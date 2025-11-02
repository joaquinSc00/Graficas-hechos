#target "indesign"

(function () {
  if (app.documents.length === 0) {
    alert("Abrí el documento final antes de correr el script.");
    return;
  }

  var doc = app.activeDocument;
  var slots = [];

  function isSlot(rectangle) {
    if (!rectangle.isValid) {
      return false;
    }

    var label = (rectangle.label || "").toLowerCase();
    if (label === "root" || label.indexOf("slot") === 0) {
      return true;
    }

    try {
      var tagged = rectangle.extractLabel("slot") || rectangle.extractLabel("root");
      if (tagged && tagged.length) {
        return true;
      }
    } catch (e) {}

    var objStyle = rectangle.appliedObjectStyle;
    if (objStyle && objStyle.isValid) {
      var styleName = objStyle.name.toLowerCase();
      if (styleName.indexOf("slot") === 0) {
        return true;
      }
    }

    return false;
  }

  function asNumber(value) {
    return Math.round(value * 1000) / 1000;
  }

  function getItemType(item) {
    try {
      if (item && item.reflect && item.reflect.name) {
        return String(item.reflect.name);
      }
    } catch (e) {}
    return "";
  }

  function allowedSlotType(typeName) {
    if (!typeName) {
      return false;
    }
    var lower = typeName.toLowerCase();
    return lower === "rectangle" || lower === "textframe" || lower === "polygon";
  }

  function safeBounds(item) {
    try {
      var gb = item.geometricBounds;
      if (gb && gb.length === 4) {
        return [gb[0], gb[1], gb[2], gb[3]];
      }
    } catch (e) {}
    return null;
  }

  function collectFromPage(page) {
    var pageSlots = [];
    var items = [];
    try {
      items = page.allPageItems ? page.allPageItems : [];
    } catch (e) {
      items = [];
    }

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (!item || !item.isValid) {
        continue;
      }

      var typeName = getItemType(item);
      if (!allowedSlotType(typeName)) {
        continue;
      }

      if (!isSlot(item)) {
        continue;
      }

      var bounds = safeBounds(item);
      if (!bounds) {
        continue;
      }

      var y1 = bounds[0];
      var x1 = bounds[1];
      var y2 = bounds[2];
      var x2 = bounds[3];

      pageSlots.push({
        page: page.name,
        id: item.id,
        label: item.label || "",
        objectStyle: item.appliedObjectStyle ? item.appliedObjectStyle.name : "",
        type: typeName,
        x_pt: asNumber(x1),
        y_pt: asNumber(y1),
        w_pt: asNumber(x2 - x1),
        h_pt: asNumber(y2 - y1)
      });
    }
    return pageSlots;
  }

  for (var p = 0; p < doc.pages.length; p++) {
    var page = doc.pages[p];
    slots = slots.concat(collectFromPage(page));
  }

  if (!slots.length) {
    alert("No se encontraron rectángulos candidatos a slots.");
    return;
  }

  var outputText = JSON.stringify(slots, null, 2);
  $.writeln(outputText);

  try {
    var defaultPath = Folder.myDocuments.fullName + "/layout_slots_detected.json";
    var file = File.saveDialog("Guardar reporte de slots", "JSON:*.json");
    if (!file) {
      file = new File(defaultPath);
    }
    file.encoding = "UTF-8";
    file.open("w");
    file.write(outputText);
    file.close();
    alert("Se detectaron " + slots.length + " slots. Archivo guardado en:\n" + file.fsName);
  } catch (err) {
    alert("Se detectaron " + slots.length + " slots, pero no se pudo guardar el archivo: " + err);
  }
})();

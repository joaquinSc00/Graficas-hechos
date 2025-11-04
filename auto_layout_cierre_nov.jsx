#target "indesign"

(function () {
  if (app.documents.length === 0) {
    alert("Abrí el INDD final antes de correr el script.");
    return;
  }
  var doc = app.activeDocument;

  function ensureJSON() {
    if (!$.global.JSON) {
      $.global.JSON = {};
    }
    if (typeof $.global.JSON.parse !== "function") {
      $.global.JSON.parse = function (text) {
        var s = String(text);
        return new Function("return (" + s + ");")();
      };
    }
  }
  ensureJSON();

  function mm2pt(mm) {
    return mm * 2.834645669291339;
  }

  function zero(n, w) {
    n = String(n);
    while (n.length < w) {
      n = "0" + n;
    }
    return n;
  }

  function readFile(f) {
    if (!f.exists) {
      return null;
    }
    f.encoding = "UTF-8";
    f.open("r");
    var s = f.read();
    f.close();
    return s;
  }

  function pickPlanFile() {
    var f = File.openDialog("Elegí plan_bloques.json", "*.json");
    if (!f) {
      throw "Sin JSON.";
    }
    var raw = readFile(f);
    if (!raw) {
      throw "No se pudo leer el JSON.";
    }
    var j = JSON.parse(raw);
    if (!(j instanceof Array)) {
      throw "El JSON esperado es un ARRAY plano de bloques.";
    }
    return j;
  }

  function getOutTxtFolder() {
    var folderPath = "C:/Users/joaqu/OneDrive/Escritorio/joaquin/HECHOS/Cier nov/out_txt";
    var f = new Folder(folderPath);
    if (!f.exists) {
      throw "No se encontró la carpeta out_txt esperada:\n" + folderPath;
    }
    return f;
  }

  function splitNoteId(noteId) {
    var p = String(noteId || "").split("#");
    var page = parseInt(p[0], 10);
    var idx = parseInt(p[1], 10);
    return { page: page, idx: idx };
  }

  function pickNoteFiles(rootFolder, page, idx) {
    var pagFolder = new Folder(rootFolder.fsName + "/PAG" + String(page));
    if (!pagFolder.exists) {
      pagFolder = new Folder(rootFolder.fsName + "/PAG" + zero(page, 2));
    }
    var nn = zero(idx, 2);
    return {
      title: new File(pagFolder.fsName + "/" + nn + "_title.txt"),
      body: new File(pagFolder.fsName + "/" + nn + "_body.txt")
    };
  }

  function ensureUnits() {
    var oldH = doc.viewPreferences.horizontalMeasurementUnits;
    var oldV = doc.viewPreferences.verticalMeasurementUnits;
    doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
    doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
    return function restore() {
      doc.viewPreferences.horizontalMeasurementUnits = oldH;
      doc.viewPreferences.verticalMeasurementUnits = oldV;
    };
  }

  function placeTextFrame(pageIndex1, x_pt, y_pt, w_pt, h_pt, contents) {
    var idx0 = pageIndex1 - 1;
    if (idx0 < 0 || idx0 >= doc.pages.length) {
      return null;
    }
    var page = doc.pages[idx0];
    var tf = page.textFrames.add();
    tf.geometricBounds = [y_pt, x_pt, y_pt + h_pt, x_pt + w_pt];
    tf.contents = contents;
    tf.textFramePreferences.firstBaselineOffset = FirstBaseline.LEADING_OFFSET;
    tf.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
    tf.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;
    tf.label = "AUTO_NOTA";
    return tf;
  }

  function composeNoteText(titleTxt, bodyTxt) {
    if (!titleTxt || titleTxt.replace(/\s+/g, "") === "") {
      titleTxt = "[SIN TÍTULO]";
    }
    if (!bodyTxt || bodyTxt.replace(/\s+/g, "") === "") {
      bodyTxt = "[SIN CUERPO]";
    }
    return titleTxt.replace(/\r?\n/g, " ").trim() + "\r" + bodyTxt;
  }

  function tryStyle(tf) {
    try {
      if (doc.paragraphStyles.itemByName("Titulo Nota").isValid) {
        tf.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("Titulo Nota");
      } else {
        var r = tf.paragraphs[0].characters.everyItem();
        r.appliedFont = app.fonts.itemByName("Minion Pro\tBold") || r.appliedFont;
        tf.paragraphs[0].pointSize = 14;
      }
      if (doc.paragraphStyles.itemByName("Cuerpo Nota").isValid) {
        for (var i = 1; i < tf.paragraphs.length; i++) {
          tf.paragraphs[i].appliedParagraphStyle = doc.paragraphStyles.itemByName("Cuerpo Nota");
        }
      }
    } catch (e) {}
  }

  try {
    var restoreUnits = ensureUnits();
    var blocks = pickPlanFile();
    var outRoot = getOutTxtFolder();

    blocks.sort(function (a, b) {
      if (a.page !== b.page) return a.page - b.page;
      if (a.y_mm !== b.y_mm) return a.y_mm - b.y_mm;
      return a.x_mm - b.x_mm;
    });

    var placed = 0;
    var missing = 0;

    for (var i = 0; i < blocks.length; i++) {
      var b = blocks[i];
      if (!b || !b.page) {
        continue;
      }
      var x = mm2pt(b.x_mm || 0);
      var y = mm2pt(b.y_mm || 0);
      var w = mm2pt(b.w_mm || 0);
      var h = mm2pt(b.h_mm || 0);

      var ids = splitNoteId(b.note_id || (b.page + "#1"));
      var files = pickNoteFiles(outRoot, ids.page || b.page, ids.idx || 1);
      var tTitle = readFile(files.title);
      var tBody = readFile(files.body);
      if (tTitle === null && tBody === null) {
        missing++;
      }

      var tf = placeTextFrame(b.page, x, y, w, h, composeNoteText(tTitle || "", tBody || ""));
      if (tf) {
        tryStyle(tf);
        placed++;
      }
    }

    alert(
      "Listo.\nMarcos colocados: " +
        placed +
        (missing ? "\nNotas sin TXT detectado: " + missing : "")
    );
    restoreUnits();
  } catch (err) {
    alert("Error: " + err);
    try {
      if (restoreUnits) {
        restoreUnits();
      }
    } catch (e) {}
  }
})();


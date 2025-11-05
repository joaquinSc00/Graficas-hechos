#target "indesign"

/**
 * Auto layout desde plan_bloques.json (array plano) + out_txt/
 * Compatible con ExtendScript (CS/CC). Incluye JSON polyfill y trim helper.
 */

// ---------- Polyfills ----------
if (typeof JSON === "undefined") {
  JSON = {};
  JSON.parse = function (s) {
    return eval("(" + s + ")");
  };
  JSON.stringify = function (obj) {
    function esc(str) {
      return (
        '"' +
        String(str).replace(/["\\\n\r\t]/g, function (c) {
          return { '"': '\\"', "\\": "\\\\", "\n": "\\n", "\r": "\\r", "\t": "\\t" }[c];
        }) +
        '"'
      );
    }
    if (obj === null) return "null";
    var t = typeof obj;
    if (t === "number" || t === "boolean") return String(obj);
    if (t === "string") return esc(obj);
    if (obj instanceof Array) return "[" + obj.map(JSON.stringify).join(",") + "]";
    var parts = [];
    for (var k in obj) if (obj.hasOwnProperty(k)) parts.push(esc(k) + ":" + JSON.stringify(obj[k]));
    return "{" + parts.join(",") + "}";
  };
}
function trimStr(s) {
  return String(s).replace(/^\s+|\s+$/g, "");
}

// ---------- Utilidades ----------
var MM2PT = 2.834645669; // 1 mm = 2.8346 pt

function getRepoRoot() {
  try {
    var scriptFile = File($.fileName);
    var indesignFolder = scriptFile.parent;
    if (!indesignFolder) return null;
    var scriptsFolder = indesignFolder.parent;
    if (!scriptsFolder) return null;
    return scriptsFolder.parent || null;
  } catch (e) {
    return null;
  }
}

function getDefaultDataPaths() {
  var repoRoot = getRepoRoot();
  if (!repoRoot) {
    return { plan: null, outTxt: null };
  }
  var dataRoot = new Folder(repoRoot.fsName + "/data");
  return {
    plan: new File(dataRoot.fsName + "/reports/plan_bloques.json"),
    outTxt: new Folder(dataRoot.fsName + "/out_txt"),
  };
}

function askFile(promptTxt, filter, defaultPath) {
  if (defaultPath) {
    var defFile = defaultPath instanceof File ? defaultPath : File(defaultPath);
    if (defFile.exists) {
      $.writeln("Usando plan por defecto: " + defFile.fsName);
      return defFile;
    }
  }
  var f = File.openDialog(promptTxt, filter || "*.*");
  if (!f) throw new Error("Operación cancelada.");
  return f;
}
function askFolder(promptTxt, defaultPath) {
  if (defaultPath) {
    var defFolder = defaultPath instanceof Folder ? defaultPath : Folder(defaultPath);
    if (defFolder.exists) {
      $.writeln("Usando TXT por defecto: " + defFolder.fsName);
      return defFolder;
    }
  }
  var f = Folder.selectDialog(promptTxt);
  if (!f) throw new Error("Operación cancelada.");
  return f;
}
function readTextFile(f) {
  if (!f || !f.exists) return "";
  f.encoding = "UTF-8";
  f.open("r");
  var txt = f.read();
  f.close();
  return txt;
}
function ensureLayer(doc, name) {
  var lyr;
  try {
    lyr = doc.layers.getByName(name);
  } catch (e) {
    lyr = null;
  }
  if (!lyr) lyr = doc.layers.add({ name: name });
  return lyr;
}
function getPageByNumber(doc, num) {
  for (var i = 0; i < doc.pages.length; i++) {
    var p = doc.pages[i];
    if (parseInt(p.name, 10) === parseInt(num, 10)) return p;
  }
  return null;
}
function pad2(n) {
  n = parseInt(n, 10);
  return n < 10 ? "0" + n : "" + n;
}

// Devuelve {title, body} leyendo out_txt/<page>/<NN>_title.txt y <NN>_body.txt
function readNoteTexts(outRoot, pageNum, noteId) {
  // noteId viene como "7#1" → NN = 01
  var idx = String(noteId).split("#")[1] || "1";
  var nn = pad2(idx);
  var pFolder = new Folder(outRoot.fsName + "/" + pageNum);
  var title = "", body = "";

  if (pFolder.exists) {
    var fTitle = new File(pFolder.fsName + "/" + nn + "_title.txt");
    var fBody = new File(pFolder.fsName + "/" + nn + "_body.txt");
    title = trimStr(readTextFile(fTitle));
    body = trimStr(readTextFile(fBody));
    // Fallbacks si no existieran con dos dígitos
    if (!title && !body) {
      fTitle = new File(pFolder.fsName + "/" + idx + "_title.txt");
      fBody = new File(pFolder.fsName + "/" + idx + "_body.txt");
      title = trimStr(readTextFile(fTitle));
      body = trimStr(readTextFile(fBody));
    }
    // Ultimo fallback: meta.json con fields "title"/"body" si existiera
    if (!title || !body) {
      var fMeta = new File(pFolder.fsName + "/meta.json");
      if (fMeta.exists) {
        try {
          var meta = JSON.parse(readTextFile(fMeta));
          if (!title && meta && meta.title) title = trimStr(meta.title);
          if (!body && meta && meta.body) body = trimStr(meta.body);
        } catch (_) {}
      }
    }
  }
  return { title: title || "", body: body || "" };
}

// Crea marcos de título / cuerpo y opcionalmente placeholder de imagen
function placeBlock(doc, layer, page, blk, noteTxts) {
  var x = blk.x_mm * MM2PT,
    y = blk.y_mm * MM2PT;
  var w = blk.w_mm * MM2PT,
    h = blk.h_mm * MM2PT;

  // Alturas en pt desde mm
  var tH = (blk.title_height_mm || 0) * MM2PT;
  var iH = (blk.image_height_mm || 0) * MM2PT;
  var bH = h - tH - iH;
  if (bH < 0) bH = 0;

  // Grupo contenedor
  var groupItems = [];

  // Título
  if (tH > 0) {
    var tfT = page.textFrames.add(layer, {
      geometricBounds: [y, x, y + tH, x + w],
    });
    tfT.contents = noteTxts.title || "";
    groupItems.push(tfT);
  }

  // Imagen (placeholder) – si corresponde
  if (iH > 0) {
    var imgY = y + tH; // debajo del título
    var rImg = page.rectangles.add(layer, {
      geometricBounds: [imgY, x, imgY + iH, x + w],
      strokeWeight: 0.5,
    });
    rImg.fillColor = doc.swatches.itemByName("None");
    rImg.strokeColor = doc.swatches.itemByName("Black");
    // Etiqueta con modo (horizontal/vertical/none)
    try {
      rImg.label = "IMG:" + (blk.img_mode || "auto");
    } catch (e) {}
    groupItems.push(rImg);
  }

  // Cuerpo (ajustado a body_chars_fit)
  var bodyTop = y + tH + iH;
  if (bH > 0) {
    var tfB = page.textFrames.add(layer, {
      geometricBounds: [bodyTop, x, bodyTop + bH, x + w],
    });
    var bodyText = noteTxts.body || "";
    var maxChars = parseInt(blk.body_chars_fit || 0, 10);
    if (maxChars > 0 && bodyText.length > maxChars) {
      bodyText = bodyText.substr(0, maxChars);
    }
    tfB.contents = bodyText;
    groupItems.push(tfB);
  }

  // Etiqueta de depuración (note_id, span, etc.)
  var lbl = page.textFrames.add(layer, {
    geometricBounds: [y - 4 * MM2PT, x, y, x + 30 * MM2PT],
  });
  lbl.contents = String(blk.note_id || "") + "  [" + (blk.span || 1) + " col]";
  lbl.textFramePreferences.firstBaselineOffset = FirstBaseline.capHeight;
  lbl.fillColor = doc.swatches.itemByName("None");
  lbl.strokeColor = doc.swatches.itemByName("Black");
  lbl.strokeWeight = 0.25;
  groupItems.push(lbl);

  // Agrupar todo (opcional)
  try {
    page.groups.add(groupItems);
  } catch (e) {}
}

// ---------- Main ----------
(function main() {
  if (!app.documents.length) {
    alert("Abrí primero el documento de InDesign (la maqueta).");
    return;
  }
  var doc = app.activeDocument;

  var defaults = getDefaultDataPaths();

  // Seleccionar plan_bloques.json
  var planFile = askFile(
    "Seleccioná plan_bloques.json (array plano de bloques)",
    "*.json",
    defaults.plan
  );
  var planRaw = readTextFile(planFile);
  var plan;
  try {
    plan = JSON.parse(planRaw);
  } catch (e) {
    alert("No pude parsear el JSON:\n" + e);
    return;
  }
  if (!plan || !(plan instanceof Array)) {
    alert("Se esperaba un ARRAY de bloques. Revisá plan_bloques.json");
    return;
  }

  // Seleccionar carpeta out_txt
  var outRoot = askFolder(
    "Seleccioná la carpeta raíz de out_txt (contiene subcarpetas 2,3,5,...)",
    defaults.outTxt
  );

  // Agrupar por página
  var byPage = {};
  for (var i = 0; i < plan.length; i++) {
    var blk = plan[i];
    var pnum = parseInt(blk.page, 10);
    if (!byPage[pnum]) byPage[pnum] = [];
    byPage[pnum].push(blk);
  }

  var layer = ensureLayer(doc, "auto_layout");

  // Iterar páginas
  for (var p in byPage) {
    if (!byPage.hasOwnProperty(p)) continue;
    var pageNum = parseInt(p, 10);
    var page = getPageByNumber(doc, pageNum);
    if (!page) {
      $.writeln("WARNING: No encuentro la página " + pageNum + " en el doc.");
      continue;
    }
    var blocks = byPage[p];

    for (var j = 0; j < blocks.length; j++) {
      var blk = blocks[j];
      var txts = readNoteTexts(outRoot, pageNum, blk.note_id || "");
      try {
        placeBlock(doc, layer, page, blk, txts);
      } catch (e) {
        $.writeln("ERROR ubicando bloque " + (blk.note_id || "?") + " pág " + pageNum + ": " + e);
      }
    }
  }

  alert("Listo. Se volcaron " + plan.length + " bloques sobre la maqueta.");
})();

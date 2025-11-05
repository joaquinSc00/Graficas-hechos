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

function parseRectLike(rectLike) {
  if (!rectLike) return null;

  var raw = rectLike;
  if (typeof raw === "string") {
    try {
      raw = JSON.parse(raw);
    } catch (e) {
      var parts = String(raw)
        .replace(/[\[\]()]/g, "")
        .split(/[,\s]+/);
      if (parts.length >= 4) {
        raw = [parts[0], parts[1], parts[2], parts[3]];
      }
    }
  }

  if (!raw) return null;

  if (raw.rect) {
    return parseRectLike(raw.rect);
  }

  if (raw instanceof Array) {
    if (raw.length < 4) return null;
    return {
      x_mm: parseFloat(raw[0]) || 0,
      y_mm: parseFloat(raw[1]) || 0,
      w_mm: parseFloat(raw[2]) || 0,
      h_mm: parseFloat(raw[3]) || 0,
    };
  }

  function _num(value) {
    var n = parseFloat(value);
    return isNaN(n) ? 0 : n;
  }

  var x = raw.hasOwnProperty("x_mm") ? raw.x_mm : raw.x;
  var y = raw.hasOwnProperty("y_mm") ? raw.y_mm : raw.y;
  var w = raw.hasOwnProperty("w_mm") ? raw.w_mm : raw.w;
  if (raw.hasOwnProperty("width_mm")) w = raw.width_mm;
  var h = raw.hasOwnProperty("h_mm") ? raw.h_mm : raw.h;
  if (raw.hasOwnProperty("height_mm")) h = raw.height_mm;

  return {
    x_mm: _num(x),
    y_mm: _num(y),
    w_mm: _num(w),
    h_mm: _num(h),
  };
}

function rectToBounds(rect) {
  if (!rect) return [0, 0, 0, 0];
  var x = (rect.x_mm || 0) * MM2PT;
  var y = (rect.y_mm || 0) * MM2PT;
  var w = (rect.w_mm || 0) * MM2PT;
  var h = (rect.h_mm || 0) * MM2PT;
  return [y, x, y + h, x + w];
}

function extractPlanNotes(plan) {
  if (!plan) return [];
  if (plan.notes && plan.notes instanceof Array) return plan.notes;
  if (plan.blocks && plan.blocks instanceof Array) return plan.blocks;
  if (plan instanceof Array) return plan;
  return [];
}

function normalizeNoteEntry(entry) {
  if (!entry) return null;

  var pageNum = parseInt(entry.page, 10);
  if (isNaN(pageNum)) return null;

  var noteId = entry.note_id || entry.id || "";
  var columnStart = parseInt(
    entry.column_index || (entry.columns && entry.columns.start) || 1,
    10
  );
  if (isNaN(columnStart)) columnStart = 1;
  var columnSpan = parseInt(entry.span || (entry.columns && entry.columns.span) || 1, 10);
  if (isNaN(columnSpan) || columnSpan <= 0) columnSpan = 1;

  var frameRect = parseRectLike(entry.frame || entry);
  var titleRect = null;
  if (entry.title) {
    titleRect = parseRectLike(entry.title.rect ? entry.title.rect : entry.title);
  }
  var imageRect = null;
  if (entry.image) {
    imageRect = parseRectLike(entry.image.rect ? entry.image.rect : entry.image);
  } else if (entry.image_rect_mm) {
    imageRect = parseRectLike(entry.image_rect_mm);
  }

  var bodySegments = [];
  if (entry.body && entry.body instanceof Array) {
    for (var i = 0; i < entry.body.length; i++) {
      var seg = entry.body[i] || {};
      var segRect = parseRectLike(seg.rect ? seg.rect : seg);
      if (!segRect) continue;
      var segColumn = parseInt(
        seg.column != null ? seg.column : columnStart + i,
        10
      );
      if (isNaN(segColumn)) segColumn = columnStart + i;
      var segRel = parseInt(
        seg.relative_column != null ? seg.relative_column : i,
        10
      );
      if (isNaN(segRel)) segRel = i;
      bodySegments.push({
        column: segColumn,
        relative_column: segRel,
        rect: segRect,
      });
    }
  }

  function _fallbackRect() {
    if (frameRect) return frameRect;
    if (bodySegments.length > 0) return bodySegments[0].rect;
    if (titleRect) return titleRect;
    if (imageRect) return imageRect;
    return { x_mm: 0, y_mm: 0, w_mm: 0, h_mm: 0 };
  }

  if (!frameRect) {
    var base = _fallbackRect();
    frameRect = {
      x_mm: base.x_mm,
      y_mm: base.y_mm,
      w_mm: base.w_mm,
      h_mm: base.h_mm,
    };
  }

  var metrics = entry.metrics || {};
  if (!metrics.body_chars_fit && entry.body_chars_fit)
    metrics.body_chars_fit = entry.body_chars_fit;
  if (!metrics.body_chars_overflow && entry.body_chars_overflow)
    metrics.body_chars_overflow = entry.body_chars_overflow;
  if (!metrics.title_height_mm && entry.title_height_mm)
    metrics.title_height_mm = entry.title_height_mm;
  if (!metrics.body_height_mm && entry.body_height_mm)
    metrics.body_height_mm = entry.body_height_mm;
  if (!metrics.image_height_mm && entry.image_height_mm)
    metrics.image_height_mm = entry.image_height_mm;
  if (!metrics.title_lines && entry.title_lines)
    metrics.title_lines = entry.title_lines;
  if (!metrics.body_lines && entry.body_lines)
    metrics.body_lines = entry.body_lines;

  if (bodySegments.length === 0) {
    var titleH = titleRect ? titleRect.h_mm || 0 : metrics.title_height_mm || 0;
    var imageH = imageRect ? imageRect.h_mm || 0 : metrics.image_height_mm || 0;
    var fallbackRect = {
      x_mm: frameRect.x_mm,
      y_mm: frameRect.y_mm + titleH + imageH,
      w_mm: frameRect.w_mm,
      h_mm: Math.max(frameRect.h_mm - titleH - imageH, 0),
    };
    bodySegments.push({
      column: columnStart,
      relative_column: 0,
      rect: fallbackRect,
    });
  }

  var hasBodyHeight = false;
  for (var bi = 0; bi < bodySegments.length; bi++) {
    if (bodySegments[bi].rect && bodySegments[bi].rect.h_mm > 0) {
      hasBodyHeight = true;
      break;
    }
  }
  if (!hasBodyHeight) {
    bodySegments = [
      {
        column: columnStart,
        relative_column: 0,
        rect: {
          x_mm: frameRect.x_mm,
          y_mm: frameRect.y_mm,
          w_mm: frameRect.w_mm,
          h_mm: frameRect.h_mm,
        },
      },
    ];
  }

  var imageMode =
    (entry.image && entry.image.mode) || entry.img_mode || (metrics.img_mode || "none");
  metrics.img_mode = imageMode;
  var imageSpan =
    (entry.image && entry.image.span) || entry.image_span || metrics.image_span || 0;
  metrics.image_span = imageSpan;

  return {
    page: pageNum,
    note_id: noteId,
    columns: { start: columnStart, span: columnSpan },
    frame: frameRect,
    title: titleRect ? { rect: titleRect, lines: metrics.title_lines || 0 } : null,
    body: bodySegments,
    image: imageRect
      ? { rect: imageRect, mode: imageMode, span: imageSpan, height_mm: metrics.image_height_mm || 0 }
      : null,
    metrics: metrics,
  };
}

// Crea marcos de título / cuerpo y opcionalmente placeholder de imagen
function placeNote(doc, layer, page, note, noteTxts) {
  var frameRect = note.frame || { x_mm: 0, y_mm: 0, w_mm: 0, h_mm: 0 };
  var frameBounds = rectToBounds(frameRect);
  var x = frameBounds[1];
  var y = frameBounds[0];
  var w = frameBounds[3] - frameBounds[1];
  var h = frameBounds[2] - frameBounds[0];

  var groupItems = [];

  // Título
  var titleRect = note.title ? note.title.rect : null;
  if (titleRect && (titleRect.h_mm || 0) > 0) {
    var titleBounds = rectToBounds(titleRect);
    var tfT = page.textFrames.add(layer, { geometricBounds: titleBounds });
    tfT.contents = noteTxts.title || "";
    groupItems.push(tfT);
  }

  // Imagen (placeholder) – si corresponde
  if (note.image && note.image.rect && (note.image.rect.h_mm || 0) > 0) {
    var imgBounds = rectToBounds(note.image.rect);
    var rImg = page.rectangles.add(layer, {
      geometricBounds: imgBounds,
      strokeWeight: 0.5,
    });
    rImg.fillColor = doc.swatches.itemByName("None");
    rImg.strokeColor = doc.swatches.itemByName("Black");
    try {
      rImg.label = "IMG:" + (note.image.mode || "auto");
    } catch (e) {}
    groupItems.push(rImg);
  }

  // Cuerpo (una entrada por columna)
  var bodySegments = note.body || [];
  var bodyFrames = [];
  var previousBodyFrame = null;
  for (var i = 0; i < bodySegments.length; i++) {
    var seg = bodySegments[i];
    if (!seg || !seg.rect) continue;
    if (seg.rect.h_mm <= 0 || seg.rect.w_mm <= 0) continue;
    var segBounds = rectToBounds(seg.rect);
    var tfB = page.textFrames.add(layer, { geometricBounds: segBounds });
    if (previousBodyFrame) {
      try {
        previousBodyFrame.nextTextFrame = tfB;
      } catch (e) {}
    }
    previousBodyFrame = tfB;
    bodyFrames.push(tfB);
    groupItems.push(tfB);
  }

  var bodyText = noteTxts.body || "";
  var maxChars = parseInt((note.metrics && note.metrics.body_chars_fit) || 0, 10);
  if (maxChars > 0 && bodyText.length > maxChars) {
    bodyText = bodyText.substr(0, maxChars);
  }
  if (bodyFrames.length > 0) {
    bodyFrames[0].contents = bodyText;
  }

  // Etiqueta de depuración (note_id, span, etc.)
  var span = (note.columns && note.columns.span) || 1;
  var lbl = page.textFrames.add(layer, {
    geometricBounds: [y - 4 * MM2PT, x, y, x + 30 * MM2PT],
  });
  lbl.contents = String(note.note_id || "") + "  [" + span + " col]";
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
    "Seleccioná plan_bloques.json (estructura por nota)",
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

  var rawNotes = extractPlanNotes(plan);
  if (!rawNotes || !rawNotes.length) {
    alert("plan_bloques.json no contiene notas serializadas.");
    return;
  }

  var notes = [];
  for (var ni = 0; ni < rawNotes.length; ni++) {
    var normalized = normalizeNoteEntry(rawNotes[ni]);
    if (normalized) notes.push(normalized);
  }
  if (!notes.length) {
    alert("Ninguna nota pudo normalizarse desde plan_bloques.json.");
    return;
  }

  // Seleccionar carpeta out_txt
  var outRoot = askFolder(
    "Seleccioná la carpeta raíz de out_txt (contiene subcarpetas 2,3,5,...)",
    defaults.outTxt
  );

  // Agrupar por página
  var byPage = {};
  for (var i = 0; i < notes.length; i++) {
    var note = notes[i];
    var pnum = parseInt(note.page, 10);
    if (!byPage[pnum]) byPage[pnum] = [];
    byPage[pnum].push(note);
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
    var pageNotes = byPage[p];

    for (var j = 0; j < pageNotes.length; j++) {
      var note = pageNotes[j];
      var txts = readNoteTexts(outRoot, pageNum, note.note_id || "");
      try {
        placeNote(doc, layer, page, note, txts);
      } catch (e) {
        $.writeln(
          "ERROR ubicando nota " + (note.note_id || "?") + " pág " + pageNum + ": " + e
        );
      }
    }
  }

  alert("Listo. Se volcaron " + notes.length + " notas sobre la maqueta.");
})();

#target "indesign"

(function () {
  if (app.documents.length === 0) {
    alert("Abrí el documento final antes de correr el script.");
    return;
  }

  var DEFAULT_CONFIG = {
    styles: {
      body: { name: "Textos", pt_base: 9.5, pt_min: 9.0, pt_max: 10.0 },
      title: { name: "Titulo 1", pt_base: 25, pt_min: 24, pt_max: 26 }
    },
    overset: {
      max_expand_mm: 12,
      body_step_pt: 0.25,
      body_max_drop_pt: 0.5,
      title_max_drop_pt: 0.5
    },
    images: {
      default_object_style: "",
      fallback_object_style: ""
    },
    "export": {
      pdf_per_page: false,
      pdf_preset_name: "Smallest File Size"
    },
    importPrefs: {
      split_titles_by_bold: true
    }
  };

  var CSV_COLUMNS = [
    "pagina",
    "nota",
    "chars",
    "body_pt",
    "title_pt",
    "overset",
    "excedente_aprox",
    "warnings",
    "errors"
  ];

  function clone(obj) {
    if (obj === null || typeof obj !== "object") {
      return obj;
    }
    var out = obj instanceof Array ? [] : {};
    for (var key in obj) {
      if (obj.hasOwnProperty(key)) {
        out[key] = clone(obj[key]);
      }
    }
    return out;
  }

  function deepMerge(target, source) {
    for (var key in source) {
      if (!source.hasOwnProperty(key)) {
        continue;
      }
      var value = source[key];
      if (value && typeof value === "object" && !(value instanceof File) && !(value instanceof Folder)) {
        if (!target[key]) {
          target[key] = value instanceof Array ? [] : {};
        }
        deepMerge(target[key], value);
      } else {
        target[key] = value;
      }
    }
    return target;
  }

  function readConfig(file) {
    var cfg = clone(DEFAULT_CONFIG);
    if (!file || !file.exists) {
      return cfg;
    }
    try {
      file.open("r");
      var data = file.read();
      file.close();
      if (data && data.length) {
        var parsed = JSON.parse(data);
        cfg = deepMerge(cfg, parsed);
      }
    } catch (e) {
      try { file.close(); } catch (_) {}
      $.writeln("Config error: " + e);
    }
    return cfg;
  }

  function selectRootFolder() {
    return Folder.selectDialog("Seleccionar carpeta raíz de cierre");
  }

  function listPageFolders(root) {
    if (!root || !root.exists) {
      return [];
    }
    var items = root.getFiles();
    var folders = [];
    for (var i = 0; i < items.length; i++) {
      if (items[i] instanceof Folder && /\d+/.test(items[i].name)) {
        folders.push(items[i]);
      }
    }
    folders.sort(function (a, b) {
      return a.name > b.name ? 1 : (a.name < b.name ? -1 : 0);
    });
    return folders;
  }

  function extractNumbers(name) {
    if (!name) {
      return [];
    }
    var matches = name.match(/\d+/g);
    if (!matches) {
      return [];
    }
    var out = [];
    for (var i = 0; i < matches.length; i++) {
      var n = parseInt(matches[i], 10);
      if (!isNaN(n)) {
        out.push(n);
      }
    }
    return out;
  }

  function resolveDocumentPages(doc, numbers, warnings) {
    var pages = [];
    if (!numbers || numbers.length === 0) {
      return pages;
    }
    for (var i = 0; i < numbers.length; i++) {
      var name = String(numbers[i]);
      var page = null;
      try {
        page = doc.pages.itemByName(name);
        if (!page || !page.isValid) {
          throw new Error("missing");
        }
        pages.push(page);
      } catch (e) {
        if (warnings) {
          warnings.push("page_missing:" + name);
        }
      }
    }
    return pages;
  }

  function createScratchFrame(doc) {
    var page = doc.pages[0];
    var tf = page.textFrames.add();
    tf.geometricBounds = [10000, 10000, 10100, 10100];
    tf.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
    return tf;
  }

  function importDocxToScratchStory(doc, file, importPrefs) {
    var tf = createScratchFrame(doc);
    var prefs = app.wordImportPreferences;
    var previous = {};
    try {
      for (var key in importPrefs) {
        if (importPrefs.hasOwnProperty(key)) {
          previous[key] = prefs[key];
          prefs[key] = importPrefs[key];
        }
      }
    } catch (e) {
      $.writeln("word prefs: " + e);
    }
    try {
      tf.place(file);
      tf.parentStory.recompose();
      return { frame: tf, story: tf.parentStory };
    } catch (err) {
      $.writeln("place docx: " + err);
      try { tf.remove(); } catch (_remove) {}
      return null;
    } finally {
      try {
        for (var key2 in previous) {
          prefs[key2] = previous[key2];
        }
      } catch (e2) {}
    }
  }

  function paragraphIsBold(paragraph) {
    if (!paragraph || paragraph.contents.length === 0) {
      return false;
    }
    var chars = paragraph.characters;
    if (!chars || chars.length === 0) {
      return false;
    }
    var boldCount = 0;
    for (var i = 0; i < chars.length; i++) {
      var ch = chars[i];
      var isBold = false;
      try {
        if (ch.characterAttributes && ch.characterAttributes.bold !== undefined) {
          isBold = ch.characterAttributes.bold;
        }
      } catch (_) {}
      if (!isBold) {
        try {
          var style = String(ch.fontStyle || "").toLowerCase();
          if (style.indexOf("bold") >= 0) {
            isBold = true;
          }
        } catch (err) {}
      }
      if (isBold) {
        boldCount++;
      }
    }
    return boldCount / chars.length >= 0.8;
  }

  function cleanParagraphText(text) {
    if (!text) {
      return "";
    }
    return String(text).replace(/\r+$/, "").replace(/^\s+|\s+$/g, "");
  }

  function splitStoryByBoldTitles(story) {
    var segments = [];
    if (!story || !story.isValid) {
      return segments;
    }
    var paras = story.paragraphs;
    var current = null;
    for (var i = 0; i < paras.length; i++) {
      var paragraph = paras[i];
      var text = cleanParagraphText(paragraph.contents);
      if (!text) {
        continue;
      }
      if (paragraphIsBold(paragraph)) {
        if (current && (current.body.length || current.titleText)) {
          current.bodyText = current.body.join("\r");
          segments.push(current);
        }
        current = { titleText: text, body: [] };
      } else {
        if (!current) {
          continue;
        }
        current.body.push(text);
      }
    }
    if (current && (current.body.length || current.titleText)) {
      current.bodyText = current.body.join("\r");
      segments.push(current);
    }
    return segments;
  }

  function disposeScratch(result) {
    if (!result) {
      return;
    }
    try {
      if (result.frame && result.frame.isValid) {
        result.frame.remove();
      }
    } catch (_) {}
  }

  function findLabeledItem(pages, label) {
    if (!label) {
      return null;
    }
    for (var i = 0; i < pages.length; i++) {
      var page = pages[i];
      if (!page || !page.isValid) {
        continue;
      }
      var items = page.allPageItems;
      for (var j = 0; j < items.length; j++) {
        if (items[j].label === label) {
          return items[j];
        }
      }
    }
    return null;
  }

  function clearStory(textFrame) {
    if (!textFrame || !textFrame.isValid) {
      return;
    }
    try {
      var story = textFrame.parentStory;
      story.contents = "";
      story.recompose();
    } catch (e) {
      $.writeln("clearStory: " + e);
    }
  }

  function writeStory(textFrame, text) {
    if (!textFrame || !textFrame.isValid) {
      return null;
    }
    var story = textFrame.parentStory;
    clearStory(textFrame);
    if (!text) {
      story.recompose();
      return story;
    }
    story.insertionPoints[0].contents = text;
    story.recompose();
    return story;
  }

  function applyStyleAndSize(story, styleName, targetSize, minSize, maxSize) {
    if (!story || !story.isValid || !styleName) {
      return;
    }
    var doc = app.activeDocument;
    var style = null;
    try {
      style = doc.paragraphStyles.itemByName(styleName);
      if (!style || !style.isValid) {
        return;
      }
    } catch (_) {
      return;
    }
    var paragraphs = story.paragraphs;
    for (var i = 0; i < paragraphs.length; i++) {
      var para = paragraphs[i];
      para.appliedParagraphStyle = style;
      if (targetSize !== undefined && targetSize !== null) {
        var size = targetSize;
        if (minSize !== undefined && size < minSize) {
          size = minSize;
        }
        if (maxSize !== undefined && size > maxSize) {
          size = maxSize;
        }
        para.pointSize = size;
      }
    }
    story.recompose();
  }

  function applyStyleToFirstParagraph(story, styleName, size, min, max) {
    if (!story || !story.isValid || !styleName) {
      return;
    }
    var doc = app.activeDocument;
    var style = null;
    try {
      style = doc.paragraphStyles.itemByName(styleName);
      if (!style || !style.isValid) {
        return;
      }
    } catch (_) {
      return;
    }
    if (story.paragraphs.length > 0) {
      var para = story.paragraphs[0];
      para.appliedParagraphStyle = style;
      if (size !== undefined && size !== null) {
        var pt = size;
        if (min !== undefined && pt < min) {
          pt = min;
        }
        if (max !== undefined && pt > max) {
          pt = max;
        }
        para.pointSize = pt;
      }
    }
    story.recompose();
  }

  function expandFrameHeightMM(textFrame, delta) {
    if (!textFrame || !textFrame.isValid || !delta) {
      return;
    }
    var pts = delta * 2.834645669;
    try {
      var gb = textFrame.geometricBounds;
      gb[2] = gb[2] + pts;
      textFrame.geometricBounds = gb;
      textFrame.parentStory.recompose();
    } catch (e) {
      $.writeln("expandFrameHeightMM: " + e);
    }
  }

  function nudgeStyle(textFrame, styleName, delta, min, max) {
    if (!textFrame || !textFrame.isValid || !delta) {
      return null;
    }
    var story = textFrame.parentStory;
    if (!story || !story.isValid) {
      return null;
    }
    var paragraphs = story.paragraphs;
    var applied = null;
    for (var i = 0; i < paragraphs.length; i++) {
      var para = paragraphs[i];
      if (!para.appliedParagraphStyle || para.appliedParagraphStyle.name !== styleName) {
        continue;
      }
      var current = para.pointSize;
      var next = current + delta;
      if (min !== undefined && next < min) {
        next = min;
      }
      if (max !== undefined && next > max) {
        next = max;
      }
      if (next !== current) {
        para.pointSize = next;
        applied = next;
      }
    }
    if (applied !== null) {
      story.recompose();
    }
    return applied;
  }

  function storyOversetInfo(story) {
    var info = { over: false, exceed: 0 };
    if (!story || !story.isValid) {
      return info;
    }
    info.over = story.overflows;
    if (info.over) {
      try {
        var containers = story.textContainers;
        var visible = 0;
        for (var i = 0; i < containers.length; i++) {
          visible += containers[i].characters.length;
        }
        info.exceed = Math.max(0, story.characters.length - visible);
      } catch (_) {}
    }
    return info;
  }

  function currentStylePointSize(textFrame, styleName) {
    if (!textFrame || !textFrame.isValid || !styleName) {
      return "";
    }
    var story = textFrame.parentStory;
    if (!story || !story.isValid) {
      return "";
    }
    var paragraphs = story.paragraphs;
    for (var i = 0; i < paragraphs.length; i++) {
      var para = paragraphs[i];
      if (para.appliedParagraphStyle && para.appliedParagraphStyle.name === styleName) {
        return para.pointSize;
      }
    }
    return "";
  }

  function resolveOverset(bodyFrame, titleFrame, cfg, stylesCfg) {
    var result = {
      overset: "",
      exceed: 0,
      bodyPointSize: currentStylePointSize(bodyFrame, stylesCfg.body.name),
      titlePointSize: titleFrame ? currentStylePointSize(titleFrame, stylesCfg.title.name) : ""
    };

    if (!bodyFrame || !bodyFrame.isValid) {
      return result;
    }

    var story = bodyFrame.parentStory;
    var info = storyOversetInfo(story);
    if (!info.over) {
      result.overset = "";
      result.exceed = info.exceed;
      result.bodyPointSize = currentStylePointSize(bodyFrame, stylesCfg.body.name);
      result.titlePointSize = titleFrame ? currentStylePointSize(titleFrame, stylesCfg.title.name) : "";
      return result;
    }

    var expanded = 0;
    while (info.over && expanded < cfg.max_expand_mm) {
      var step = Math.min(2, cfg.max_expand_mm - expanded);
      expandFrameHeightMM(bodyFrame, step);
      expanded += step;
      info = storyOversetInfo(story);
    }

    if (info.over) {
      var dropped = 0;
      while (info.over && dropped < cfg.body_max_drop_pt) {
        var delta = -Math.min(cfg.body_step_pt, cfg.body_max_drop_pt - dropped);
        var applied = nudgeStyle(bodyFrame, stylesCfg.body.name, delta, stylesCfg.body.pt_min, stylesCfg.body.pt_max);
        if (applied === null) {
          break;
        }
        dropped += Math.abs(delta);
        info = storyOversetInfo(story);
      }
    }

    if (info.over && titleFrame && titleFrame.isValid) {
      var titleStory = titleFrame.parentStory;
      var titleDropped = 0;
      while (info.over && titleDropped < cfg.title_max_drop_pt) {
        var tDelta = -Math.min(cfg.body_step_pt, cfg.title_max_drop_pt - titleDropped);
        var appliedTitle = nudgeStyle(titleFrame, stylesCfg.title.name, tDelta, stylesCfg.title.pt_min, stylesCfg.title.pt_max);
        if (appliedTitle === null) {
          break;
        }
        titleDropped += Math.abs(tDelta);
        titleStory.recompose();
        info = storyOversetInfo(story);
      }
    }

    info = storyOversetInfo(story);
    result.exceed = info.exceed;
    result.bodyPointSize = currentStylePointSize(bodyFrame, stylesCfg.body.name);
    result.titlePointSize = titleFrame ? currentStylePointSize(titleFrame, stylesCfg.title.name) : result.titlePointSize;
    if (info.over) {
      result.overset = "overset_hard";
    }
    return result;
  }

  function ensureFolder(folder) {
    if (!folder.exists) {
      folder.create();
    }
    return folder;
  }

  function CsvLogger(file, columns) {
    this.columns = columns.slice(0);
    this.file = file;
    if (file.exists) {
      file.remove();
    }
    file.open("w");
    this.writeArray(columns);
  }

  CsvLogger.prototype.writeRow = function (data) {
    var row = [];
    for (var i = 0; i < this.columns.length; i++) {
      var key = this.columns[i];
      var value = data && data.hasOwnProperty(key) ? data[key] : "";
      if (value === null || value === undefined) {
        value = "";
      }
      row.push(String(value));
    }
    this.writeArray(row);
  };

  CsvLogger.prototype.writeArray = function (values) {
    var parts = [];
    for (var i = 0; i < values.length; i++) {
      var text = String(values[i]);
      if (text.indexOf("\"") >= 0 || text.indexOf(",") >= 0) {
        text = '"' + text.replace(/"/g, '""') + '"';
      }
      parts.push(text);
    }
    this.file.write(parts.join(",") + "\n");
  };

  CsvLogger.prototype.close = function () {
    if (this.file && this.file.opened) {
      this.file.close();
    }
  };

  function collectImages(pageFolder) {
    var files = pageFolder.getFiles();
    var map = {};
    for (var i = 0; i < files.length; i++) {
      var file = files[i];
      if (!(file instanceof File)) {
        continue;
      }
      var match = /nota(\d+)_foto(\d+)/i.exec(file.name);
      if (!match) {
        continue;
      }
      var noteIndex = parseInt(match[1], 10);
      var photoIndex = parseInt(match[2], 10);
      if (isNaN(noteIndex) || isNaN(photoIndex)) {
        continue;
      }
      if (!map[noteIndex]) {
        map[noteIndex] = [];
      }
      map[noteIndex].push({ index: photoIndex, file: file });
    }
    for (var key in map) {
      if (!map.hasOwnProperty(key)) continue;
      map[key].sort(function (a, b) {
        return a.index - b.index;
      });
    }
    return map;
  }

  function exportPagePDF(doc, page, outFile, presetName) {
    if (!doc || !page || !page.isValid || !outFile) {
      return "missing_info";
    }
    try {
      var preset = null;
      if (presetName) {
        try {
          preset = app.pdfExportPresets.itemByName(presetName);
          if (!preset.isValid) {
            preset = null;
          }
        } catch (_) {
          preset = null;
        }
      }
      var previous = doc.exportPreferences.pageRange;
      doc.exportPreferences.pageRange = String(page.documentOffset + 1);
      doc.exportFile(ExportFormat.PDF_TYPE, outFile, false, preset);
      doc.exportPreferences.pageRange = previous;
      return "";
    } catch (e) {
      $.writeln("export PDF error: " + e);
      return String(e);
    }
  }

  function run() {
    var doc = app.activeDocument;
    var root = selectRootFolder();
    if (!root) {
      return;
    }

    var config = readConfig(File(root.fsName + "/config.json"));
    var csvPath = File(root.fsName + "/reporte.csv");
    var csv = new CsvLogger(csvPath, CSV_COLUMNS);
    var processedPages = [];

    var folders = listPageFolders(root);
    for (var fi = 0; fi < folders.length; fi++) {
      var pageFolder = folders[fi];
      var warnings = [];
      var numbers = extractNumbers(pageFolder.name);
      var pages = resolveDocumentPages(doc, numbers, warnings);
      if (!pages.length) {
        continue;
      }
      for (var pi = 0; pi < pages.length; pi++) {
        processedPages.push(pages[pi]);
      }

      var docxFiles = pageFolder.getFiles(function (f) {
        if (!(f instanceof File)) return false;
        return /\.docx$/i.test(f.name);
      });

      if (!docxFiles || docxFiles.length === 0) {
        warnings.push("no_docx_found");
        csv.writeRow({
          pagina: numbers.join("-"),
          nota: "",
          warnings: warnings.join(";"),
          errors: ""
        });
      }

      var noteCounter = 0;
      var imagesMap = collectImages(pageFolder);

      for (var df = 0; df < docxFiles.length; df++) {
        var docxFile = docxFiles[df];
        var scratch = importDocxToScratchStory(doc, docxFile, config.importPrefs || {});
        if (!scratch) {
          continue;
        }

        var segments = config.importPrefs && config.importPrefs.split_titles_by_bold !== false
          ? splitStoryByBoldTitles(scratch.story)
          : [{ titleText: cleanParagraphText(scratch.story.paragraphs.length ? scratch.story.paragraphs[0].contents : docxFile.displayName), bodyText: scratch.story.contents }];

        if (!segments || segments.length === 0) {
          var warnCopy = warnings.slice(0);
          warnCopy.push("no_bold_titles_detected");
          csv.writeRow({
            pagina: numbers.join("-"),
            nota: docxFile.displayName,
            warnings: warnCopy.join(";"),
            errors: ""
          });
          disposeScratch(scratch);
          continue;
        }

        for (var si = 0; si < segments.length; si++) {
          noteCounter++;
          var noteId = "nota" + noteCounter;
          var noteWarnings = warnings.slice(0);
          var titleText = segments[si].titleText || "";
          var bodyText = segments[si].bodyText || "";

          var textFrame = findLabeledItem(pages, noteId + "_texto");
          var titleFrame = findLabeledItem(pages, noteId + "_titulo");

          if (!textFrame || !textFrame.isValid) {
            noteWarnings.push("text_frame_missing:" + noteId + "_texto");
            csv.writeRow({
              pagina: numbers.join("-"),
              nota: noteId,
              chars: bodyText.length + titleText.length,
              warnings: noteWarnings.join(";"),
              errors: ""
            });
            continue;
          }

          var fullText = bodyText;
          var titleStory = null;

          if (titleFrame && titleFrame.isValid) {
            titleStory = writeStory(titleFrame, titleText);
            applyStyleAndSize(titleStory, config.styles.title.name, config.styles.title.pt_base, config.styles.title.pt_min, config.styles.title.pt_max);
          } else if (titleText) {
            fullText = titleText + "\r" + (bodyText || "");
            noteWarnings.push("title_frame_missing:" + noteId + "_titulo");
          }

          var bodyStory = writeStory(textFrame, fullText);
          if (titleFrame && titleFrame.isValid && !titleText) {
            clearStory(titleFrame);
          }

          if (bodyStory && bodyStory.isValid) {
            applyStyleAndSize(bodyStory, config.styles.body.name, config.styles.body.pt_base, config.styles.body.pt_min, config.styles.body.pt_max);
          }

          if (bodyStory && bodyStory.isValid && (!titleFrame || !titleFrame.isValid) && titleText) {
            applyStyleToFirstParagraph(bodyStory, config.styles.title.name, config.styles.title.pt_base, config.styles.title.pt_min, config.styles.title.pt_max);
          }

          var oversetResult = resolveOverset(textFrame, titleFrame, config.overset, config.styles);

          if (oversetResult.overset === "overset_hard") {
            noteWarnings.push("overset_hard");
          }

          var images = imagesMap[noteCounter] || [];
          for (var im = 0; im < images.length; im++) {
            var photo = images[im];
            var frame = findLabeledItem(pages, noteId + "_foto" + photo.index);
            if (!frame || !frame.isValid) {
              noteWarnings.push("image_frame_missing:" + noteId + "_foto" + photo.index);
              continue;
            }
            try {
              var appliedStyle = false;
              if (config.images.default_object_style) {
                try {
                  var style = app.activeDocument.objectStyles.itemByName(config.images.default_object_style);
                  if (style && style.isValid) {
                    frame.appliedObjectStyle = style;
                    appliedStyle = true;
                  }
                } catch (_) {}
              }
              if (!appliedStyle && config.images.fallback_object_style) {
                try {
                  var fb = app.activeDocument.objectStyles.itemByName(config.images.fallback_object_style);
                  if (fb && fb.isValid) {
                    frame.appliedObjectStyle = fb;
                    appliedStyle = true;
                  }
                } catch (_) {}
              }
              frame.place(photo.file);
              frame.fit(FitOptions.PROPORTIONALLY);
              frame.fit(FitOptions.CENTER_CONTENT);
            } catch (e) {
              noteWarnings.push("place_fail:" + photo.file.displayName);
            }
          }

          var row = {
            pagina: numbers.join("-"),
            nota: noteId,
            chars: bodyText.length + titleText.length,
            body_pt: oversetResult.bodyPointSize,
            title_pt: oversetResult.titlePointSize,
            overset: oversetResult.overset,
            excedente_aprox: oversetResult.exceed,
            warnings: noteWarnings.join(";"),
            errors: ""
          };
          csv.writeRow(row);
        }

        disposeScratch(scratch);
      }
    }

    csv.close();

    var exportConfig = config["export"];
    if (exportConfig && exportConfig.pdf_per_page) {
      var pdfFolder = ensureFolder(Folder(root.fsName + "/muestras"));
      var processedUnique = {};
      for (var pp = 0; pp < processedPages.length; pp++) {
        var page = processedPages[pp];
        if (!page || !page.isValid) {
          continue;
        }
        var key = page.documentOffset;
        if (processedUnique[key]) {
          continue;
        }
        processedUnique[key] = true;
        var pdfFile = File(pdfFolder.fsName + "/" + (page.documentOffset + 1) + ".pdf");
        var err = exportPagePDF(doc, page, pdfFile, exportConfig.pdf_preset_name);
        if (err) {
          $.writeln("PDF export warning: " + err);
        }
      }
    }

    alert("Proceso finalizado. Reporte en " + csvPath.fsName + ".");
  }

  run();
})();


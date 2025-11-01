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
    layout_solver: {
      enabled: true,
      max_attempts_small: 6,
      max_attempts_medium: 30,
      max_attempts_large: 120,
      strategy: "best_first",
      photo: {
        two_col_w_cm: 10.145,
        min_h_cm: 5.35,
        slot_strict: true
      }
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
    "slot_usado",
    "ancho_col",
    "alto_final",
    "pt_titulo",
    "pt_cuerpo",
    "overflow_chars",
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

  function pointsFromCm(value) {
    if (!value && value !== 0) {
      return 0;
    }
    return value * 28.34645669;
  }

  function readLayoutSlotsReport(file) {
    var out = {};
    if (!file || !file.exists) {
      return out;
    }
    var data = null;
    try {
      file.open("r");
      var raw = file.read();
      file.close();
      if (raw && raw.length) {
        data = JSON.parse(raw);
      }
    } catch (e) {
      try { file.close(); } catch (_) {}
      $.writeln("layout_slots_report read error: " + e);
      return out;
    }

    if (!data || !(data instanceof Array)) {
      return out;
    }

    var perPageCounter = {};
    for (var i = 0; i < data.length; i++) {
      var slot = data[i];
      if (!slot || typeof slot.page === "undefined") {
        continue;
      }
      var pageKey = String(slot.page);
      if (!out[pageKey]) {
        out[pageKey] = [];
        perPageCounter[pageKey] = 0;
      }
      perPageCounter[pageKey]++;
      var width = slot.w_pt ? slot.w_pt : pointsFromCm(slot.w_cm || 0);
      var height = slot.h_pt ? slot.h_pt : pointsFromCm(slot.h_cm || 0);
      var x = slot.x_pt ? slot.x_pt : pointsFromCm(slot.x_cm || 0);
      var y = slot.y_pt ? slot.y_pt : pointsFromCm(slot.y_cm || 0);
      var bounds = [y, x, y + height, x + width];
      out[pageKey].push({
        id: slot.id ? slot.id : (pageKey + "_slot" + perPageCounter[pageKey]),
        page: slot.page,
        bounds: bounds,
        width: width,
        height: height,
        raw: slot,
        isPhoto: slot.is_photo_slot || slot.is_2col_photo_slot || false,
        fitsPhoto2ColHeight: slot.fits_photo_2col_h || false
      });
    }
    return out;
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
    var prefs = null;
    var previous = {};
    try {
      if (typeof app.wordImportPreferences !== "undefined") {
        prefs = app.wordImportPreferences;
      } else if (typeof app.wordRTFImportPreferences !== "undefined") {
        prefs = app.wordRTFImportPreferences;
      }
    } catch (_) {}

    if (prefs) {
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
      if (prefs) {
        try {
          for (var key2 in previous) {
            prefs[key2] = previous[key2];
          }
        } catch (e2) {}
      }
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

  function expandFrameHeightMM(textFrame, delta, limitBottom) {
    if (!textFrame || !textFrame.isValid || !delta) {
      return;
    }
    var pts = delta * 2.834645669;
    try {
      var gb = textFrame.geometricBounds;
      var nextBottom = gb[2] + pts;
      if (limitBottom !== undefined && limitBottom !== null) {
        if (nextBottom > limitBottom) {
          nextBottom = limitBottom;
        }
      }
      gb[2] = nextBottom;
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

  function resolveOverset(bodyFrame, titleFrame, cfg, stylesCfg, limits) {
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
      expandFrameHeightMM(bodyFrame, step, limits && limits.bodyMaxBottom !== undefined ? limits.bodyMaxBottom : null);
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

    if (info.over) {
      var titleContainer = null;
      if (titleFrame && titleFrame.isValid) {
        titleContainer = titleFrame;
      } else {
        titleContainer = bodyFrame;
      }
      if (titleContainer && titleContainer.isValid) {
        var titleStory = titleContainer.parentStory;
        var titleDropped = 0;
        while (info.over && titleDropped < cfg.title_max_drop_pt) {
          var tDelta = -Math.min(cfg.body_step_pt, cfg.title_max_drop_pt - titleDropped);
          var appliedTitle = nudgeStyle(titleContainer, stylesCfg.title.name, tDelta, stylesCfg.title.pt_min, stylesCfg.title.pt_max);
          if (appliedTitle === null) {
            break;
          }
          titleDropped += Math.abs(tDelta);
          titleStory.recompose();
          info = storyOversetInfo(story);
        }
      }
    }

    info = storyOversetInfo(story);
    result.exceed = info.exceed;
    result.bodyPointSize = currentStylePointSize(bodyFrame, stylesCfg.body.name);
    result.titlePointSize = currentStylePointSize(titleFrame && titleFrame.isValid ? titleFrame : bodyFrame, stylesCfg.title.name);
    if (info.over) {
      result.overset = "overset_hard";
    }
    return result;
  }

  function toCm(points) {
    return points / 28.34645669;
  }

  function formatNumber(value, decimals) {
    var factor = Math.pow(10, decimals || 2);
    return Math.round(value * factor) / factor;
  }

  function cloneBounds(bounds) {
    if (!bounds) {
      return null;
    }
    return [bounds[0], bounds[1], bounds[2], bounds[3]];
  }

  function slotsForPage(slotsMap, page) {
    if (!slotsMap || !page || !page.isValid) {
      return [];
    }
    var nameKey = String(page.name);
    if (slotsMap.hasOwnProperty(nameKey)) {
      return slotsMap[nameKey].slice(0);
    }
    var indexKey = String(page.documentOffset + 1);
    if (slotsMap.hasOwnProperty(indexKey)) {
      return slotsMap[indexKey].slice(0);
    }
    return [];
  }

  function splitSlotsByType(slots) {
    var out = { text: [], photo: [] };
    if (!slots) {
      return out;
    }
    for (var i = 0; i < slots.length; i++) {
      var slot = slots[i];
      if (!slot) {
        continue;
      }
      if (slot.isPhoto) {
        out.photo.push(slot);
      } else {
        out.text.push(slot);
      }
    }
    return out;
  }

  function slotColumnCount(slot) {
    if (!slot || !slot.raw) {
      return 1;
    }
    var raw = slot.raw;
    if (raw.text_column_count) {
      return Math.max(1, raw.text_column_count);
    }
    if (raw.columns) {
      return Math.max(1, raw.columns);
    }
    if (raw.col_count) {
      return Math.max(1, raw.col_count);
    }
    return 1;
  }

  function createMeasurementFrameForSlot(doc, slot, columnCount) {
    var page = doc.pages[0];
    var top = 10000;
    var left = 10000;
    var height = slot && slot.height ? slot.height : pointsFromCm(25);
    var width = slot && slot.width ? slot.width : pointsFromCm(4);
    var frame = page.textFrames.add();
    frame.geometricBounds = [top, left, top + height, left + width];
    frame.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
    try {
      frame.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    } catch (_) {}
    try {
      frame.textFramePreferences.textColumnCount = Math.max(1, columnCount);
    } catch (_) {}
    try {
      frame.textFramePreferences.useFlexibleColumnWidth = false;
    } catch (_) {}
    return frame;
  }

  function applyTitleSpan(paragraph, span) {
    if (!paragraph || !paragraph.isValid) {
      return;
    }
    try {
      if (span > 1) {
        if (paragraph.spanColumnType !== undefined) {
          paragraph.spanColumnType = SpanColumnType.SPAN_COLUMNS;
        }
        if (paragraph.spanColumnCount !== undefined) {
          paragraph.spanColumnCount = span;
        }
      } else {
        if (paragraph.spanColumnType !== undefined) {
          paragraph.spanColumnType = SpanColumnType.SINGLE_COLUMN;
        }
      }
    } catch (_) {}
  }

  function LayoutSolver(doc, pages, slots, config, stylesCfg) {
    this.doc = doc;
    this.pages = pages;
    this.slots = slots.slice(0);
    this.config = config;
    this.stylesCfg = stylesCfg;
    this.sizeProfiles = this.buildSizeProfiles();
    this.measurementCache = {};
  }

  LayoutSolver.prototype.determineMaxAttempts = function (noteCount) {
    var solverCfg = this.config.layout_solver || {};
    if (noteCount <= 1) {
      return solverCfg.max_attempts_small || 6;
    }
    if (noteCount === 2) {
      return solverCfg.max_attempts_small || 6;
    }
    if (noteCount <= 4) {
      return solverCfg.max_attempts_medium || 30;
    }
    return solverCfg.max_attempts_large || 120;
  };

  LayoutSolver.prototype.buildSizeProfiles = function () {
    var profiles = [];
    var body = this.stylesCfg.body;
    var title = this.stylesCfg.title;
    var bodyStep = this.config.overset.body_step_pt || 0.25;
    var deltas = [0, -bodyStep, bodyStep, -2 * bodyStep];
    var titleDeltas = [0, -0.5, 0.5];
    var added = {};

    function clamp(val, min, max) {
      if (val < min) return min;
      if (val > max) return max;
      return val;
    }

    for (var i = 0; i < deltas.length; i++) {
      var bodyPt = clamp(body.pt_base + deltas[i], body.pt_min, body.pt_max);
      for (var j = 0; j < titleDeltas.length; j++) {
        var titlePt = clamp(title.pt_base + titleDeltas[j], title.pt_min, title.pt_max);
        var key = bodyPt + "_" + titlePt;
        if (!added[key]) {
          profiles.push({ body: bodyPt, title: titlePt, key: key });
          added[key] = true;
        }
      }
    }

    if (!profiles.length) {
      profiles.push({ body: body.pt_base, title: title.pt_base, key: body.pt_base + "_" + title.pt_base });
    }
    return profiles;
  };

  LayoutSolver.prototype.measureBody = function (note, slot, profile) {
    var cacheKey = ["body", note.noteId, slot.id, profile.key].join("|");
    if (this.measurementCache.hasOwnProperty(cacheKey)) {
      return this.measurementCache[cacheKey];
    }
    var result = {
      overflow: note.bodyText ? note.bodyText.length : 0,
      oversetCode: "overset_hard",
      bodyPointSize: profile.body,
      height: slot.height,
      columnCount: Math.max(1, slotColumnCount(slot)),
      columnWidth: slot.width,
    };
    var frame = null;
    try {
      frame = createMeasurementFrameForSlot(this.doc, slot, result.columnCount);
      frame.geometricBounds = [frame.geometricBounds[0], frame.geometricBounds[1], frame.geometricBounds[0] + slot.height, frame.geometricBounds[1] + slot.width];
      var story = writeStory(frame, note.bodyText || "");
      applyStyleAndSize(story, this.stylesCfg.body.name, profile.body, this.stylesCfg.body.pt_min, this.stylesCfg.body.pt_max);
      var overset = resolveOverset(frame, null, this.config.overset || {}, this.stylesCfg, { bodyMaxBottom: frame.geometricBounds[0] + slot.height });
      var info = storyOversetInfo(story);
      result.overflow = overset.exceed || info.exceed || 0;
      result.oversetCode = overset.overset || (info.over ? "overset_hard" : "");
      result.bodyPointSize = overset.bodyPointSize || profile.body;
      result.height = frame.geometricBounds[2] - frame.geometricBounds[0];
      result.columnWidth = slot.width / result.columnCount;
    } catch (e) {
      $.writeln("measureBody: " + e);
    } finally {
      if (frame && frame.isValid) {
        try { frame.remove(); } catch (_) {}
      }
    }
    this.measurementCache[cacheKey] = result;
    return result;
  };

  LayoutSolver.prototype.measureTitle = function (note, slot, profile, span) {
    if (!note.titleText) {
      return null;
    }
    var cacheKey = ["title", note.noteId, slot.id, profile.key, span].join("|");
    if (this.measurementCache.hasOwnProperty(cacheKey)) {
      return this.measurementCache[cacheKey];
    }
    var result = {
      overflow: 0,
      oversetCode: "",
      titlePointSize: profile.title,
      height: 0,
      span: span,
      columnCount: Math.max(1, slotColumnCount(slot))
    };
    var frame = null;
    try {
      frame = createMeasurementFrameForSlot(this.doc, slot, result.columnCount);
      frame.geometricBounds = [frame.geometricBounds[0], frame.geometricBounds[1], frame.geometricBounds[0] + slot.height, frame.geometricBounds[1] + slot.width];
      var story = writeStory(frame, note.titleText || "");
      applyStyleAndSize(story, this.stylesCfg.title.name, profile.title, this.stylesCfg.title.pt_min, this.stylesCfg.title.pt_max);
      if (story.paragraphs.length > 0) {
        applyTitleSpan(story.paragraphs[0], span);
      }
      story.recompose();
      var info = storyOversetInfo(story);
      if (info.over) {
        var dropped = 0;
        var step = this.config.overset.body_step_pt || 0.25;
        while (info.over && dropped < (this.config.overset.title_max_drop_pt || 0.5)) {
          var para = story.paragraphs.length ? story.paragraphs[0] : null;
          if (!para) {
            break;
          }
          var next = para.pointSize - step;
          if (next < this.stylesCfg.title.pt_min) {
            next = this.stylesCfg.title.pt_min;
          }
          if (next === para.pointSize) {
            break;
          }
          para.pointSize = next;
          dropped += step;
          story.recompose();
          info = storyOversetInfo(story);
        }
        result.titlePointSize = story.paragraphs.length ? story.paragraphs[0].pointSize : profile.title;
        result.overflow = info.exceed || 0;
        if (info.over) {
          result.oversetCode = "overset_hard";
        }
      }
      result.height = frame.geometricBounds[2] - frame.geometricBounds[0];
    } catch (e) {
      $.writeln("measureTitle: " + e);
    } finally {
      if (frame && frame.isValid) {
        try { frame.remove(); } catch (_) {}
      }
    }
    this.measurementCache[cacheKey] = result;
    return result;
  };

  LayoutSolver.prototype.measureCombined = function (note, slot, profile, span) {
    var cacheKey = ["combo", note.noteId, slot.id, profile.key, span].join("|");
    if (this.measurementCache.hasOwnProperty(cacheKey)) {
      return this.measurementCache[cacheKey];
    }
    var result = {
      overflow: 0,
      oversetCode: "",
      bodyPointSize: profile.body,
      titlePointSize: profile.title,
      height: slot.height,
      span: span,
      columnCount: Math.max(1, slotColumnCount(slot)),
      columnWidth: slot.width
    };
    var frame = null;
    try {
      frame = createMeasurementFrameForSlot(this.doc, slot, result.columnCount);
      frame.geometricBounds = [frame.geometricBounds[0], frame.geometricBounds[1], frame.geometricBounds[0] + slot.height, frame.geometricBounds[1] + slot.width];
      var text = note.bodyText || "";
      if (note.titleText) {
        text = note.titleText + (text ? "
" + text : "");
      }
      var story = writeStory(frame, text);
      applyStyleAndSize(story, this.stylesCfg.body.name, profile.body, this.stylesCfg.body.pt_min, this.stylesCfg.body.pt_max);
      if (note.titleText) {
        applyStyleToFirstParagraph(story, this.stylesCfg.title.name, profile.title, this.stylesCfg.title.pt_min, this.stylesCfg.title.pt_max);
        if (story.paragraphs.length > 0) {
          applyTitleSpan(story.paragraphs[0], span);
        }
      }
      var overset = resolveOverset(frame, null, this.config.overset || {}, this.stylesCfg, { bodyMaxBottom: frame.geometricBounds[0] + slot.height });
      var info = storyOversetInfo(story);
      result.overflow = overset.exceed || info.exceed || 0;
      result.oversetCode = overset.overset || (info.over ? "overset_hard" : "");
      result.bodyPointSize = overset.bodyPointSize || profile.body;
      result.titlePointSize = overset.titlePointSize || profile.title;
      result.height = frame.geometricBounds[2] - frame.geometricBounds[0];
      result.columnWidth = slot.width / result.columnCount;
    } catch (e) {
      $.writeln("measureCombined: " + e);
    } finally {
      if (frame && frame.isValid) {
        try { frame.remove(); } catch (_) {}
      }
    }
    this.measurementCache[cacheKey] = result;
    return result;
  };

  LayoutSolver.prototype.sortNotes = function (notes) {
    var ordered = notes.slice(0);
    ordered.sort(function (a, b) {
      var aTitle = a.titleText ? a.titleText.length : 0;
      var bTitle = b.titleText ? b.titleText.length : 0;
      if (bTitle !== aTitle) {
        return bTitle - aTitle;
      }
      var aBody = a.bodyText ? a.bodyText.length : 0;
      var bBody = b.bodyText ? b.bodyText.length : 0;
      return bBody - aBody;
    });
    return ordered;
  };

  LayoutSolver.prototype.pageForSlot = function (slot) {
    return slot && slot.page !== undefined ? slot.page : null;
  };

  LayoutSolver.prototype.computeOptionsForNote = function (note) {
    var bodyCandidates = [];
    var titleCandidates = [];
    var combinedCandidates = [];
    for (var si = 0; si < this.slots.length; si++) {
      var slot = this.slots[si];
      for (var pi = 0; pi < this.sizeProfiles.length; pi++) {
        var profile = this.sizeProfiles[pi];
        var bodyMetrics = this.measureBody(note, slot, profile);
        bodyCandidates.push({
          slot: slot,
          profile: profile,
          overflow: bodyMetrics.overflow,
          oversetCode: bodyMetrics.oversetCode,
          bodyPointSize: bodyMetrics.bodyPointSize,
          height: bodyMetrics.height,
          columnCount: bodyMetrics.columnCount,
          columnWidth: bodyMetrics.columnWidth
        });
        if (note.titleText) {
          for (var span = 1; span <= 5; span++) {
            var comboMetrics = this.measureCombined(note, slot, profile, span);
            combinedCandidates.push({
              slot: slot,
              profile: profile,
              span: span,
              overflow: comboMetrics.overflow,
              oversetCode: comboMetrics.oversetCode,
              bodyPointSize: comboMetrics.bodyPointSize,
              titlePointSize: comboMetrics.titlePointSize,
              height: comboMetrics.height,
              columnCount: comboMetrics.columnCount,
              columnWidth: comboMetrics.columnWidth
            });
            var titleMetrics = this.measureTitle(note, slot, profile, span);
            if (titleMetrics) {
              titleCandidates.push({
                slot: slot,
                profile: profile,
                span: span,
                overflow: titleMetrics.overflow,
                oversetCode: titleMetrics.oversetCode,
                titlePointSize: titleMetrics.titlePointSize,
                height: titleMetrics.height,
                columnCount: titleMetrics.columnCount
              });
            }
          }
        }
      }
    }

    function sortByOverflow(a, b) {
      if (a.oversetCode === "overset_hard" && b.oversetCode !== "overset_hard") {
        return 1;
      }
      if (b.oversetCode === "overset_hard" && a.oversetCode !== "overset_hard") {
        return -1;
      }
      if (a.overflow !== b.overflow) {
        return a.overflow - b.overflow;
      }
      var aWidth = a.columnWidth || (a.slot ? a.slot.width : 0);
      var bWidth = b.columnWidth || (b.slot ? b.slot.width : 0);
      return bWidth - aWidth;
    }

    bodyCandidates.sort(sortByOverflow);
    titleCandidates.sort(sortByOverflow);
    combinedCandidates.sort(sortByOverflow);

    var limit = 8;
    if (bodyCandidates.length > limit) {
      bodyCandidates = bodyCandidates.slice(0, limit);
    }
    if (titleCandidates.length > limit) {
      titleCandidates = titleCandidates.slice(0, limit);
    }
    if (combinedCandidates.length > limit) {
      combinedCandidates = combinedCandidates.slice(0, limit);
    }

    var options = [];
    var minOverflow = null;

    for (var ci = 0; ci < combinedCandidates.length; ci++) {
      var comb = combinedCandidates[ci];
      var totalOverflow = comb.overflow || 0;
      var hard = comb.oversetCode === "overset_hard";
      options.push({
        type: "combined",
        bodySlot: comb.slot,
        titleSlot: comb.slot,
        bodyPointSize: comb.bodyPointSize,
        titlePointSize: comb.titlePointSize,
        titleSpan: comb.span,
        totalOverflow: totalOverflow,
        oversetFlags: hard ? ["overset_hard"] : [],
        bodyHeight: comb.height,
        columnWidth: comb.columnWidth,
        profile: comb.profile,
        titleHeight: comb.height
      });
      if (minOverflow === null || totalOverflow < minOverflow) {
        minOverflow = totalOverflow;
      }
    }

    if (note.titleText) {
      for (var bi = 0; bi < bodyCandidates.length; bi++) {
        var body = bodyCandidates[bi];
        for (var ti = 0; ti < titleCandidates.length; ti++) {
          var title = titleCandidates[ti];
          if (body.profile.key !== title.profile.key) {
            continue;
          }
          if (!body.slot || !title.slot) {
            continue;
          }
          if (body.slot.id === title.slot.id) {
            continue;
          }
          if (this.pageForSlot(body.slot) !== this.pageForSlot(title.slot)) {
            continue;
          }
          var overflowTotal = (body.overflow || 0) + (title.overflow || 0);
          var hardFlags = [];
          if (body.oversetCode === "overset_hard") {
            hardFlags.push("body_overset");
          }
          if (title.oversetCode === "overset_hard") {
            hardFlags.push("title_overset");
          }
          options.push({
            type: "separate",
            bodySlot: body.slot,
            titleSlot: title.slot,
            bodyPointSize: body.bodyPointSize,
            titlePointSize: title.titlePointSize,
            titleSpan: title.span,
            totalOverflow: overflowTotal,
            oversetFlags: hardFlags,
            bodyHeight: body.height,
            titleHeight: title.height,
            columnWidth: body.columnWidth,
            profile: body.profile
          });
          if (minOverflow === null || overflowTotal < minOverflow) {
            minOverflow = overflowTotal;
          }
        }
      }
    } else {
      for (var bi2 = 0; bi2 < bodyCandidates.length; bi2++) {
        var bodyOnly = bodyCandidates[bi2];
        var totalOverflow2 = bodyOnly.overflow || 0;
        var flags = bodyOnly.oversetCode === "overset_hard" ? ["overset_hard"] : [];
        options.push({
          type: "body",
          bodySlot: bodyOnly.slot,
          titleSlot: null,
          bodyPointSize: bodyOnly.bodyPointSize,
          titlePointSize: "",
          titleSpan: 1,
          totalOverflow: totalOverflow2,
          oversetFlags: flags,
          bodyHeight: bodyOnly.height,
          titleHeight: 0,
          columnWidth: bodyOnly.columnWidth,
          profile: bodyOnly.profile
        });
        if (minOverflow === null || totalOverflow2 < minOverflow) {
          minOverflow = totalOverflow2;
        }
      }
    }

    if (!options.length) {
      options.push({
        type: "missing",
        bodySlot: null,
        titleSlot: null,
        bodyPointSize: this.stylesCfg.body.pt_base,
        titlePointSize: this.stylesCfg.title.pt_base,
        titleSpan: 1,
        totalOverflow: note.bodyText ? note.bodyText.length : 0,
        oversetFlags: ["slot_missing", "no_slot_available"],
        bodyHeight: 0,
        titleHeight: 0,
        columnWidth: 0,
        profile: this.sizeProfiles.length ? this.sizeProfiles[0] : { body: this.stylesCfg.body.pt_base, title: this.stylesCfg.title.pt_base, key: "default" }
      });
      minOverflow = minOverflow === null ? options[0].totalOverflow : minOverflow;
    }

    function optionScore(option, note) {
      var base = option.totalOverflow;
      for (var oi = 0; oi < option.oversetFlags.length; oi++) {
        if (option.oversetFlags[oi] === "overset_hard" || option.oversetFlags[oi] === "body_overset" || option.oversetFlags[oi] === "title_overset") {
          base += 500;
        }
      }
      var titleLen = note.titleText ? note.titleText.length : 0;
      if (titleLen > 30) {
        base -= (option.titleSpan || 1) * 0.5;
      }
      var bodyLen = note.bodyText ? note.bodyText.length : 0;
      var width = option.bodySlot ? option.bodySlot.width : 0;
      if (bodyLen > 500 && width) {
        base -= width / 120;
      }
      return base;
    }

    for (var oi = 0; oi < options.length; oi++) {
      options[oi].score = optionScore(options[oi], note);
    }

    options.sort(function (a, b) {
      if (a.score !== b.score) {
        return a.score - b.score;
      }
      return a.totalOverflow - b.totalOverflow;
    });

    return { options: options, minOverflow: minOverflow === null ? 0 : minOverflow };
  };

  LayoutSolver.prototype.solve = function (notes) {
    var orderedNotes = this.sortNotes(notes);
    var optionsMap = {};
    var minOverflowMap = {};
    for (var i = 0; i < orderedNotes.length; i++) {
      var note = orderedNotes[i];
      var info = this.computeOptionsForNote(note);
      optionsMap[note.noteId] = info.options;
      minOverflowMap[note.noteId] = info.minOverflow;
    }

    var maxAttempts = this.determineMaxAttempts(notes.length);
    var solverCfg = this.config.layout_solver || {};
    var attempts = 0;
    var bestState = null;

    function initialState() {
      return {
        index: 0,
        usedSlots: {},
        placements: [],
        totalOverflow: 0,
        hardOverset: false
      };
    }

    function cloneState(state) {
      return {
        index: state.index,
        usedSlots: clone(state.usedSlots),
        placements: state.placements.slice(0),
        totalOverflow: state.totalOverflow,
        hardOverset: state.hardOverset
      };
    }

    function remainingHeuristic(state) {
      var sum = 0;
      for (var ri = state.index; ri < orderedNotes.length; ri++) {
        var pendingNote = orderedNotes[ri];
        sum += minOverflowMap[pendingNote.noteId] || 0;
      }
      return sum;
    }

    function stateScore(state) {
      return state.totalOverflow + remainingHeuristic(state) + (state.hardOverset ? 1000 : 0);
    }

    var open = [initialState()];

    while (open.length && attempts < maxAttempts) {
      open.sort(function (a, b) { return stateScore(a) - stateScore(b); });
      var current = open.shift();
      if (current.index >= orderedNotes.length) {
        if (!bestState) {
          bestState = current;
        } else {
          var better = false;
          if (!current.hardOverset && bestState.hardOverset) {
            better = true;
          } else if (current.hardOverset === bestState.hardOverset) {
            if (current.totalOverflow < bestState.totalOverflow) {
              better = true;
            }
          }
          if (better) {
            bestState = current;
          }
        }
        if (bestState && !bestState.hardOverset && bestState.totalOverflow === 0 && solverCfg.strategy === "best_first") {
          break;
        }
        attempts++;
        continue;
      }

      var note = orderedNotes[current.index];
      var noteOptions = optionsMap[note.noteId] || [];
      if (!noteOptions.length) {
        var fallback = cloneState(current);
        fallback.index++;
        fallback.placements.push({
          note: note,
          option: {
            type: "missing",
            bodySlot: null,
            titleSlot: null,
            bodyPointSize: this.stylesCfg.body.pt_base,
            titlePointSize: this.stylesCfg.title.pt_base,
            titleSpan: 1,
            totalOverflow: note.bodyText ? note.bodyText.length : 0,
            oversetFlags: ["slot_missing"],
            bodyHeight: 0,
            titleHeight: 0,
            columnWidth: 0
          }
        });
        fallback.totalOverflow += note.bodyText ? note.bodyText.length : 0;
        fallback.hardOverset = true;
        open.push(fallback);
        attempts++;
        continue;
      }

      for (var oi = 0; oi < noteOptions.length && attempts < maxAttempts; oi++) {
        var option = noteOptions[oi];
        var blocked = false;
        if (option.bodySlot && current.usedSlots[option.bodySlot.id]) {
          blocked = true;
        }
        if (!blocked && option.titleSlot && option.titleSlot !== option.bodySlot && current.usedSlots[option.titleSlot.id]) {
          blocked = true;
        }
        if (blocked) {
          continue;
        }
        var nextState = cloneState(current);
        nextState.index++;
        if (option.bodySlot) {
          nextState.usedSlots[option.bodySlot.id] = true;
        }
        if (option.titleSlot && option.titleSlot !== option.bodySlot) {
          nextState.usedSlots[option.titleSlot.id] = true;
        }
        nextState.totalOverflow += option.totalOverflow;
        if (option.oversetFlags && option.oversetFlags.length) {
          for (var fi = 0; fi < option.oversetFlags.length; fi++) {
            if (option.oversetFlags[fi] === "overset_hard" || option.oversetFlags[fi] === "body_overset" || option.oversetFlags[fi] === "title_overset") {
              nextState.hardOverset = true;
              break;
            }
          }
        }
        nextState.placements.push({ note: note, option: option });
        open.push(nextState);
        attempts++;
      }
    }

    var output = { noteResults: [], success: false, totalOverflow: 0 };
    if (!bestState) {
      return output;
    }

    var placementMap = {};
    for (var pi = 0; pi < bestState.placements.length; pi++) {
      placementMap[bestState.placements[pi].note.noteId] = bestState.placements[pi].option;
    }

    for (var ni = 0; ni < notes.length; ni++) {
      var originalNote = notes[ni];
      var chosen = placementMap[originalNote.noteId];
      if (!chosen) {
        chosen = {
          type: "missing",
          bodySlot: null,
          titleSlot: null,
          bodyPointSize: this.stylesCfg.body.pt_base,
          titlePointSize: this.stylesCfg.title.pt_base,
          titleSpan: 1,
          totalOverflow: originalNote.bodyText ? originalNote.bodyText.length : 0,
          oversetFlags: ["slot_missing"],
          bodyHeight: 0,
          titleHeight: 0,
          columnWidth: 0
        };
      }
      output.noteResults.push({
        note: originalNote,
        slot: chosen.bodySlot,
        titleSlot: chosen.titleSlot,
        layoutType: chosen.type,
        oversetFlags: chosen.oversetFlags || [],
        bodyPointSize: chosen.bodyPointSize,
        titlePointSize: chosen.titlePointSize,
        titleSpan: chosen.titleSpan,
        columnWidth: chosen.columnWidth,
        bodyHeight: chosen.bodyHeight,
        titleHeight: chosen.titleHeight,
        totalOverflow: chosen.totalOverflow
      });
    }

    output.totalOverflow = bestState.totalOverflow;
    output.success = !bestState.hardOverset && bestState.totalOverflow === 0;
    return output;
  };
  function removeItemsByLabel(page, label) {
    if (!page || !page.isValid || !label) {
      return;
    }
    var items = page.allPageItems;
    for (var i = items.length - 1; i >= 0; i--) {
      try {
        if (items[i] && items[i].isValid && items[i].label === label) {
          items[i].remove();
        }
      } catch (_) {}
    }
  }

  function createTextFrameForBounds(page, bounds, label) {
    if (!page || !page.isValid || !bounds) {
      return null;
    }
    var frame = page.textFrames.add();
    frame.geometricBounds = cloneBounds(bounds);
    if (label) {
      frame.label = label;
    }
    frame.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
    return frame;
  }

  function ensureOversetSwatch(doc) {
    var name = "AutoLayoutOversetRed";
    if (!doc || !doc.isValid) {
      return null;
    }
    try {
      var existing = doc.colors.itemByName(name);
      if (existing && existing.isValid) {
        return existing;
      }
    } catch (_) {}
    try {
      return doc.colors.add({ name: name, model: ColorModel.PROCESS, space: ColorSpace.CMYK, colorValue: [0, 100, 100, 0] });
    } catch (e) {
      try {
        var swatch = doc.swatches.itemByName("[Black]");
        if (swatch && swatch.isValid) {
          return swatch;
        }
      } catch (_) {}
    }
    return null;
  }

  function markOversetFrame(frame) {
    if (!frame || !frame.isValid) {
      return;
    }
    try {
      var doc = app.activeDocument;
      var swatch = ensureOversetSwatch(doc);
      frame.strokeWeight = 2;
      if (swatch && swatch.isValid) {
        frame.strokeColor = swatch;
      }
      frame.strokeTint = 100;
      frame.strokeType = doc.strokeStyles.itemByName("Solid");
    } catch (_) {}
  }

  function applyImageToFrame(frame, photoFile, config, warnings) {
    if (!frame || !frame.isValid || !photoFile) {
      return;
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
      frame.place(photoFile);
      frame.fit(FitOptions.PROPORTIONALLY);
      frame.fit(FitOptions.CENTER_CONTENT);
    } catch (e) {
      if (warnings) {
        warnings.push("place_fail:" + photoFile.displayName);
      }
    }
  }

  function placeNoteImages(note, targetPage, photoSlots, config, warnings) {
    if (!note || !note.images || !note.images.length) {
      return;
    }
    var solverPhotoCfg = config.layout_solver && config.layout_solver.photo ? config.layout_solver.photo : {};
    for (var i = 0; i < note.images.length; i++) {
      var photo = note.images[i];
      var slot = null;
      for (var si = 0; si < photoSlots.length; si++) {
        var candidate = photoSlots[si];
        if (!candidate || candidate.used) {
          continue;
        }
        if (targetPage && candidate.pageRef && candidate.pageRef !== targetPage) {
          continue;
        }
        if (solverPhotoCfg.slot_strict) {
          if (!candidate.isPhoto || !candidate.fitsPhoto2ColHeight) {
            continue;
          }
          if (solverPhotoCfg.two_col_w_cm) {
            var widthCm = toCm(candidate.width || (candidate.bounds ? (candidate.bounds[3] - candidate.bounds[1]) : 0));
            if (Math.abs(widthCm - solverPhotoCfg.two_col_w_cm) > 0.4) {
              continue;
            }
          }
          if (solverPhotoCfg.min_h_cm) {
            var heightCm = toCm(candidate.height || (candidate.bounds ? (candidate.bounds[2] - candidate.bounds[0]) : 0));
            if (heightCm < solverPhotoCfg.min_h_cm) {
              continue;
            }
          }
        }
        slot = candidate;
        break;
      }
      if (!slot) {
        if (warnings) {
          warnings.push("photo_slot_missing:" + note.noteId);
        }
        continue;
      }
      slot.used = true;
      var pageForSlot = slot.pageRef || targetPage;
      if (!pageForSlot || !pageForSlot.isValid) {
        if (warnings) {
          warnings.push("photo_slot_page_missing:" + slot.id);
        }
        continue;
      }
      var bounds = cloneBounds(slot.bounds);
      removeItemsByLabel(pageForSlot, note.noteId + "_foto" + photo.index);
      var frame = pageForSlot.rectangles.add();
      frame.geometricBounds = bounds;
      frame.label = note.noteId + "_foto" + photo.index;
      applyImageToFrame(frame, photo.file, config, warnings);
    }
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
    var layoutSlots = readLayoutSlotsReport(File(root.fsName + "/layout_slots_report.json"));
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
        var warnNoDoc = warnings.slice(0);
        warnNoDoc.push("no_docx_found");
        csv.writeRow({
          pagina: numbers.join("-"),
          nota: "",
          slot_usado: "",
          ancho_col: "",
          alto_final: "",
          pt_titulo: "",
          pt_cuerpo: "",
          overflow_chars: "",
          warnings: warnNoDoc.join(";"),
          errors: "",
        });
        continue;
      }

      var noteCounter = 0;
      var imagesMap = collectImages(pageFolder);
      var notesData = [];

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
            slot_usado: "",
            ancho_col: "",
            alto_final: "",
            pt_titulo: "",
            pt_cuerpo: "",
            overflow_chars: "",
            warnings: warnCopy.join(";"),
            errors: "",
          });
          disposeScratch(scratch);
          continue;
        }

        for (var si = 0; si < segments.length; si++) {
          noteCounter++;
          var noteId = "nota" + noteCounter;
          var titleText = segments[si].titleText || "";
          var bodyText = segments[si].bodyText || "";
          var noteWarnings = warnings.slice(0);
          var noteImages = imagesMap[noteCounter] || [];
          notesData.push({
            noteId: noteId,
            titleText: titleText,
            bodyText: bodyText,
            warnings: noteWarnings,
            images: noteImages,
            pageLabel: numbers.join("-"),
          });
        }

        disposeScratch(scratch);
      }

      if (!notesData.length) {
        continue;
      }

      var textSlotsAll = [];
      var photoSlotsAll = [];
      for (var pg = 0; pg < pages.length; pg++) {
        var page = pages[pg];
        var pageSlots = splitSlotsByType(slotsForPage(layoutSlots, page));
        for (var ts = 0; ts < pageSlots.text.length; ts++) {
          var tSlot = pageSlots.text[ts];
          tSlot.pageRef = page;
          textSlotsAll.push(tSlot);
        }
        for (var ps = 0; ps < pageSlots.photo.length; ps++) {
          var pSlot = pageSlots.photo[ps];
          pSlot.pageRef = page;
          pSlot.used = false;
          photoSlotsAll.push(pSlot);
        }
      }

      var finalResults = [];
      if (!textSlotsAll.length) {
        for (var nd = 0; nd < notesData.length; nd++) {
          var missingNote = notesData[nd];
          finalResults.push({
            note: missingNote,
            slot: null,
            titleSlot: null,
            layoutType: "missing",
            oversetFlags: ["slot_missing", "no_slot_available"],
            bodyPointSize: config.styles.body.pt_base,
            titlePointSize: config.styles.title.pt_base,
            titleSpan: 1,
            columnWidth: 0,
            bodyHeight: 0,
            titleHeight: 0,
            totalOverflow: missingNote.bodyText ? missingNote.bodyText.length : 0
          });
        }
      } else if (config.layout_solver && config.layout_solver.enabled === false) {
        for (var nd2 = 0; nd2 < notesData.length; nd2++) {
          var manualNote = notesData[nd2];
          var manualSlot = nd2 < textSlotsAll.length ? textSlotsAll[nd2] : null;
          finalResults.push({
            note: manualNote,
            slot: manualSlot,
            titleSlot: manualSlot,
            layoutType: manualSlot ? "combined" : "missing",
            oversetFlags: manualSlot ? [] : ["slot_missing", "no_slot_available"],
            bodyPointSize: config.styles.body.pt_base,
            titlePointSize: config.styles.title.pt_base,
            titleSpan: 1,
            columnWidth: manualSlot ? manualSlot.width : 0,
            bodyHeight: manualSlot ? (manualSlot.bounds[2] - manualSlot.bounds[0]) : 0,
            titleHeight: manualSlot ? (manualSlot.bounds[2] - manualSlot.bounds[0]) : 0,
            totalOverflow: 0
          });
        }
      } else {
        var solver = new LayoutSolver(doc, pages, textSlotsAll, config, config.styles);
        var solution = solver.solve(notesData);
        finalResults = solution.noteResults || [];
      }

      if (finalResults.length > notesData.length) {
        finalResults = finalResults.slice(0, notesData.length);
      }

      if (finalResults.length < notesData.length) {
        for (var pad = finalResults.length; pad < notesData.length; pad++) {
          var padNote = notesData[pad];
          finalResults.push({
            note: padNote,
            slot: null,
            titleSlot: null,
            layoutType: "missing",
            oversetFlags: ["solver_no_result", "slot_missing"],
            bodyPointSize: config.styles.body.pt_base,
            titlePointSize: config.styles.title.pt_base,
            titleSpan: 1,
            columnWidth: 0,
            bodyHeight: 0,
            titleHeight: 0,
            totalOverflow: padNote.bodyText ? padNote.bodyText.length : 0
          });
        }
      }

      for (var nr = 0; nr < finalResults.length; nr++) {
        var result = finalResults[nr];
        var note = result.note || notesData[nr];
        if (!result.note) {
          result.note = note;
        }
        var bodySlot = result.slot;
        var titleSlot = result.titleSlot;
        var layoutType = result.layoutType || (bodySlot ? "combined" : "missing");
        var targetPage = null;
        if (bodySlot && bodySlot.pageRef) {
          targetPage = bodySlot.pageRef;
        } else if (titleSlot && titleSlot.pageRef) {
          targetPage = titleSlot.pageRef;
        } else if (pages.length) {
          targetPage = pages[0];
        }

        var noteWarnings = (note.warnings || []).slice(0);
        if (result.oversetFlags && result.oversetFlags.length) {
          for (var ow = 0; ow < result.oversetFlags.length; ow++) {
            if (noteWarnings.indexOf(result.oversetFlags[ow]) === -1) {
              noteWarnings.push(result.oversetFlags[ow]);
            }
          }
        }

        var csvSlotUsed = "";
        var csvWidth = "";
        var csvHeight = "";
        var overflowChars = result.totalOverflow || 0;
        var bodyFrame = null;
        var titleFrame = null;

        if (bodySlot && layoutType !== "missing" && targetPage && targetPage.isValid) {
          csvSlotUsed = bodySlot.id;
          csvWidth = formatNumber(toCm(result.columnWidth || (bodySlot.bounds[3] - bodySlot.bounds[1])), 2);
          csvHeight = formatNumber(toCm(result.bodyHeight || (bodySlot.bounds[2] - bodySlot.bounds[0])), 2);
          if (layoutType !== "separate") {
            removeItemsByLabel(targetPage, note.noteId + "_titulo");
          }
          removeItemsByLabel(targetPage, note.noteId + "_texto");
          bodyFrame = createTextFrameForBounds(targetPage, bodySlot.bounds, note.noteId + "_texto");
          if (bodyFrame) {
            try {
              bodyFrame.textFramePreferences.textColumnCount = Math.max(1, slotColumnCount(bodySlot));
            } catch (_) {}
            var bodyText = note.bodyText || "";
            if (layoutType === "combined" && note.titleText) {
              bodyText = note.titleText + (bodyText ? "\r" + bodyText : "");
            }
            var bodyStory = writeStory(bodyFrame, bodyText);
            if (bodyStory && bodyStory.isValid) {
              applyStyleAndSize(bodyStory, config.styles.body.name, result.bodyPointSize, result.bodyPointSize, result.bodyPointSize);
              if (layoutType === "combined" && note.titleText) {
                applyStyleToFirstParagraph(bodyStory, config.styles.title.name, result.titlePointSize, result.titlePointSize, result.titlePointSize);
                if (bodyStory.paragraphs.length > 0) {
                  applyTitleSpan(bodyStory.paragraphs[0], result.titleSpan || 1);
                }
              }
            }
            var bodyInfo = bodyStory ? storyOversetInfo(bodyStory) : { over: false, exceed: 0 };
            if (bodyInfo.over) {
              overflowChars += bodyInfo.exceed;
              if (noteWarnings.indexOf("overset_hard") === -1) {
                noteWarnings.push("overset_hard");
              }
              markOversetFrame(bodyFrame);
            } else if (result.oversetFlags && (result.oversetFlags.indexOf("overset_hard") >= 0 || result.oversetFlags.indexOf("body_overset") >= 0)) {
              markOversetFrame(bodyFrame);
            } else {
              try { bodyFrame.strokeWeight = 0; } catch (_) {}
            }
          }
        } else if (!bodySlot) {
          if (noteWarnings.indexOf("slot_missing:" + note.noteId) === -1) {
            noteWarnings.push("slot_missing:" + note.noteId);
          }
        }

        if (layoutType === "separate" && titleSlot && titleSlot !== bodySlot) {
          var titlePage = titleSlot.pageRef || targetPage;
          if (titlePage && titlePage.isValid) {
            csvSlotUsed = csvSlotUsed ? csvSlotUsed + "+" + titleSlot.id : titleSlot.id;
            removeItemsByLabel(titlePage, note.noteId + "_titulo");
            titleFrame = createTextFrameForBounds(titlePage, titleSlot.bounds, note.noteId + "_titulo");
            if (titleFrame) {
              try {
                titleFrame.textFramePreferences.textColumnCount = Math.max(1, slotColumnCount(titleSlot));
              } catch (_) {}
              var titleStory = writeStory(titleFrame, note.titleText || "");
              if (titleStory && titleStory.isValid) {
                applyStyleAndSize(titleStory, config.styles.title.name, result.titlePointSize, result.titlePointSize, result.titlePointSize);
                if (titleStory.paragraphs.length > 0) {
                  applyTitleSpan(titleStory.paragraphs[0], result.titleSpan || 1);
                }
              }
              var titleInfo = titleStory ? storyOversetInfo(titleStory) : { over: false, exceed: 0 };
              if (titleInfo.over) {
                overflowChars += titleInfo.exceed;
                if (noteWarnings.indexOf("overset_hard") === -1) {
                  noteWarnings.push("overset_hard");
                }
                markOversetFrame(titleFrame);
              } else if (result.oversetFlags && result.oversetFlags.indexOf("title_overset") >= 0) {
                markOversetFrame(titleFrame);
              } else {
                try { titleFrame.strokeWeight = 0; } catch (_) {}
              }
            }
          }
        }

        if (layoutType !== "missing" && targetPage && targetPage.isValid) {
          placeNoteImages(note, targetPage, photoSlotsAll, config, noteWarnings);
        }

        csv.writeRow({
          pagina: targetPage && targetPage.isValid ? targetPage.name : note.pageLabel,
          nota: note.noteId,
          slot_usado: csvSlotUsed,
          ancho_col: csvWidth,
          alto_final: csvHeight,
          pt_titulo: result.titlePointSize || "",
          pt_cuerpo: result.bodyPointSize || "",
          overflow_chars: overflowChars,
          warnings: noteWarnings.join(";"),
          errors: ""
        });
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


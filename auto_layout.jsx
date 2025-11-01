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
      var columnSpan = slot.column_span || slot.col_span || slot.span || null;
      var columnCount = slot.text_column_count || slot.column_count || slot.columns || null;
      var columnWidth = null;
      var columnGutter = null;
      if (slot.column_width_pt || slot.col_width_pt) {
        columnWidth = slot.column_width_pt || slot.col_width_pt;
      } else if (slot.column_width_cm || slot.col_width_cm) {
        columnWidth = pointsFromCm(slot.column_width_cm || slot.col_width_cm);
      }
      if (slot.column_gutter_pt || slot.col_gutter_pt) {
        columnGutter = slot.column_gutter_pt || slot.col_gutter_pt;
      } else if (slot.column_gutter_cm || slot.col_gutter_cm) {
        columnGutter = pointsFromCm(slot.column_gutter_cm || slot.col_gutter_cm);
      }

      out[pageKey].push({
        id: slot.id ? slot.id : (pageKey + "_slot" + perPageCounter[pageKey]),
        page: slot.page,
        bounds: bounds,
        width: width,
        height: height,
        raw: slot,
        isPhoto: slot.is_photo_slot || slot.is_2col_photo_slot || false,
        is2ColPhotoSlot: slot.is_2col_photo_slot || false,
        fitsPhoto2ColHeight: slot.fits_photo_2col_h || false,
        columnSpan: columnSpan,
        columnCount: columnCount,
        columnWidth: columnWidth,
        columnGutter: columnGutter
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

  function measurementSpreadForSlot(slot, fallbackPage) {
    if (slot && slot.pageObj && slot.pageObj.isValid) {
      return slot.pageObj.parent;
    }
    if (fallbackPage && fallbackPage.isValid) {
      return fallbackPage.parent;
    }
    var doc = app.activeDocument;
    if (doc && doc.isValid && doc.spreads.length) {
      return doc.spreads[0];
    }
    return null;
  }

  function createPasteboardFrame(spread, width, height) {
    if (!spread || !spread.isValid) {
      return null;
    }
    try {
      var frame = spread.textFrames.add();
      var top = -10000;
      var left = -10000;
      var bottom = top + (height || 1000);
      var right = left + (width || 1000);
      frame.geometricBounds = [top, left, bottom, right];
      frame.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
      frame.textFramePreferences.autoSizingType = AutoSizingTypeEnum.OFF;
      frame.textFramePreferences.useMinimumHeightForAutoSizing = false;
      frame.textFramePreferences.verticalJustification = VerticalJustification.TOP_ALIGN;
      return frame;
    } catch (e) {
      $.writeln("createPasteboardFrame error: " + e);
      return null;
    }
  }

  function applySlotColumnsToFrame(frame, slot) {
    if (!frame || !frame.isValid || !slot) {
      return;
    }
    try {
      var prefs = frame.textFramePreferences;
      if (slot.columnCount) {
        prefs.textColumnCount = slot.columnCount;
      } else if (slot.raw && slot.raw.text_column_count) {
        prefs.textColumnCount = slot.raw.text_column_count;
      }
      if (slot.columnGutter !== null && slot.columnGutter !== undefined) {
        prefs.textColumnGutter = slot.columnGutter;
      } else if (slot.raw && slot.raw.text_column_gutter_pt) {
        prefs.textColumnGutter = slot.raw.text_column_gutter_pt;
      }
      if (slot.columnWidth !== null && slot.columnWidth !== undefined) {
        prefs.useFixedColumnWidth = true;
        prefs.textColumnFixedWidth = slot.columnWidth;
      } else if (slot.raw && slot.raw.text_column_width_pt) {
        prefs.useFixedColumnWidth = true;
        prefs.textColumnFixedWidth = slot.raw.text_column_width_pt;
      }
    } catch (e) {
      $.writeln("applySlotColumnsToFrame error: " + e);
    }
  }

  function findFrameByAutoSlotId(page, autoId) {
    if (!page || !page.isValid || !autoId) {
      return null;
    }
    var frames = page.textFrames;
    for (var i = 0; i < frames.length; i++) {
      var frame = frames[i];
      if (!frame || !frame.isValid) {
        continue;
      }
      try {
        if (frame.extractLabel && frame.extractLabel("auto_slot_id") === autoId) {
          return frame;
        }
      } catch (_) {}
    }
    return null;
  }

  function instantiateTextFrameForSlot(page, slot, bounds, label) {
    var frame = null;
    var targetBounds = bounds || (slot && slot.bounds ? cloneBounds(slot.bounds) : null);
    if (slot && slot.autoId) {
      frame = findFrameByAutoSlotId(page, slot.autoId);
      if (frame && frame.isValid) {
        try { frame.label = label || frame.label; } catch (_) {}
        if (targetBounds) {
          try { frame.geometricBounds = cloneBounds(targetBounds); } catch (_) {}
        }
        applySlotColumnsToFrame(frame, slot);
        clearStory(frame);
        try { frame.insertLabel("auto_slot_id", slot.autoId); } catch (_) {}
        return frame;
      }
    }
    removeItemsByLabel(page, label);
    frame = createTextFrameForBounds(page, targetBounds, label);
    if (frame && slot && slot.autoId) {
      try { frame.insertLabel("auto_slot_id", slot.autoId); } catch (_) {}
    }
    applySlotColumnsToFrame(frame, slot);
    return frame;
  }

  function hasAnySlots(slotsMap) {
    if (!slotsMap) {
      return false;
    }
    for (var key in slotsMap) {
      if (!slotsMap.hasOwnProperty(key)) {
        continue;
      }
      var arr = slotsMap[key];
      if (arr && arr.length) {
        return true;
      }
    }
    return false;
  }

  function frameIsCandidateSlot(frame) {
    if (!frame || !frame.isValid) {
      return false;
    }
    if (frame.locked || !frame.visible) {
      return false;
    }
    var gb = null;
    try {
      gb = frame.geometricBounds;
    } catch (e) {
      gb = null;
    }
    if (!gb) {
      return false;
    }
    var width = Math.abs(gb[3] - gb[1]);
    var height = Math.abs(gb[2] - gb[0]);
    var area = width * height;
    if (width < 60 || height < 60) {
      return false;
    }
    if (area < 10000) {
      return false;
    }
    if (gb[0] < -1000 || gb[1] < -1000) {
      return false;
    }
    return true;
  }

  function buildDocumentSlotMap(doc) {
    var map = {};
    if (!doc || !doc.isValid) {
      return map;
    }
    var pages = doc.pages;
    for (var i = 0; i < pages.length; i++) {
      var page = pages[i];
      if (!page || !page.isValid) {
        continue;
      }
      var frames = page.textFrames;
      var slots = [];
      var counter = 0;
      for (var j = 0; j < frames.length; j++) {
        var frame = frames[j];
        if (!frameIsCandidateSlot(frame)) {
          continue;
        }
        counter++;
        var gb = cloneBounds(frame.geometricBounds);
        var prefs = frame.textFramePreferences;
        var slotId = String(page.name) + "_auto_" + counter;
        var autoId = "auto_slot_" + (page.documentOffset + 1) + "_" + counter;
        var slot = {
          id: slotId,
          page: page.documentOffset + 1,
          bounds: gb,
          width: Math.abs(gb[3] - gb[1]),
          height: Math.abs(gb[2] - gb[0]),
          columnCount: prefs ? prefs.textColumnCount : null,
          columnGutter: prefs ? prefs.textColumnGutter : null,
          columnWidth: (prefs && prefs.useFixedColumnWidth) ? prefs.textColumnFixedWidth : null,
          isPhoto: false,
          source: "document_frame",
          autoId: autoId
        };
        if (!slot.columnWidth && prefs && prefs.textColumnCount) {
          try {
            slot.columnWidth = slot.width / Math.max(1, prefs.textColumnCount);
          } catch (_) {}
        }
        try {
          frame.insertLabel("auto_slot_id", autoId);
        } catch (_) {}
        slots.push(slot);
      }
      if (slots.length) {
        var nameKey = String(page.name);
        var indexKey = String(page.documentOffset + 1);
        map[nameKey] = slots.slice(0);
        map[indexKey] = slots.slice(0);
      }
    }
    return map;
  }

  function resolveTitleOverset(frame, cfg, stylesCfg, limits) {
    var result = {
      overset: "",
      exceed: 0,
      titlePointSize: currentStylePointSize(frame, stylesCfg.title.name)
    };
    if (!frame || !frame.isValid) {
      return result;
    }
    var story = frame.parentStory;
    var info = storyOversetInfo(story);
    if (!info.over) {
      return result;
    }
    var expanded = 0;
    while (info.over && expanded < cfg.max_expand_mm) {
      var step = Math.min(2, cfg.max_expand_mm - expanded);
      expandFrameHeightMM(frame, step, limits && limits.bodyMaxBottom !== undefined ? limits.bodyMaxBottom : null);
      expanded += step;
      info = storyOversetInfo(story);
    }
    if (info.over) {
      var dropped = 0;
      while (info.over && dropped < cfg.title_max_drop_pt) {
        var delta = -Math.min(cfg.body_step_pt, cfg.title_max_drop_pt - dropped);
        var applied = nudgeStyle(frame, stylesCfg.title.name, delta, stylesCfg.title.pt_min, stylesCfg.title.pt_max);
        if (applied === null) {
          break;
        }
        dropped += Math.abs(delta);
        info = storyOversetInfo(story);
      }
    }
    result.titlePointSize = currentStylePointSize(frame, stylesCfg.title.name);
    if (info.over) {
      result.overset = "overset_hard";
      result.exceed = info.exceed;
    }
    return result;
  }

  function slotsForPage(slotsMap, page) {
    if (!slotsMap || !page || !page.isValid) {
      return [];
    }
    function cloneWithPage(slots) {
      var out = [];
      for (var i = 0; i < slots.length; i++) {
        var original = slots[i];
        if (!original) {
          continue;
        }
        var copy = clone(original);
        copy.pageObj = page;
        out.push(copy);
      }
      return out;
    }
    var nameKey = String(page.name);
    if (slotsMap.hasOwnProperty(nameKey)) {
      return cloneWithPage(slotsMap[nameKey]);
    }
    var indexKey = String(page.documentOffset + 1);
    if (slotsMap.hasOwnProperty(indexKey)) {
      return cloneWithPage(slotsMap[indexKey]);
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

  function generateSlotPermutations(slots, noteCount, limit) {
    var combos = [];
    var choose = Math.min(noteCount, slots.length);
    if (choose <= 0) {
      combos.push([]);
      return combos;
    }

    var used = {};
    function step(prefix) {
      if (combos.length >= limit) {
        return;
      }
      if (prefix.length === choose) {
        var finalCombo = prefix.slice(0);
        while (finalCombo.length < noteCount) {
          finalCombo.push(-1);
        }
        combos.push(finalCombo);
        return;
      }
      for (var si = 0; si < slots.length; si++) {
        if (used[si]) {
          continue;
        }
        used[si] = true;
        prefix.push(si);
        step(prefix);
        prefix.pop();
        used[si] = false;
        if (combos.length >= limit) {
          break;
        }
      }
    }

    step([]);

    if (!combos.length) {
      combos.push([]);
    }
    return combos;
  }

  function LayoutSolver(pages, slots, config, stylesCfg) {
    if (pages && !(pages instanceof Array)) {
      pages = [pages];
    }
    this.pages = pages ? pages.slice(0) : [];
    this.page = this.pages.length ? this.pages[0] : null;
    this.slots = slots.slice(0);
    this.config = config;
    this.stylesCfg = stylesCfg;
    this.sizeProfiles = this.buildSizeProfiles();
    this.measurementCache = {};
    this.noteOptionCache = {};
    this.oversetCfg = this.config.overset || {};
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
          profiles.push({ body: bodyPt, title: titlePt });
          added[key] = true;
        }
      }
    }

    if (!profiles.length) {
      profiles.push({ body: body.pt_base, title: title.pt_base });
    }
    return profiles;
  };

  LayoutSolver.prototype.profileKey = function (profile) {
    if (!profile) {
      return "";
    }
    return [profile.body, profile.title].join("_");
  };

  LayoutSolver.prototype.measurementKey = function (note, slot, profile, type, span) {
    var parts = [];
    parts.push(note ? note.noteId : "");
    parts.push(slot ? slot.id : "");
    parts.push(this.profileKey(profile));
    parts.push(type || "");
    parts.push(span || 0);
    return parts.join("|");
  };

  LayoutSolver.prototype.cloneMeasurement = function (data) {
    if (!data) {
      return null;
    }
    var out = {};
    for (var key in data) {
      if (data.hasOwnProperty(key)) {
        out[key] = data[key];
      }
    }
    return out;
  };

  LayoutSolver.prototype.createMeasurementFrameForSlot = function (slot) {
    var spread = measurementSpreadForSlot(slot, this.page);
    if (!spread) {
      return null;
    }
    var height = slot && slot.height ? slot.height : 1000;
    var width = slot && slot.width ? slot.width : 1000;
    var frame = createPasteboardFrame(spread, width, height);
    if (frame && frame.isValid) {
      var top = -10000;
      var left = -10000;
      var bounds = [top, left, top + height, left + width];
      try {
        frame.geometricBounds = bounds;
      } catch (_) {}
      applySlotColumnsToFrame(frame, slot);
    }
    return frame;
  };

  LayoutSolver.prototype.measureCombined = function (note, slot, profile) {
    var key = this.measurementKey(note, slot, profile, "combined", 0);
    if (this.measurementCache.hasOwnProperty(key)) {
      return this.cloneMeasurement(this.measurementCache[key]);
    }
    var result = {
      type: "combined",
      slot: slot,
      overset: { overset: "overset_hard", exceed: 0 },
      bodyPointSize: profile.body,
      titlePointSize: profile.title,
      height: 0,
      success: false
    };
    var frame = this.createMeasurementFrameForSlot(slot);
    if (!frame || !frame.isValid) {
      result.overset.exceed = (note.bodyText ? note.bodyText.length : 0) + (note.titleText ? note.titleText.length : 0);
      this.measurementCache[key] = result;
      return this.cloneMeasurement(result);
    }
    var top = frame.geometricBounds[0];
    var limitBottom = top + (slot && slot.height ? slot.height : 0);
    var text = note.bodyText || "";
    if (note.titleText) {
      text = note.titleText + (text ? "\r" + text : "");
    }
    var story = writeStory(frame, text);
    if (story && story.isValid) {
      applyStyleAndSize(story, this.stylesCfg.body.name, profile.body, this.stylesCfg.body.pt_min, this.stylesCfg.body.pt_max);
      if (note.titleText) {
        applyStyleToFirstParagraph(story, this.stylesCfg.title.name, profile.title, this.stylesCfg.title.pt_min, this.stylesCfg.title.pt_max);
      }
    }
    var overset = resolveOverset(frame, null, this.oversetCfg, this.stylesCfg, { bodyMaxBottom: limitBottom });
    result.overset = overset;
    result.bodyPointSize = overset.bodyPointSize || profile.body;
    result.titlePointSize = overset.titlePointSize || profile.title;
    result.height = frame.geometricBounds[2] - frame.geometricBounds[0];
    result.success = overset.overset !== "overset_hard";
    try { frame.remove(); } catch (_) {}
    this.measurementCache[key] = result;
    return this.cloneMeasurement(result);
  };

  LayoutSolver.prototype.measureBody = function (note, slot, profile) {
    var key = this.measurementKey(note, slot, profile, "body", 0);
    if (this.measurementCache.hasOwnProperty(key)) {
      return this.cloneMeasurement(this.measurementCache[key]);
    }
    var result = {
      type: "body",
      slot: slot,
      overset: { overset: "", exceed: 0 },
      bodyPointSize: profile.body,
      height: 0,
      success: true
    };
    if (!note.bodyText) {
      this.measurementCache[key] = result;
      return this.cloneMeasurement(result);
    }
    var frame = this.createMeasurementFrameForSlot(slot);
    if (!frame || !frame.isValid) {
      result.overset = { overset: "overset_hard", exceed: note.bodyText.length };
      result.success = false;
      this.measurementCache[key] = result;
      return this.cloneMeasurement(result);
    }
    var top = frame.geometricBounds[0];
    var limitBottom = top + (slot && slot.height ? slot.height : 0);
    var story = writeStory(frame, note.bodyText || "");
    if (story && story.isValid) {
      applyStyleAndSize(story, this.stylesCfg.body.name, profile.body, this.stylesCfg.body.pt_min, this.stylesCfg.body.pt_max);
    }
    var overset = resolveOverset(frame, null, this.oversetCfg, this.stylesCfg, { bodyMaxBottom: limitBottom });
    result.overset = overset;
    result.bodyPointSize = overset.bodyPointSize || profile.body;
    result.height = frame.geometricBounds[2] - frame.geometricBounds[0];
    result.success = overset.overset !== "overset_hard";
    try { frame.remove(); } catch (_) {}
    this.measurementCache[key] = result;
    return this.cloneMeasurement(result);
  };

  LayoutSolver.prototype.measureTitle = function (note, slot, profile, span) {
    var spanCount = span || 1;
    var key = this.measurementKey(note, slot, profile, "title", spanCount);
    if (this.measurementCache.hasOwnProperty(key)) {
      return this.cloneMeasurement(this.measurementCache[key]);
    }
    var result = {
      type: "title",
      slot: slot,
      span: spanCount,
      overset: { overset: "", exceed: 0 },
      titlePointSize: profile.title,
      height: 0,
      success: true
    };
    if (!note.titleText) {
      this.measurementCache[key] = result;
      return this.cloneMeasurement(result);
    }
    var frame = this.createMeasurementFrameForSlot(slot);
    if (!frame || !frame.isValid) {
      result.overset = { overset: "overset_hard", exceed: note.titleText.length };
      result.success = false;
      this.measurementCache[key] = result;
      return this.cloneMeasurement(result);
    }
    var top = frame.geometricBounds[0];
    var limitBottom = top + (slot && slot.height ? slot.height : 0);
    try {
      var prefs = frame.textFramePreferences;
      var columnCount = slot && slot.columnCount ? slot.columnCount : null;
      if (!columnCount && slot && slot.raw && slot.raw.text_column_count) {
        columnCount = slot.raw.text_column_count;
      }
      if (!columnCount) {
        columnCount = Math.max(spanCount, 1);
      } else {
        columnCount = Math.max(columnCount, spanCount);
      }
      prefs.textColumnCount = columnCount;
    } catch (_) {}
    var story = writeStory(frame, note.titleText || "");
    if (story && story.isValid) {
      applyStyleAndSize(story, this.stylesCfg.title.name, profile.title, this.stylesCfg.title.pt_min, this.stylesCfg.title.pt_max);
      try {
        if (story.paragraphs.length > 0) {
          story.paragraphs[0].spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
          story.paragraphs[0].spanColumnCount = spanCount;
        }
      } catch (_) {}
    }
    var overset = resolveTitleOverset(frame, this.oversetCfg, this.stylesCfg, { bodyMaxBottom: limitBottom });
    result.overset = overset;
    result.titlePointSize = overset.titlePointSize || profile.title;
    result.height = frame.geometricBounds[2] - frame.geometricBounds[0];
    result.success = overset.overset !== "overset_hard";
    try { frame.remove(); } catch (_) {}
    this.measurementCache[key] = result;
    return this.cloneMeasurement(result);
  };

  LayoutSolver.prototype.resolveMaxSpanForSlot = function (slot) {
    if (!slot) {
      return 1;
    }
    var maxSpan = 1;
    if (slot.columnSpan) {
      maxSpan = Math.max(maxSpan, slot.columnSpan);
    }
    if (slot.columnCount) {
      maxSpan = Math.max(maxSpan, slot.columnCount);
    }
    if (slot.raw) {
      if (slot.raw.column_span) {
        maxSpan = Math.max(maxSpan, slot.raw.column_span);
      }
      if (slot.raw.col_span) {
        maxSpan = Math.max(maxSpan, slot.raw.col_span);
      }
      if (slot.raw.span) {
        maxSpan = Math.max(maxSpan, slot.raw.span);
      }
      if (slot.raw.text_column_count) {
        maxSpan = Math.max(maxSpan, slot.raw.text_column_count);
      }
    }
    if (maxSpan > 5) {
      maxSpan = 5;
    }
    if (maxSpan < 1) {
      maxSpan = 1;
    }
    return maxSpan;
  };

  LayoutSolver.prototype.generateOptionsForNote = function (note) {
    var cacheKey = note ? note.noteId : "";
    if (this.noteOptionCache.hasOwnProperty(cacheKey)) {
      return this.noteOptionCache[cacheKey].slice(0);
    }
    var options = [];
    if (!note) {
      this.noteOptionCache[cacheKey] = options;
      return options;
    }
    var hasBody = note.bodyText && note.bodyText.length;
    var hasTitle = note.titleText && note.titleText.length;
    var recommendedSpan = 1;
    if (hasTitle) {
      recommendedSpan = Math.min(5, Math.max(1, Math.round(note.titleText.length / 35)));
    }
    var MAX_BODY_CANDIDATES = 6;
    var MAX_TITLE_CANDIDATES = 6;

    var bodyCandidates = [];
    var titleCandidates = [];
    var combinedCandidates = [];

    for (var si = 0; si < this.slots.length; si++) {
      var slot = this.slots[si];
      if (!slot) {
        continue;
      }
      for (var pi = 0; pi < this.sizeProfiles.length; pi++) {
        var profile = this.sizeProfiles[pi];
        if (hasBody) {
          var bodyMeasurement = this.measureBody(note, slot, profile);
          if (bodyMeasurement) {
            bodyCandidates.push({
              slot: slot,
              profile: profile,
              measurement: bodyMeasurement,
              overflow: bodyMeasurement.overset ? bodyMeasurement.overset.exceed : 0,
              success: bodyMeasurement.overset ? bodyMeasurement.overset.overset !== "overset_hard" : true,
              area: slot.width * slot.height
            });
          }
        }
        if (hasTitle) {
          var maxSpan = this.resolveMaxSpanForSlot(slot);
          for (var span = 1; span <= maxSpan; span++) {
            var titleMeasurement = this.measureTitle(note, slot, profile, span);
            if (titleMeasurement) {
              titleCandidates.push({
                slot: slot,
                profile: profile,
                measurement: titleMeasurement,
                span: span,
                overflow: titleMeasurement.overset ? titleMeasurement.overset.exceed : 0,
                success: titleMeasurement.overset ? titleMeasurement.overset.overset !== "overset_hard" : true,
                spanDistance: Math.abs(span - recommendedSpan)
              });
            }
          }
        }
        if (hasBody) {
          var combinedMeasurement = this.measureCombined(note, slot, profile);
          if (combinedMeasurement) {
            combinedCandidates.push({
              slot: slot,
              profile: profile,
              measurement: combinedMeasurement,
              overflow: combinedMeasurement.overset ? combinedMeasurement.overset.exceed : 0,
              success: combinedMeasurement.overset ? combinedMeasurement.overset.overset !== "overset_hard" : true
            });
          }
        }
      }
    }

    bodyCandidates.sort(function (a, b) {
      if (a.success !== b.success) {
        return a.success ? -1 : 1;
      }
      if (a.overflow !== b.overflow) {
        return a.overflow - b.overflow;
      }
      return b.area - a.area;
    });
    titleCandidates.sort(function (a, b) {
      if (a.success !== b.success) {
        return a.success ? -1 : 1;
      }
      if (a.spanDistance !== b.spanDistance) {
        return a.spanDistance - b.spanDistance;
      }
      if (a.overflow !== b.overflow) {
        return a.overflow - b.overflow;
      }
      return a.span - b.span;
    });
    combinedCandidates.sort(function (a, b) {
      if (a.success !== b.success) {
        return a.success ? -1 : 1;
      }
      if (a.overflow !== b.overflow) {
        return a.overflow - b.overflow;
      }
      return b.slot.height - a.slot.height;
    });

    if (bodyCandidates.length > MAX_BODY_CANDIDATES) {
      bodyCandidates = bodyCandidates.slice(0, MAX_BODY_CANDIDATES);
    }
    if (titleCandidates.length > MAX_TITLE_CANDIDATES) {
      titleCandidates = titleCandidates.slice(0, MAX_TITLE_CANDIDATES);
    }

    for (var cc = 0; cc < combinedCandidates.length; cc++) {
      var comb = combinedCandidates[cc];
      var option = {
        note: note,
        profile: comb.profile,
        combined: true,
        body: comb.measurement,
        title: hasTitle ? comb.measurement : null,
        bodySlot: comb.slot,
        titleSlot: hasTitle ? comb.slot : null,
        overflow: comb.overflow,
        success: comb.success,
        warnings: [],
        slotsUsed: {}
      };
      option.slotsUsed[comb.slot.id] = true;
      options.push(option);
    }

    if (hasBody) {
      for (var bi = 0; bi < bodyCandidates.length; bi++) {
        var bodyCandidate = bodyCandidates[bi];
        if (hasTitle) {
          var anyTitle = false;
          for (var ti = 0; ti < titleCandidates.length; ti++) {
            var titleCandidate = titleCandidates[ti];
            if (titleCandidate.slot.id === bodyCandidate.slot.id) {
              continue;
            }
            anyTitle = true;
            var optionSplit = {
              note: note,
              profile: bodyCandidate.profile,
              combined: false,
              body: bodyCandidate.measurement,
              title: titleCandidate.measurement,
              bodySlot: bodyCandidate.slot,
              titleSlot: titleCandidate.slot,
              titleSpan: titleCandidate.span,
              overflow: (bodyCandidate.overflow || 0) + (titleCandidate.overflow || 0),
              success: bodyCandidate.success && titleCandidate.success,
              warnings: [],
              slotsUsed: {}
            };
            optionSplit.slotsUsed[bodyCandidate.slot.id] = true;
            optionSplit.slotsUsed[titleCandidate.slot.id] = true;
            options.push(optionSplit);
          }
          if (!anyTitle) {
            var warnOption = {
              note: note,
              profile: bodyCandidate.profile,
              combined: false,
              body: bodyCandidate.measurement,
              title: null,
              bodySlot: bodyCandidate.slot,
              titleSlot: null,
              overflow: (bodyCandidate.overflow || 0) + (note.titleText ? note.titleText.length : 0),
              success: false,
              warnings: ["title_slot_missing"],
              slotsUsed: {}
            };
            warnOption.slotsUsed[bodyCandidate.slot.id] = true;
            options.push(warnOption);
          }
        } else {
          var optionBodyOnly = {
            note: note,
            profile: bodyCandidate.profile,
            combined: false,
            body: bodyCandidate.measurement,
            title: null,
            bodySlot: bodyCandidate.slot,
            titleSlot: null,
            overflow: bodyCandidate.overflow || 0,
            success: bodyCandidate.success,
            warnings: [],
            slotsUsed: {}
          };
          optionBodyOnly.slotsUsed[bodyCandidate.slot.id] = true;
          options.push(optionBodyOnly);
        }
      }
    } else if (hasTitle) {
      for (var ti2 = 0; ti2 < titleCandidates.length; ti2++) {
        var onlyTitle = titleCandidates[ti2];
        var optionTitleOnly = {
          note: note,
          profile: onlyTitle.profile,
          combined: false,
          body: null,
          title: onlyTitle.measurement,
          bodySlot: null,
          titleSlot: onlyTitle.slot,
          titleSpan: onlyTitle.span,
          overflow: onlyTitle.overflow || 0,
          success: onlyTitle.success,
          warnings: [],
          slotsUsed: {}
        };
        optionTitleOnly.slotsUsed[onlyTitle.slot.id] = true;
        options.push(optionTitleOnly);
      }
    }

    if (!options.length) {
      options.push({
        note: note,
        profile: this.sizeProfiles.length ? this.sizeProfiles[0] : null,
        combined: false,
        body: null,
        title: null,
        bodySlot: null,
        titleSlot: null,
        overflow: (note.bodyText ? note.bodyText.length : 0) + (note.titleText ? note.titleText.length : 0),
        success: false,
        warnings: ["no_slot_available"],
        slotsUsed: {}
      });
    }

    for (var oi = 0; oi < options.length; oi++) {
      var opt = options[oi];
      opt.primarySlot = opt.bodySlot ? opt.bodySlot : (opt.titleSlot ? opt.titleSlot : null);
    }

    this.noteOptionCache[cacheKey] = options.slice(0);
    return options;
  };

  LayoutSolver.prototype.sortNoteOrder = function (notes) {
    var decorated = [];
    for (var i = 0; i < notes.length; i++) {
      var note = notes[i];
      var bodyChars = note && note.bodyText ? note.bodyText.length : 0;
      var titleChars = note && note.titleText ? note.titleText.length : 0;
      decorated.push({
        index: i,
        bodyChars: bodyChars,
        titleChars: titleChars,
        density: bodyChars + Math.round(titleChars / 2)
      });
    }
    decorated.sort(function (a, b) {
      if (a.density !== b.density) {
        return b.density - a.density;
      }
      if (a.titleChars !== b.titleChars) {
        return b.titleChars - a.titleChars;
      }
      return a.index - b.index;
    });
    var order = [];
    for (var di = 0; di < decorated.length; di++) {
      order.push(decorated[di].index);
    }
    return order;
  };

  LayoutSolver.prototype.buildResultFromAssignments = function (notes, assignments) {
    var result = {
      noteResults: [],
      totalOverflow: 0,
      success: true
    };
    for (var i = 0; i < notes.length; i++) {
      var note = notes[i];
      var option = assignments && assignments[i] ? assignments[i] : null;
      var noteOverflow = 0;
      var overset = { overset: "", exceed: 0 };
      var warnings = [];
      var bodySlot = null;
      var titleSlot = null;
      var bodyBounds = null;
      var titleBounds = null;
      var bodyPointSize = this.stylesCfg.body.pt_base;
      var titlePointSize = this.stylesCfg.title.pt_base;
      var titleSpan = option && option.titleSpan ? option.titleSpan : 1;
      var totalHeight = 0;
      if (option) {
        noteOverflow = option.overflow || 0;
        bodySlot = option.bodySlot || null;
        titleSlot = option.titleSlot || null;
        if (option.body) {
          overset = option.body.overset ? option.body.overset : overset;
          bodyPointSize = option.body.bodyPointSize || bodyPointSize;
          bodyBounds = bodySlot ? cloneBounds(bodySlot.bounds) : null;
          if (option.body.height) {
            totalHeight = option.body.height;
          }
        }
        if (option.title && option.title.overset) {
          if (option.body && option.body.overset && option.body.overset.overset === "overset_hard") {
            // keep body overset priority
          } else {
            overset = option.title.overset;
          }
          titlePointSize = option.title.titlePointSize || titlePointSize;
          titleBounds = titleSlot ? cloneBounds(titleSlot.bounds) : null;
        }
        if (option.combined && option.body) {
          titleBounds = bodyBounds ? cloneBounds(bodyBounds) : null;
          titlePointSize = option.body.titlePointSize || titlePointSize;
        }
        if (option.warnings && option.warnings.length) {
          for (var w = 0; w < option.warnings.length; w++) {
            warnings.push(option.warnings[w]);
          }
        }
        if (!option.success) {
          result.success = false;
        }
      } else {
        result.success = false;
        noteOverflow = (note.bodyText ? note.bodyText.length : 0) + (note.titleText ? note.titleText.length : 0);
        overset = { overset: "overset_hard", exceed: noteOverflow };
        warnings.push("solver_no_option");
      }
      result.totalOverflow += noteOverflow;
      result.noteResults.push({
        note: note,
        slot: bodySlot,
        bodySlot: bodySlot,
        titleSlot: titleSlot,
        bodyBounds: bodyBounds,
        titleBounds: titleBounds,
        overset: overset,
        warnings: warnings,
        bodyPointSize: bodyPointSize,
        titlePointSize: titlePointSize,
        titleSpan: titleSpan,
        columnWidth: bodySlot ? bodySlot.width : 0,
        totalHeight: totalHeight || (bodyBounds ? (bodyBounds[2] - bodyBounds[0]) : 0),
        combined: option ? option.combined : false,
        body: option ? option.body : null,
        title: option ? option.title : null
      });
    }
    return result;
  };

  LayoutSolver.prototype.solve = function (notes) {
    var result = {
      noteResults: [],
      success: false,
      totalOverflow: 0
    };
    if (!notes || !notes.length) {
      result.success = true;
      return result;
    }

    var noteOrder = this.sortNoteOrder(notes);
    var optionsPerNote = [];
    for (var i = 0; i < notes.length; i++) {
      var opts = this.generateOptionsForNote(notes[i]);
      opts.sort(function (a, b) {
        if (a.success !== b.success) {
          return a.success ? -1 : 1;
        }
        if ((a.overflow || 0) !== (b.overflow || 0)) {
          return (a.overflow || 0) - (b.overflow || 0);
        }
        var aSpan = a.titleSpan || (a.title && a.title.span) || 0;
        var bSpan = b.titleSpan || (b.title && b.title.span) || 0;
        return bSpan - aSpan;
      });
      optionsPerNote[i] = opts;
    }

    var maxAttempts = this.determineMaxAttempts(notes.length);
    var solverCfg = this.config.layout_solver || {};
    var open = [];
    var initial = {
      index: 0,
      assignments: new Array(notes.length),
      usedSlots: {},
      totalOverflow: 0,
      failureCount: 0,
      score: 0
    };
    open.push(initial);
    var bestState = null;
    var attempts = 0;

    while (open.length && attempts < maxAttempts) {
      open.sort(function (a, b) {
        return a.score - b.score;
      });
      var state = open.shift();
      if (!state) {
        break;
      }
      if (state.index >= noteOrder.length) {
        if (!bestState || state.totalOverflow < bestState.totalOverflow || (state.totalOverflow === bestState.totalOverflow && state.failureCount < bestState.failureCount)) {
          bestState = state;
        }
        if (state.failureCount === 0 && solverCfg.strategy === "best_first") {
          bestState = state;
          break;
        }
        attempts++;
        continue;
      }

      var noteIdx = noteOrder[state.index];
      var noteOptions = optionsPerNote[noteIdx];
      var note = notes[noteIdx];
      var expanded = false;

      if (!noteOptions || !noteOptions.length) {
        noteOptions = [];
      }

      for (var oi = 0; oi < noteOptions.length; oi++) {
        var option = noteOptions[oi];
        var conflict = false;
        for (var sid in option.slotsUsed) {
          if (option.slotsUsed.hasOwnProperty(sid) && state.usedSlots[sid]) {
            conflict = true;
            break;
          }
        }
        if (conflict) {
          continue;
        }
        expanded = true;
        var nextState = {
          index: state.index + 1,
          assignments: state.assignments.slice(0),
          usedSlots: {},
          totalOverflow: state.totalOverflow + (option.overflow || 0),
          failureCount: state.failureCount + (option.success ? 0 : 1)
        };
        for (var used in state.usedSlots) {
          if (state.usedSlots.hasOwnProperty(used)) {
            nextState.usedSlots[used] = true;
          }
        }
        for (var usedNew in option.slotsUsed) {
          if (option.slotsUsed.hasOwnProperty(usedNew)) {
            nextState.usedSlots[usedNew] = true;
          }
        }
        nextState.assignments[noteIdx] = option;
        var remaining = noteOrder.length - nextState.index;
        nextState.score = nextState.totalOverflow + nextState.failureCount * 1000 + remaining * 10;
        open.push(nextState);
      }

      if (!expanded) {
        var fallbackOverflow = (note && note.bodyText ? note.bodyText.length : 0) + (note && note.titleText ? note.titleText.length : 0);
        var fallbackState = {
          index: state.index + 1,
          assignments: state.assignments.slice(0),
          usedSlots: {},
          totalOverflow: state.totalOverflow + fallbackOverflow,
          failureCount: state.failureCount + 1
        };
        for (var usedPrev in state.usedSlots) {
          if (state.usedSlots.hasOwnProperty(usedPrev)) {
            fallbackState.usedSlots[usedPrev] = true;
          }
        }
        fallbackState.assignments[noteIdx] = null;
        var remainingFallback = noteOrder.length - fallbackState.index;
        fallbackState.score = fallbackState.totalOverflow + fallbackState.failureCount * 1000 + remainingFallback * 10;
        open.push(fallbackState);
      }

      attempts++;
    }

    if (!bestState) {
      for (var oi2 = 0; oi2 < open.length; oi2++) {
        var candidate = open[oi2];
        if (!candidate) {
          continue;
        }
        if (!bestState || candidate.totalOverflow < bestState.totalOverflow || (candidate.totalOverflow === bestState.totalOverflow && candidate.failureCount < bestState.failureCount)) {
          bestState = candidate;
        }
      }
    }

    if (!bestState) {
      bestState = initial;
    }

    var finalResult = this.buildResultFromAssignments(notes, bestState.assignments);
    finalResult.totalOverflow = bestState.totalOverflow || 0;
    finalResult.success = bestState.failureCount === 0 && finalResult.success;
    return finalResult;
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

  function placeNoteImages(note, page, photoSlots, config, warnings) {
    if (!note || !note.images || !note.images.length) {
      return;
    }
    var solverPhotoCfg = config.layout_solver && config.layout_solver.photo ? config.layout_solver.photo : {};
    for (var i = 0; i < note.images.length; i++) {
      var photo = note.images[i];
      var slot = null;
      for (var si = 0; si < photoSlots.length; si++) {
        var candidate = photoSlots[si];
        if (candidate.used) {
          continue;
        }
        if (solverPhotoCfg.slot_strict) {
          if (!candidate.is2ColPhotoSlot || !candidate.fitsPhoto2ColHeight) {
            continue;
          }
        }
        slot = candidate;
        photoSlots[si].used = true;
        break;
      }
      if (!slot) {
        if (warnings) {
          warnings.push("photo_slot_missing:" + note.noteId);
        }
        continue;
      }
      if (!solverPhotoCfg.slot_strict && !slot.isPhoto) {
        if (warnings) {
          warnings.push("photo_slot_nonphoto:" + slot.id);
        }
      }
      var bounds = cloneBounds(slot.bounds);
      removeItemsByLabel(page, note.noteId + "_foto" + photo.index);
      var frame = page.rectangles.add();
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
    if (!hasAnySlots(layoutSlots)) {
      layoutSlots = buildDocumentSlotMap(doc);
    }
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

      var remainingIndex = 0;
      var solverDisabled = config.layout_solver && config.layout_solver.enabled === false;
      var batches = [];

      if (pages.length === 2 && !solverDisabled) {
        var combinedTextSlots = [];
        var combinedPhotoSlots = [];
        for (var sp = 0; sp < pages.length; sp++) {
          var spreadPage = pages[sp];
          var spreadSlots = splitSlotsByType(slotsForPage(layoutSlots, spreadPage));
          var spreadText = spreadSlots.text;
          var spreadPhoto = spreadSlots.photo;
          for (var st = 0; st < spreadText.length; st++) {
            combinedTextSlots.push(spreadText[st]);
          }
          for (var pp = 0; pp < spreadPhoto.length; pp++) {
            spreadPhoto[pp].used = false;
            combinedPhotoSlots.push(spreadPhoto[pp]);
          }
        }
        if (combinedTextSlots.length) {
          var notesForSpread = [];
          while (remainingIndex < notesData.length && notesForSpread.length < combinedTextSlots.length) {
            notesForSpread.push(notesData[remainingIndex]);
            remainingIndex++;
          }
          if (notesForSpread.length) {
            var solverSpread = new LayoutSolver(pages, combinedTextSlots, config, config.styles);
            var solutionSpread = solverSpread.solve(notesForSpread);
            batches.push({
              notes: notesForSpread,
              results: solutionSpread.noteResults || [],
              photoSlots: combinedPhotoSlots,
              defaultPage: pages[0]
            });
          }
        }
      } else {
        for (var pg = 0; pg < pages.length; pg++) {
          if (remainingIndex >= notesData.length) {
            break;
          }
          var page = pages[pg];
          var slots = splitSlotsByType(slotsForPage(layoutSlots, page));
          var textSlots = slots.text;
          var photoSlots = slots.photo;
          for (var ps = 0; ps < photoSlots.length; ps++) {
            photoSlots[ps].used = false;
          }

          if (!textSlots.length) {
            continue;
          }

          var notesForPage = [];
          while (remainingIndex < notesData.length && notesForPage.length < textSlots.length) {
            notesForPage.push(notesData[remainingIndex]);
            remainingIndex++;
          }

          if (!notesForPage.length) {
            continue;
          }

          var results = [];
          if (solverDisabled) {
            for (var nf = 0; nf < notesForPage.length; nf++) {
              var manualSlot = nf < textSlots.length ? textSlots[nf] : null;
              results.push({
                note: notesForPage[nf],
                slot: manualSlot,
                overset: { overset: manualSlot ? "" : "overset_hard", exceed: 0 },
                warnings: [],
                bodyPointSize: config.styles.body.pt_base,
                titlePointSize: config.styles.title.pt_base,
                columnWidth: manualSlot ? manualSlot.width : 0,
                totalHeight: manualSlot ? (manualSlot.bounds[2] - manualSlot.bounds[0]) : 0,
                bodyBounds: manualSlot ? cloneBounds(manualSlot.bounds) : null,
              });
            }
          } else {
            var solver = new LayoutSolver(page, textSlots, config, config.styles);
            var solution = solver.solve(notesForPage);
            results = solution.noteResults || [];
          }

          batches.push({
            notes: notesForPage,
            results: results,
            photoSlots: photoSlots,
            defaultPage: page
          });
        }
      }

      for (var bi = 0; bi < batches.length; bi++) {
        var batch = batches[bi];
        var notesForPage = batch.notes || [];
        if (!notesForPage.length) {
          continue;
        }
        var results = batch.results || [];
        var photoSlots = batch.photoSlots || [];
        var page = batch.defaultPage || (pages.length ? pages[0] : null);
        for (var psReset = 0; psReset < photoSlots.length; psReset++) {
          photoSlots[psReset].used = false;
        }

        for (var nr = 0; nr < notesForPage.length; nr++) {
          var note = notesForPage[nr];
          var result = results[nr] || {
            note: note,
            bodySlot: null,
            titleSlot: null,
            overset: { overset: "overset_hard", exceed: (note.bodyText ? note.bodyText.length : 0) + (note.titleText ? note.titleText.length : 0) },
            warnings: ["solver_no_result"],
            bodyPointSize: config.styles.body.pt_base,
            titlePointSize: config.styles.title.pt_base,
            titleSpan: 1,
            bodyBounds: null,
            titleBounds: null,
            combined: false
          };

          var bodySlot = result.bodySlot || result.slot || null;
          var titleSlot = result.titleSlot || null;
          var bodyPage = bodySlot && bodySlot.pageObj ? bodySlot.pageObj : page;
          var titlePage = titleSlot && titleSlot.pageObj ? titleSlot.pageObj : bodyPage;
          var csvWidth = "";
          var csvHeight = "";
          var overflowChars = 0;
          var noteWarnings = (note.warnings || []).slice(0);
          if (result.warnings && result.warnings.length) {
            for (var w = 0; w < result.warnings.length; w++) {
              noteWarnings.push(result.warnings[w]);
            }
          }

          if (bodySlot) {
            csvWidth = formatNumber(toCm(result.columnWidth || (bodySlot.bounds[3] - bodySlot.bounds[1])), 2);
            csvHeight = formatNumber(toCm(result.totalHeight || (bodySlot.bounds[2] - bodySlot.bounds[0])), 2);
            var bodyLabel = note.noteId + "_texto";
            var bodyFrame = instantiateTextFrameForSlot(bodyPage, bodySlot, result.bodyBounds || (bodySlot ? bodySlot.bounds : null), bodyLabel);
            if (bodyFrame) {
              var bodyStoryText = note.bodyText || "";
              if (result.combined) {
                if (note.titleText) {
                  bodyStoryText = note.titleText + (bodyStoryText ? "\r" + bodyStoryText : "");
                }
              }
              var bodyStory = writeStory(bodyFrame, bodyStoryText);
              if (bodyStory && bodyStory.isValid) {
                applyStyleAndSize(bodyStory, config.styles.body.name, result.bodyPointSize, result.bodyPointSize, result.bodyPointSize);
                if (result.combined && note.titleText) {
                  applyStyleToFirstParagraph(bodyStory, config.styles.title.name, result.titlePointSize, result.titlePointSize, result.titlePointSize);
                }
              }
              var bodyInfo = storyOversetInfo(bodyStory);
              if (bodyInfo.over) {
                overflowChars += bodyInfo.exceed;
                noteWarnings.push("overset_hard_body");
                markOversetFrame(bodyFrame);
              } else if (result.body && result.body.overset && result.body.overset.overset === "overset_hard") {
                overflowChars += result.body.overset.exceed || 0;
                noteWarnings.push("overset_hard_body");
                markOversetFrame(bodyFrame);
              } else {
                try { bodyFrame.strokeWeight = 0; } catch (_) {}
              }
            }
          }

          if (!result.combined && titleSlot && note.titleText) {
            var titleFrame = instantiateTextFrameForSlot(titlePage, titleSlot, result.titleBounds || (titleSlot ? titleSlot.bounds : null), note.noteId + "_titulo");
            if (titleFrame) {
              var titleStory = writeStory(titleFrame, note.titleText || "");
              if (titleStory && titleStory.isValid) {
                applyStyleAndSize(titleStory, config.styles.title.name, result.titlePointSize, result.titlePointSize, result.titlePointSize);
                try {
                  if (titleStory.paragraphs.length > 0) {
                    titleStory.paragraphs[0].spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
                    titleStory.paragraphs[0].spanColumnCount = result.titleSpan || 1;
                  }
                } catch (_) {}
              }
              var titleInfo = storyOversetInfo(titleStory);
              if (titleInfo.over) {
                overflowChars += titleInfo.exceed;
                noteWarnings.push("overset_hard_title");
                markOversetFrame(titleFrame);
              } else if (result.title && result.title.overset && result.title.overset.overset === "overset_hard") {
                overflowChars += result.title.overset.exceed || 0;
                noteWarnings.push("overset_hard_title");
                markOversetFrame(titleFrame);
              } else {
                try { titleFrame.strokeWeight = 0; } catch (_) {}
              }
            }
          }

          if (!bodySlot && !titleSlot) {
            noteWarnings.push("slot_missing:" + note.noteId);
            overflowChars = result.overset ? result.overset.exceed : overflowChars;
          }

          placeNoteImages(note, bodyPage, photoSlots, config, noteWarnings);

          var slotIdForCsv = bodySlot ? bodySlot.id : (titleSlot ? titleSlot.id : "");
          if (!csvWidth && titleSlot) {
            csvWidth = formatNumber(toCm(titleSlot.bounds[3] - titleSlot.bounds[1]), 2);
          }
          if (!csvHeight && titleSlot) {
            csvHeight = formatNumber(toCm(titleSlot.bounds[2] - titleSlot.bounds[0]), 2);
          }

          csv.writeRow({
            pagina: bodyPage ? bodyPage.name : page.name,
            nota: note.noteId,
            slot_usado: slotIdForCsv,
            ancho_col: csvWidth,
            alto_final: csvHeight,
            pt_titulo: result.titlePointSize || "",
            pt_cuerpo: result.bodyPointSize || "",
            overflow_chars: overflowChars,
            warnings: noteWarnings.join(";"),
            errors: "",
          });
        }
      }

      while (remainingIndex < notesData.length) {
        var leftover = notesData[remainingIndex];
        remainingIndex++;
        var leftoverWarnings = (leftover.warnings || []).slice(0);
        leftoverWarnings.push("no_slot_available");
        csv.writeRow({
          pagina: leftover.pageLabel,
          nota: leftover.noteId,
          slot_usado: "",
          ancho_col: "",
          alto_final: "",
          pt_titulo: "",
          pt_cuerpo: "",
          overflow_chars: leftover.bodyText ? leftover.bodyText.length : 0,
          warnings: leftoverWarnings.join(";"),
          errors: "",
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


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

  function LayoutSolver(page, slots, config, stylesCfg) {
    this.page = page;
    this.slots = slots.slice(0);
    this.config = config;
    this.stylesCfg = stylesCfg;
    this.sizeProfiles = this.buildSizeProfiles();
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

  LayoutSolver.prototype.evaluateAssignment = function (notes, assignment, profile) {
    var attempt = {
      noteResults: [],
      totalOverflow: 0,
      success: true
    };
    var createdFrames = [];
    var oversetCfg = this.config.overset || {};

    for (var ni = 0; ni < notes.length; ni++) {
      var note = notes[ni];
      var slotIndex = assignment[ni];
      if (slotIndex === undefined || slotIndex === null) {
        slotIndex = -1;
      }
      var slot = slotIndex >= 0 && slotIndex < this.slots.length ? this.slots[slotIndex] : null;
      if (!slot) {
        attempt.success = false;
        attempt.noteResults.push({
          note: note,
          slot: null,
          overset: { overset: "overset_hard", exceed: note.bodyText ? note.bodyText.length : 0 },
          warnings: ["slot_missing"],
          bodyPointSize: profile.body,
          titlePointSize: profile.title,
          bodyBounds: null,
          columnWidth: 0,
          totalHeight: 0
        });
        continue;
      }

      var frame = this.page.textFrames.add();
      frame.geometricBounds = [slot.bounds[0], slot.bounds[1], slot.bounds[2], slot.bounds[3]];
      frame.label = note.noteId + "_texto_temp";
      createdFrames.push(frame);

      var fullText = note.bodyText || "";
      if (note.titleText) {
        if (fullText) {
          fullText = note.titleText + "\r" + fullText;
        } else {
          fullText = note.titleText;
        }
      }

      var story = writeStory(frame, fullText);
      if (story && story.isValid) {
        applyStyleAndSize(story, this.stylesCfg.body.name, profile.body, this.stylesCfg.body.pt_min, this.stylesCfg.body.pt_max);
        if (note.titleText) {
          applyStyleToFirstParagraph(story, this.stylesCfg.title.name, profile.title, this.stylesCfg.title.pt_min, this.stylesCfg.title.pt_max);
        }
      }

      var overset = resolveOverset(frame, null, oversetCfg, this.stylesCfg, { bodyMaxBottom: slot.bounds[2] });
      attempt.totalOverflow += overset.exceed;
      if (overset.overset === "overset_hard") {
        attempt.success = false;
      }

      attempt.noteResults.push({
        note: note,
        slot: slot,
        overset: overset,
        warnings: [],
        bodyPointSize: overset.bodyPointSize || profile.body,
        titlePointSize: overset.titlePointSize || profile.title,
        bodyBounds: frame.geometricBounds.slice(0),
        columnWidth: slot.width,
        totalHeight: frame.geometricBounds[2] - frame.geometricBounds[0]
      });
    }

    for (var ci = createdFrames.length - 1; ci >= 0; ci--) {
      try {
        if (createdFrames[ci] && createdFrames[ci].isValid) {
          createdFrames[ci].remove();
        }
      } catch (_) {}
    }

    return attempt;
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

    var maxAttempts = this.determineMaxAttempts(notes.length);
    var combos = generateSlotPermutations(this.slots, notes.length, maxAttempts);
    var profiles = this.sizeProfiles;
    var solverCfg = this.config.layout_solver || {};
    var best = null;
    var attempts = 0;

    for (var ci = 0; ci < combos.length && attempts < maxAttempts; ci++) {
      for (var pi = 0; pi < profiles.length && attempts < maxAttempts; pi++) {
        var attempt = this.evaluateAssignment(notes, combos[ci], profiles[pi]);
        attempts++;
        if (!best) {
          best = attempt;
        } else {
          if (attempt.success && !best.success) {
            best = attempt;
          } else if (attempt.success === best.success) {
            if (attempt.totalOverflow < best.totalOverflow) {
              best = attempt;
            }
          }
        }
        if (attempt.success && solverCfg.strategy === "best_first") {
          best = attempt;
          ci = combos.length;
          break;
        }
      }
    }

    if (!best) {
      best = {
        noteResults: [],
        success: false,
        totalOverflow: 0
      };
    }

    return best;
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
        if (!photoSlots[si].used) {
          slot = photoSlots[si];
          photoSlots[si].used = true;
          break;
        }
      }
      if (!slot) {
        if (warnings) {
          warnings.push("photo_slot_missing:" + note.noteId);
        }
        continue;
      }
      if (solverPhotoCfg.slot_strict && !slot.isPhoto) {
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
        if (config.layout_solver && config.layout_solver.enabled === false) {
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

        for (var nr = 0; nr < notesForPage.length; nr++) {
          var note = notesForPage[nr];
          var result = results[nr] || {
            note: note,
            slot: null,
            overset: { overset: "overset_hard", exceed: note.bodyText ? note.bodyText.length : 0 },
            warnings: ["solver_no_result"],
            bodyPointSize: config.styles.body.pt_base,
            titlePointSize: config.styles.title.pt_base,
            columnWidth: 0,
            totalHeight: 0,
            bodyBounds: null,
          };

          var slot = result.slot;
          var csvWidth = "";
          var csvHeight = "";
          var overflowChars = result.overset ? result.overset.exceed : 0;
          var noteWarnings = (note.warnings || []).slice(0);
          if (result.warnings && result.warnings.length) {
            for (var w = 0; w < result.warnings.length; w++) {
              noteWarnings.push(result.warnings[w]);
            }
          }

          if (slot) {
            csvWidth = formatNumber(toCm(result.columnWidth || (slot.bounds[3] - slot.bounds[1])), 2);
            csvHeight = formatNumber(toCm(result.totalHeight || (slot.bounds[2] - slot.bounds[0])), 2);
            removeItemsByLabel(page, note.noteId + "_texto");
            var finalFrame = createTextFrameForBounds(page, result.bodyBounds || slot.bounds, note.noteId + "_texto");
            if (finalFrame) {
              var fullText = note.bodyText || "";
              if (note.titleText) {
                fullText = note.titleText + (fullText ? "\r" + fullText : "");
              }
              var story = writeStory(finalFrame, fullText);
              if (story && story.isValid) {
                applyStyleAndSize(story, config.styles.body.name, result.bodyPointSize, result.bodyPointSize, result.bodyPointSize);
                if (note.titleText) {
                  applyStyleToFirstParagraph(story, config.styles.title.name, result.titlePointSize, result.titlePointSize, result.titlePointSize);
                }
              }
              var info = storyOversetInfo(story);
              if (info.over) {
                overflowChars = info.exceed;
                noteWarnings.push("overset_hard");
                markOversetFrame(finalFrame);
              } else if (result.overset && result.overset.overset === "overset_hard") {
                noteWarnings.push("overset_hard");
                markOversetFrame(finalFrame);
              } else {
                try { finalFrame.strokeWeight = 0; } catch (_) {}
              }
            }
            placeNoteImages(note, page, photoSlots, config, noteWarnings);
          } else {
            noteWarnings.push("slot_missing:" + note.noteId);
          }

          csv.writeRow({
            pagina: page.name,
            nota: note.noteId,
            slot_usado: slot ? slot.id : "",
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


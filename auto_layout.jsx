#target "indesign"

(function(){
  var DEFAULT_CONFIG = {
    styles: {
      body: {name: "Textos", pt_base: 9.5, pt_min: 9.0, pt_max: 10.0},
      title: {name: "Titulo 1", pt_base: 25, pt_min: 24, pt_max: 26}
    },
    images: {
      default_object_style: "Foto_2col_1cm",
      fallback_object_style: "",
      allow_photo_slot_release: false,
      distribute_unmatched_images: false
    },
    layout: {
      columns: 5,
      gutter_mm: 4.25,
      photo_default_width_columns: 2,
      photo_default_height_mm: 10
    },
    pages: {
      spreads: {},
      detect_vdr_by_name: true
    },
    overset: {
      max_expand_mm: 12,
      body_step_pt: 0.25,
      body_max_drop_pt: 0.5,
      title_max_drop_pt: 0.5
    },
    pdf: {
      preset_name: "Smallest File Size"
    },
    llm: {
      enabled: false,
      endpoint: "",
      api_key_env: ""
    },
    importPrefs: {}
  };

  function cloneConfig(obj){
    if (obj === null || typeof obj !== "object"){
      return obj;
    }
    var copy = obj instanceof Array ? [] : {};
    for (var key in obj){
      if (obj.hasOwnProperty(key)){
        copy[key] = cloneConfig(obj[key]);
      }
    }
    return copy;
  }

  function deepMerge(target, source){
    for (var key in source){
      if (!source.hasOwnProperty(key)) continue;
      var value = source[key];
      if (value && typeof value === "object" && !(value instanceof File) && !(value instanceof Folder)){
        if (!target[key]){
          target[key] = value instanceof Array ? [] : {};
        }
        deepMerge(target[key], value);
      } else {
        target[key] = value;
      }
    }
    return target;
  }

  function readConfig(file){
    var cfg = cloneConfig(DEFAULT_CONFIG);
    if (!file || !file.exists){
      return cfg;
    }
    try{
      file.open("r");
      var content = file.read();
      file.close();
      if (content && content.length > 0){
        var parsed = JSON.parse(content);
        cfg = deepMerge(cfg, parsed);
      }
    }catch(e){
      try{ file.close(); }catch(e2){}
      $.writeln("Config read error: " + e);
    }
    return cfg;
  }

  function selectRootFolder(){
    return Folder.selectDialog("Seleccionar carpeta raiz de la edicion");
  }

  function extractNumericTokens(name){
    if (!name){
      return [];
    }
    var matches = name.match(/\d+/g);
    return matches ? matches : [];
  }

  function firstNumericValue(name){
    var tokens = extractNumericTokens(name);
    if (tokens.length === 0){
      return null;
    }
    var value = parseInt(tokens[0], 10);
    return isNaN(value) ? null : value;
  }

  function listPageFolders(root){
    if (!root || !root.exists){
      return [];
    }
    var folders = root.getFiles(function(f){
      if (!(f instanceof Folder)){
        return false;
      }
      return /^\d+/.test(f.name);
    });
    folders.sort(function(a,b){
      var na = firstNumericValue(a.name);
      var nb = firstNumericValue(b.name);
      if (na !== null && nb !== null && na !== nb){
        return na - nb;
      }
      var sa = a.name;
      var sb = b.name;
      if (sa < sb) return -1;
      if (sa > sb) return 1;
      return 0;
    });
    return folders;
  }

  function determinePageNames(folderName, cfg){
    var pagesCfg = cfg.pages || {};
    var spreads = pagesCfg.spreads || {};
    var detectVdr = pagesCfg.detect_vdr_by_name === undefined ? true : !!pagesCfg.detect_vdr_by_name;
    var tokens = extractNumericTokens(folderName);
    var lowerName = (folderName || "").toLowerCase();
    var key = tokens.length > 0 ? tokens[0] : folderName;
    if (spreads[folderName]){
      return spreads[folderName];
    }
    if (spreads[key]){
      return spreads[key];
    }
    if (tokens.length > 1){
      return tokens;
    }
    if (detectVdr && /vdr/.test(lowerName)){
      if (spreads[key]){
        return spreads[key];
      }
    }
    if (tokens.length === 1){
      return [tokens[0]];
    }
    return [folderName];
  }

  function choosePageNameForNote(note, availablePages){
    if (!availablePages || availablePages.length === 0){
      return null;
    }
    if (note && note.pageName){
      var target = String(note.pageName);
      for (var i=0; i<availablePages.length; i++){
        if (availablePages[i] === target){
          return availablePages[i];
        }
      }
      var tokens = extractNumericTokens(target);
      if (tokens.length > 0){
        var token = tokens[0];
        for (var j=0; j<availablePages.length; j++){
          var candidate = availablePages[j];
          var candidateTokens = extractNumericTokens(candidate);
          if (candidateTokens.length > 0 && candidateTokens[0] === token){
            return candidate;
          }
        }
      }
    }
    return availablePages[0];
  }

  function pushUnique(arr, value){
    if (!value){
      return;
    }
    if (!arr){
      return;
    }
    for (var i=0; i<arr.length; i++){
      if (arr[i] === value){
        return;
      }
    }
    arr.push(value);
  }

  function isHardWarning(warning){
    if (!warning){
      return false;
    }
    return warning === "overset_hard" || warning === "extra_images" || warning === "extra_notes_unplaced";
  }

  function findPageItemByLabel(page, label){
    if (!page){
      return null;
    }
    var items = page.allPageItems;
    for (var i=0; i<items.length; i++){
      if (items[i].label && items[i].label === label){
        return items[i];
      }
    }
    return null;
  }

  function applyParagraphStyleToStory(story, styleName){
    if (!story || !styleName){
      return;
    }
    try{
      var doc = app.activeDocument;
      if (!doc){
        return;
      }
      var style = doc.paragraphStyles.itemByName(styleName);
      if (!style || !style.isValid){
        return;
      }
      var paragraphs = story.paragraphs;
      for (var i=0; i<paragraphs.length; i++){
        paragraphs[i].appliedParagraphStyle = style;
      }
      story.recompose();
    }catch(e){
      $.writeln("applyParagraphStyleToStory error: " + e);
    }
  }

  function duplicateBasePage(doc){
    if (!doc || doc.pages.length === 0){
      return null;
    }
    var basePage = doc.pages[0];
    var dup = basePage.duplicate(LocationOptions.AT_END);
    return dup;
  }

  function resetTextFrame(textFrame){
    if (!textFrame){
      return;
    }
    try{
      textFrame.contents = "";
      var story = textFrame.parentStory;
      if (story){
        story.clearOverrides(OverrideType.ALL);
      }
    }catch(e){
      $.writeln("resetTextFrame error: " + e);
    }
  }

  function placeWordIntoTextFrame(textFrame, file, importPrefs){
    if (!textFrame || !file || !file.exists){
      return null;
    }
    resetTextFrame(textFrame);
    var previousPrefs = {};
    var wordPrefs = app.wordImportPreferences;
    try{
      for (var pref in importPrefs){
        if (!importPrefs.hasOwnProperty(pref)) continue;
        try{
          previousPrefs[pref] = wordPrefs[pref];
          wordPrefs[pref] = importPrefs[pref];
        }catch(inner){
          $.writeln("word import prefs error: " + inner);
        }
      }
    }catch(e){
      $.writeln("word import prefs loop error: " + e);
    }
    try{
      textFrame.place(file);
      var story = textFrame.parentStory;
      if (story){
        story.recompose();
      }
      return story;
    }catch(e2){
      $.writeln("placeWordIntoTextFrame error: " + e2);
      return null;
    } finally {
      try{
        for (var key2 in previousPrefs){
          if (previousPrefs.hasOwnProperty(key2)){
            wordPrefs[key2] = previousPrefs[key2];
          }
        }
      }catch(e3){
        $.writeln("restore word prefs error: " + e3);
      }
    }
  }

  function applyObjectStyleIfExists(rect, styleName){
    if (!rect || !styleName || styleName === ""){
      return false;
    }
    try{
      var doc = app.activeDocument;
      if (doc){
        var style = doc.objectStyles.itemByName(styleName);
        if (style && style.isValid){
          rect.appliedObjectStyle = style;
          return true;
        }
      }
    }catch(e){
      $.writeln("applyObjectStyleIfExists error: " + e);
    }
    return false;
  }

  function placeImageIntoRect(rect, file, objectStyleName, fallbackStyleName){
    if (!rect || !file || !file.exists){
      return;
    }
    try{
      var applied = applyObjectStyleIfExists(rect, objectStyleName);
      if (!applied && fallbackStyleName){
        applyObjectStyleIfExists(rect, fallbackStyleName);
      }
      rect.place(file);
      rect.fit(FitOptions.PROPORTIONALLY);
      rect.fit(FitOptions.CENTER_CONTENT);
    }catch(e){
      $.writeln("placeImageIntoRect error: " + e);
    }
  }

  function storyOversetInfo(story){
    var over = false;
    var exceed = 0;
    if (story && story.isValid){
      over = story.overflows;
      if (over){
        try{
          var containers = story.textContainers;
          var visibleCount = 0;
          for (var i=0; i<containers.length; i++){
            visibleCount += containers[i].characters.length;
          }
          var totalChars = story.characters.length;
          exceed = Math.max(0, totalChars - visibleCount);
        }catch(e){
          exceed = 0;
        }
      }
    }
    return {over: over, exceed_estimate: exceed};
  }

  function adjustParagraphStylePointSizeWithin(textFrame, styleName, targetPt, minPt, maxPt){
    if (!textFrame || !styleName){
      return null;
    }
    var story = textFrame.parentStory;
    if (!story || !story.isValid){
      return null;
    }
    var applied = null;
    try{
      var paragraphs = story.paragraphs;
      for (var i=0; i<paragraphs.length; i++){
        var para = paragraphs[i];
        if (para.appliedParagraphStyle && para.appliedParagraphStyle.name === styleName){
          var newPt = targetPt;
          if (newPt < minPt) newPt = minPt;
          if (newPt > maxPt) newPt = maxPt;
          para.pointSize = newPt;
          applied = newPt;
        }
      }
      story.recompose();
    }catch(e){
      $.writeln("adjustParagraphStylePointSizeWithin error: " + e);
    }
    return applied;
  }

  function nudgeParagraphStylePointSize(textFrame, styleName, delta, minPt, maxPt){
    if (!textFrame || !styleName || delta === 0){
      return null;
    }
    var story = textFrame.parentStory;
    if (!story || !story.isValid){
      return null;
    }
    var applied = null;
    try{
      var paragraphs = story.paragraphs;
      for (var i=0; i<paragraphs.length; i++){
        var para = paragraphs[i];
        if (para.appliedParagraphStyle && para.appliedParagraphStyle.name === styleName){
          var current = para.pointSize;
          var newPt = current + delta;
          if (newPt < minPt) newPt = minPt;
          if (newPt > maxPt) newPt = maxPt;
          if (newPt !== current){
            para.pointSize = newPt;
            applied = newPt;
          }
        }
      }
      story.recompose();
    }catch(e){
      $.writeln("nudgeParagraphStylePointSize error: " + e);
    }
    return applied;
  }

  function expandFrameHeightMM(textFrame, deltaMM){
    if (!textFrame || deltaMM === 0){
      return;
    }
    try{
      var pts = deltaMM * 2.834645669;
      var bounds = textFrame.geometricBounds.slice(0);
      bounds[2] = bounds[2] + pts;
      textFrame.geometricBounds = bounds;
      var story = textFrame.parentStory;
      if (story){
        story.recompose();
      }
    }catch(e){
      $.writeln("expandFrameHeightMM error: " + e);
    }
  }

  function hideOrRemove(item, context){
    if (!item || !item.isValid){
      return;
    }
    try{
      if (item.locked){
        item.locked = false;
      }
      item.remove();
      if (context && context.pageWarnings){
        var label = "";
        try{ label = item.label || ""; }catch(ignore){}
        pushUnique(context.pageWarnings, "warning: unused_frame:" + label);
      }
    }catch(e){
      $.writeln("hideOrRemove failed: " + e);
    }
  }

  function exportPagePDF(doc, page, outFile, presetName){
    if (!doc || !page || !outFile){
      return;
    }
    try{
      if (!outFile.parent.exists){
        outFile.parent.create();
      }
    }catch(e){
      $.writeln("exportPagePDF dir error: " + e);
    }
    var preset = null;
    if (presetName){
      try{
        preset = app.pdfExportPresets.itemByName(presetName);
        if (!preset.isValid) preset = null;
      }catch(e1){
        preset = null;
      }
    }
    try{
      app.pdfExportPreferences.pageRange = page.name;
      if (preset){
        doc.exportFile(ExportFormat.PDF_TYPE, outFile, false, preset);
      } else {
        doc.exportFile(ExportFormat.PDF_TYPE, outFile);
      }
    }catch(e2){
      $.writeln("exportPagePDF error: " + e2);
      throw e2;
    }
  }

  function CsvLogger(file){
    this.file = file;
    if (file){
      try{
        file.encoding = "UTF-8";
        file.open("w");
      }catch(e){
        $.writeln("CsvLogger open error: " + e);
      }
    }
  }

  CsvLogger.prototype.writeHeader = function(columns){
    this.writeArray(columns);
  };

  CsvLogger.prototype.writeData = function(columns, data){
    if (!columns){
      return;
    }
    var row = [];
    for (var i=0; i<columns.length; i++){
      var key = columns[i];
      var value = data && data.hasOwnProperty(key) ? data[key] : "";
      if (value === null || value === undefined){
        value = "";
      }
      row.push(String(value));
    }
    this.writeArray(row);
  };

  CsvLogger.prototype.writeArray = function(values){
    if (!this.file || !this.file.opened || !values){
      return;
    }
    var parts = [];
    for (var i=0; i<values.length; i++){
      var text = String(values[i]);
      if (text.indexOf('"') >= 0 || text.indexOf(',') >= 0){
        text = '"' + text.replace(/"/g,'""') + '"';
      }
      parts.push(text);
    }
    this.file.write(parts.join(",") + "\n");
  };

  CsvLogger.prototype.close = function(){
    if (this.file && this.file.opened){
      this.file.close();
    }
  };

  function readManifestIfExists(folder){
    if (!folder || !folder.exists){
      return null;
    }
    var manifestFile = File(folder.fsName + "/manifest.json");
    if (!manifestFile.exists){
      return null;
    }
    try{
      manifestFile.open("r");
      var content = manifestFile.read();
      manifestFile.close();
      if (!content){
        return null;
      }
      return JSON.parse(content);
    }catch(e){
      try{ manifestFile.close(); }catch(ignore){}
      $.writeln("readManifestIfExists error: " + e);
      return null;
    }
  }

  function discoverNotesAndImages(pageFolder, cfg){
    var result = {notas: [], warnings: [], unusedImages: []};
    if (!pageFolder || !pageFolder.exists){
      result.warnings.push("page_folder_missing");
      return result;
    }
    var manifestData = readManifestIfExists(pageFolder);
    if (manifestData){
      var manifestWarnings = manifestData.warnings;
      if (manifestWarnings && manifestWarnings.length){
        for (var mw=0; mw<manifestWarnings.length; mw++){
          result.warnings.push(String(manifestWarnings[mw]));
        }
      }
      var notes = manifestData.notes || manifestData.notas || manifestData;
      if (!(notes instanceof Array)){
        return result;
      }
      for (var i=0; i<notes.length; i++){
        var entry = notes[i] || {};
        var note = {
          id: entry.id || entry.name || ("nota" + (result.notas.length+1)),
          images: [],
          warnings: [],
          pageName: entry.page || entry.pagina || entry.pageName || null,
          chars: entry.chars || entry.characters || 0,
          hasPhoto: entry.hasPhoto === undefined ? null : !!entry.hasPhoto
        };
        if (entry.warnings){
          if (entry.warnings instanceof Array){
            for (var ew=0; ew<entry.warnings.length; ew++){
              note.warnings.push(String(entry.warnings[ew]));
            }
          } else {
            note.warnings.push(String(entry.warnings));
          }
        }
        var docName = entry.docx || entry.word || entry.file || entry.doc || entry.docFile;
        if (docName){
          var docFile = File(pageFolder.fsName + "/" + docName);
          if (docFile.exists){
            note.fileDocx = docFile;
            note.docxName = docFile.displayName;
          } else {
            note.docxName = docName;
            note.warnings.push("missing_docx:" + docName);
          }
        }
        var images = entry.images || entry.photos || [];
        if (!(images instanceof Array)){
          images = [];
        }
        for (var im=0; im<images.length; im++){
          var imgName = images[im];
          var imgFile = File(pageFolder.fsName + "/" + imgName);
          if (imgFile.exists){
            note.images.push(imgFile);
          } else {
            note.warnings.push("missing_image:" + imgName);
          }
        }
        if (note.hasPhoto === null){
          note.hasPhoto = note.images.length > 0;
        }
        result.notas.push(note);
      }
      return result;
    }

    var files = pageFolder.getFiles();
    var docFiles = [];
    var imageFiles = [];
    for (var f=0; f<files.length; f++){
      var file = files[f];
      if (!(file instanceof File)){
        continue;
      }
      var name = file.name || "";
      var lower = name.toLowerCase();
      if (/\.docx$/.test(lower)){
        docFiles.push(file);
      } else if (/\.(jpg|jpeg|png|tif|tiff)$/.test(lower)){
        imageFiles.push(file);
      }
    }
    docFiles.sort(function(a,b){
      return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0);
    });
    imageFiles.sort(function(a,b){
      return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0);
    });
    if (docFiles.length === 0){
      result.warnings.push("no_docx_found");
    }
    for (var d=0; d<docFiles.length; d++){
      var docFile = docFiles[d];
      var baseName = docFile.displayName.replace(/\.docx$/i, "");
      var noteId = null;
      var match = baseName.match(/(nota\d+)/i);
      if (match){
        noteId = match[1].toLowerCase();
      } else {
        noteId = "nota" + (d+1);
      }
      result.notas.push({
        id: noteId,
        fileDocx: docFile,
        docxName: docFile.displayName,
        images: [],
        warnings: [],
        chars: 0,
        hasPhoto: null
      });
    }
    for (var j=0; j<result.notas.length; j++){
      var noteItem = result.notas[j];
      var prefix = noteItem.id.toLowerCase();
      for (var k=imageFiles.length-1; k>=0; k--){
        var img = imageFiles[k];
        var imgNameLower = img.displayName.toLowerCase();
        if (imgNameLower.indexOf(prefix) === 0){
          noteItem.images.unshift(img);
          imageFiles.splice(k,1);
        }
      }
      noteItem.hasPhoto = noteItem.images.length > 0;
    }
    var distribute = cfg && cfg.images && cfg.images.distribute_unmatched_images;
    if (distribute && imageFiles.length > 0 && result.notas.length > 0){
      for (var m=0; m<imageFiles.length; m++){
        var idx = m % result.notas.length;
        if (result.notas[idx]){
          result.notas[idx].images.push(imageFiles[m]);
          result.notas[idx].hasPhoto = true;
        }
      }
    } else if (imageFiles.length > 0){
      for (var e=0; e<imageFiles.length; e++){
        result.unusedImages.push(imageFiles[e]);
      }
      result.warnings.push("extra_images");
    }
    return result;
  }

  function getParagraphStylePointSize(textFrame, styleName, fallback){
    if (!textFrame || !styleName){
      return fallback;
    }
    try{
      var story = textFrame.parentStory;
      if (!story){
        return fallback;
      }
      var paragraphs = story.paragraphs;
      for (var i=0; i<paragraphs.length; i++){
        var para = paragraphs[i];
        if (para.appliedParagraphStyle && para.appliedParagraphStyle.name === styleName){
          return para.pointSize;
        }
      }
    }catch(e){
      $.writeln("getParagraphStylePointSize error: " + e);
    }
    return fallback;
  }

  function hideUnusedLabeledItems(page, usedLabels, cfg, context){
    if (!page){
      return;
    }
    try{
      var items = page.allPageItems;
      for (var i=0; i<items.length; i++){
        var item = items[i];
        var label = item.label;
        if (!label){
          continue;
        }
        if (usedLabels[label]){
          continue;
        }
        if (!cfg){
          cfg = {};
        }
        var allowRelease = cfg.images ? !!cfg.images.allow_photo_slot_release : true;
        if (!allowRelease && /_foto\d+$/i.test(label)){
          continue;
        }
        if (label.match(/^nota\d+_/i)){
          hideOrRemove(item, context);
        }
      }
    }catch(e){
      $.writeln("hideUnusedLabeledItems error: " + e);
    }
  }

  function updateNoteMetrics(note, cfg){
    if (!note){
      return;
    }
    var bodyStyle = cfg.styles.body;
    var titleStyle = cfg.styles.title;
    if (note.story){
      var info = storyOversetInfo(note.story);
      note.overset = info.over;
      note.excedente = info.exceed_estimate;
      note.body_pt = getParagraphStylePointSize(note.textFrame, bodyStyle.name, bodyStyle.pt_base);
      try{
        note.chars = note.story.characters.length;
      }catch(e){
        note.chars = 0;
      }
    } else {
      if (typeof note.overset === "undefined") note.overset = false;
      if (typeof note.excedente === "undefined") note.excedente = 0;
      if (typeof note.body_pt === "undefined") note.body_pt = bodyStyle.pt_base;
      if (typeof note.chars === "undefined") note.chars = 0;
    }
    var titleContainer = null;
    if (note.titleFrame){
      titleContainer = note.titleFrame;
    } else if (note.titleStory && note.titleStory.textContainers.length > 0){
      titleContainer = note.titleStory.textContainers[0];
    }
    if (titleContainer){
      note.title_pt = getParagraphStylePointSize(titleContainer, titleStyle.name, titleStyle.pt_base);
    }
    if (typeof note.title_pt === "undefined" || note.title_pt === null){
      note.title_pt = titleStyle.pt_base;
    }
  }

  function resolveOversetForNote(note, cfg){
    if (!note || !note.story){
      return;
    }
    var oversetCfg = cfg.overset || {};
    var bodyStyle = cfg.styles.body;
    var titleStyle = cfg.styles.title;
    var maxExpand = oversetCfg.max_expand_mm || 0;
    var bodyStep = oversetCfg.body_step_pt || 0;
    var bodyMaxDrop = oversetCfg.body_max_drop_pt || 0;
    var titleMaxDrop = oversetCfg.title_max_drop_pt || 0;
    var expanded = 0;
    var expandStep = 1;
    while (note.story.overflows && expanded < maxExpand){
      var remaining = maxExpand - expanded;
      var delta = remaining < expandStep ? remaining : expandStep;
      if (delta <= 0){
        break;
      }
      expandFrameHeightMM(note.textFrame, delta);
      expanded += delta;
    }
    var totalBodyDrop = 0;
    while (note.story.overflows && bodyStep > 0 && totalBodyDrop + bodyStep <= bodyMaxDrop + 0.0001){
      var res = nudgeParagraphStylePointSize(note.textFrame, bodyStyle.name, -bodyStep, bodyStyle.pt_min, bodyStyle.pt_max);
      if (res === null){
        break;
      }
      totalBodyDrop += bodyStep;
    }
    var titleTarget = note.titleFrame ? note.titleFrame : (note.titleStory && note.titleStory.textContainers.length ? note.titleStory.textContainers[0] : note.textFrame);
    var titleStep = titleMaxDrop;
    var totalTitleDrop = 0;
    if (titleStep <= 0 && titleMaxDrop > 0){
      titleStep = titleMaxDrop;
    }
    while (note.story.overflows && titleMaxDrop > 0 && titleTarget){
      var allowed = titleMaxDrop - totalTitleDrop;
      if (allowed <= 0){
        break;
      }
      var step = titleStep;
      if (step <= 0 || step > allowed){
        step = allowed;
      }
      var tres = nudgeParagraphStylePointSize(titleTarget, titleStyle.name, -step, titleStyle.pt_min, titleStyle.pt_max);
      if (tres === null){
        break;
      }
      totalTitleDrop += step;
    }
    if (note.story.overflows){
      if (!note.warnings){
        note.warnings = [];
      }
      note.warnings.push("overset_hard");
    }
  }

  function main(){
    var root = selectRootFolder();
    if (!root){
      alert("Sin carpeta raíz");
      return;
    }
    var cfg = readConfig(File(root.fsName + "/config.json"));
    var doc = app.activeDocument;
    if (!doc){
      alert("No hay documento activo");
      return;
    }
    var outDir = new Folder(root.fsName + "/muestras");
    if (!outDir.exists){
      outDir.create();
    }
    var csv = new CsvLogger(File(root.fsName + "/reporte.csv"));
    var NOTE_COLUMNS = ["pagina","note_id","docx","chars","hasPhoto","body_pt","title_pt","overset","warnings"];
    var SUMMARY_COLUMNS = ["pagina","total_notas","total_chars","total_imagenes","hard_warnings"];
    csv.writeHeader(NOTE_COLUMNS);

    var pageFolders = listPageFolders(root);
    if (pageFolders.length === 0){
      alert("No se encontraron carpetas de páginas");
    }

    var summaryRows = [];

    for (var i=0; i<pageFolders.length; i++){
      var pf = pageFolders[i];
      var pageNames = determinePageNames(pf.name, cfg);
      if (!pageNames || pageNames.length === 0){
        pageNames = [pf.name];
      }
      var contexts = {};
      var generatedNames = [];
      for (var p=0; p<pageNames.length; p++){
        var duplicate = duplicateBasePage(doc);
        if (!duplicate){
          csv.writeData(NOTE_COLUMNS, {pagina: pf.name, note_id: "-", warnings: "duplicate_page_failed"});
          continue;
        }
        duplicate.name = pageNames[p];
        contexts[pageNames[p]] = {
          page: duplicate,
          usedLabels: {},
          pageWarnings: [],
          hardWarnings: [],
          notes: [],
          totalChars: 0,
          totalImages: 0
        };
        generatedNames.push(pageNames[p]);
      }
      if (generatedNames.length === 0){
        continue;
      }

      var discovery = discoverNotesAndImages(pf, cfg);
      var discoveryWarnings = discovery.warnings || [];
      for (var dw=0; dw<discoveryWarnings.length; dw++){
        var warnStr = String(discoveryWarnings[dw]);
        for (var ctxKey in contexts){
          if (!contexts.hasOwnProperty(ctxKey)) continue;
          pushUnique(contexts[ctxKey].pageWarnings, "warning: " + warnStr);
          if (isHardWarning(warnStr)){
            pushUnique(contexts[ctxKey].hardWarnings, warnStr);
          }
        }
      }

      var notesList = discovery.notas || [];
      for (var n=0; n<notesList.length; n++){
        var note = notesList[n];
        if (!note){
          continue;
        }
        if (!note.warnings){
          note.warnings = [];
        }
        if (!note.docxName && note.fileDocx){
          note.docxName = note.fileDocx.displayName;
        }
        var targetName = choosePageNameForNote(note, generatedNames);
        if (!contexts[targetName]){
          targetName = generatedNames[0];
        }
        var pageCtx = contexts[targetName];
        note.pageName = targetName;
        pageCtx.notes.push(note);

        var textLabel = note.id + "_texto";
        var titleLabel = note.id + "_titulo";
        var tf = findPageItemByLabel(pageCtx.page, textLabel);
        var titleFrame = findPageItemByLabel(pageCtx.page, titleLabel);
        note.textFrame = tf;
        note.titleFrame = titleFrame;
        note.titleStory = titleFrame ? titleFrame.parentStory : null;
        if (titleFrame){
          resetTextFrame(titleFrame);
          pageCtx.usedLabels[titleLabel] = true;
        }
        var story = null;
        if (tf && note.fileDocx){
          story = placeWordIntoTextFrame(tf, note.fileDocx, cfg.importPrefs);
        } else if (!tf){
          pushUnique(note.warnings, "extra_notes_unplaced");
          pushUnique(pageCtx.hardWarnings, "extra_notes_unplaced");
        } else if (!note.fileDocx){
          pushUnique(note.warnings, "missing_docx");
        }
        if (tf){
          pageCtx.usedLabels[textLabel] = true;
        }
        if (story){
          note.story = story;
          applyParagraphStyleToStory(story, cfg.styles.body.name);
          if (titleFrame){
            try{
              var paragraphs = story.paragraphs;
              if (paragraphs.length > 1){
                var titleText = paragraphs[0].contents.replace(/\r+$/, "");
                resetTextFrame(titleFrame);
                titleFrame.contents = titleText;
                note.titleStory = titleFrame.parentStory;
                if (note.titleStory){
                  applyParagraphStyleToStory(note.titleStory, cfg.styles.title.name);
                }
                paragraphs[0].remove();
                story.recompose();
                applyParagraphStyleToStory(story, cfg.styles.body.name);
              }
            }catch(splitErr){
              $.writeln("titulo split error: " + splitErr);
            }
          }
        }

        var placedImages = 0;
        var noteImages = note.images || [];
        for (var k=0; k<noteImages.length; k++){
          var imgFrameLabel = note.id + "_foto" + (k+1);
          var imgFrame = findPageItemByLabel(pageCtx.page, imgFrameLabel);
          if (imgFrame){
            placeImageIntoRect(imgFrame, noteImages[k], cfg.images.default_object_style, cfg.images.fallback_object_style);
            pageCtx.usedLabels[imgFrameLabel] = true;
            placedImages++;
          } else {
            pushUnique(note.warnings, "extra_images");
            pushUnique(pageCtx.hardWarnings, "extra_images");
          }
        }
        note.imagesPlaced = placedImages;
        if (typeof note.hasPhoto === "undefined" || note.hasPhoto === null){
          note.hasPhoto = placedImages > 0;
        } else if (!note.hasPhoto && placedImages > 0){
          note.hasPhoto = true;
        }
      }

      for (var g=0; g<generatedNames.length; g++){
        var pageName = generatedNames[g];
        var ctx = contexts[pageName];
        var notesOnPage = ctx.notes;
        for (var nn=0; nn<notesOnPage.length; nn++){
          var noteRef = notesOnPage[nn];
          if (noteRef.story){
            resolveOversetForNote(noteRef, cfg);
          }
          updateNoteMetrics(noteRef, cfg);
          if (typeof noteRef.hasPhoto === "undefined" || noteRef.hasPhoto === null){
            noteRef.hasPhoto = (noteRef.imagesPlaced || 0) > 0;
          }
          ctx.totalChars += noteRef.chars || 0;
          ctx.totalImages += noteRef.imagesPlaced || 0;
          if (noteRef.warnings && noteRef.warnings.length){
            for (var widx=0; widx<noteRef.warnings.length; widx++){
              if (isHardWarning(noteRef.warnings[widx])){
                pushUnique(ctx.hardWarnings, noteRef.warnings[widx]);
              }
            }
          }
          var warningsText = noteRef.warnings && noteRef.warnings.length ? noteRef.warnings.join(';') : "";
          csv.writeData(NOTE_COLUMNS, {
            pagina: pageName,
            note_id: noteRef.id,
            docx: noteRef.docxName || "",
            chars: noteRef.chars || 0,
            hasPhoto: noteRef.hasPhoto ? 1 : 0,
            body_pt: noteRef.body_pt || cfg.styles.body.pt_base,
            title_pt: noteRef.title_pt || cfg.styles.title.pt_base,
            overset: noteRef.overset ? 1 : 0,
            warnings: warningsText
          });
        }

        hideUnusedLabeledItems(ctx.page, ctx.usedLabels, cfg, ctx);

        try{
          exportPagePDF(doc, ctx.page, File(outDir.fsName + "/" + pageName + ".pdf"), cfg.pdf.preset_name);
        }catch(ePdf){
          pushUnique(ctx.pageWarnings, "warning: pdf_export");
          $.writeln("exportPagePDF error: " + ePdf);
        }

        for (var pw=0; pw<ctx.pageWarnings.length; pw++){
          csv.writeData(NOTE_COLUMNS, {
            pagina: pageName,
            note_id: "-",
            docx: "",
            chars: "",
            hasPhoto: "",
            body_pt: "",
            title_pt: "",
            overset: "",
            warnings: ctx.pageWarnings[pw]
          });
        }

        summaryRows.push({
          pagina: pageName,
          total_notas: notesOnPage.length,
          total_chars: ctx.totalChars,
          total_imagenes: ctx.totalImages,
          hard_warnings: ctx.hardWarnings.join(';')
        });
      }
    }

    csv.writeHeader(SUMMARY_COLUMNS);
    for (var sr=0; sr<summaryRows.length; sr++){
      csv.writeData(SUMMARY_COLUMNS, summaryRows[sr]);
    }

    csv.close();
    alert("Listo. PDFs en /muestras y reporte.csv generado.");
  }

  main();
})();

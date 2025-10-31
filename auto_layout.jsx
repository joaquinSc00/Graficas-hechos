#target "indesign"

(function(){
  var DEFAULT_CONFIG = {
    styles: {
      body: {name: "Textos", pt_base: 9.5, pt_min: 9.0, pt_max: 10.0},
      title: {name: "Titulo 1", pt_base: 25, pt_min: 24, pt_max: 26}
    },
    images: {
      default_object_style: "Foto_2col_1cm",
      fallback_object_style: ""
    },
    layout: {
      columns: 5,
      gutter_mm: 4.25,
      photo_default_width_columns: 2,
      photo_default_height_mm: 10
    },
    pdf: {
      preset_name: "Smallest File Size"
    },
    llm: {
      enabled: false,
      endpoint: "",
      api_key_env: ""
    },
    import: {}
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

  function listPageFolders(root){
    if (!root || !root.exists){
      return [];
    }
    var folders = root.getFiles(function(f){
      return f instanceof Folder && /^\d+$/.test(f.name);
    });
    folders.sort(function(a,b){
      var na = parseInt(a.name, 10);
      var nb = parseInt(b.name, 10);
      if (!isNaN(na) && !isNaN(nb) && na !== nb){
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

  function hideOrRemove(item){
    if (!item || !item.isValid){
      return;
    }
    try{
      if (item.locked){
        item.locked = false;
      }
      item.remove();
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
    this.firstRow = true;
    if (file){
      try{
        file.encoding = "UTF-8";
        file.open("w");
      }catch(e){
        $.writeln("CsvLogger open error: " + e);
      }
    }
  }

  CsvLogger.prototype.writeRow = function(obj){
    if (!this.file || !this.file.opened){
      return;
    }
    var headers;
    if (this.firstRow && obj && obj.pagina === "pagina"){
      headers = obj;
      this.firstRow = false;
      this.file.write(serializeCsvRow(headers));
      return;
    }
    if (this.firstRow){
      this.firstRow = false;
    }
    this.file.write(serializeCsvRow(obj));
  };

  CsvLogger.prototype.close = function(){
    if (this.file && this.file.opened){
      this.file.close();
    }
  };

  function serializeCsvRow(obj){
    var keys = ["pagina","nota","chars","fotos","body_pt","title_pt","overset","excedente","warning","error"];
    var parts = [];
    for (var i=0; i<keys.length; i++){
      var key = keys[i];
      var value = obj && obj.hasOwnProperty(key) ? obj[key] : "";
      if (value === null || value === undefined){
        value = "";
      }
      var text = String(value);
      if (text.indexOf("\"") >= 0 || text.indexOf(",") >= 0){
        text = '"' + text.replace(/"/g,'""') + '"';
      }
      parts.push(text);
    }
    return parts.join(",") + "\n";
  }

  function discoverNotesAndImages(pageFolder){
    var manifest = {notas: [], warnings: []};
    if (!pageFolder || !pageFolder.exists){
      manifest.warnings.push("page_folder_missing");
      return manifest;
    }
    var files = pageFolder.getFiles();
    var docFiles = [];
    var imageFiles = [];
    for (var i=0; i<files.length; i++){
      var file = files[i];
      if (!(file instanceof File)){
        continue;
      }
      var name = file.name || "";
      var lower = name.toLowerCase();
      if (lower.match(/\.docx$/)){
        docFiles.push(file);
      } else if (lower.match(/\.(jpg|jpeg|png|tif|tiff)$/)){
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
      manifest.warnings.push("no_docx_found");
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
      manifest.notas.push({
        id: noteId,
        fileDocx: docFile,
        images: [],
        chars: 0
      });
    }

    for (var j=0; j<manifest.notas.length; j++){
      var note = manifest.notas[j];
      var prefix = note.id.toLowerCase();
      for (var k=imageFiles.length-1; k>=0; k--){
        var img = imageFiles[k];
        var imgName = img.displayName.toLowerCase();
        if (imgName.indexOf(prefix) === 0){
          note.images.unshift(img);
          imageFiles.splice(k,1);
        }
      }
    }

    if (imageFiles.length > 0 && manifest.notas.length > 0){
      for (var m=0; m<imageFiles.length; m++){
        var idx = m % manifest.notas.length;
        if (manifest.notas[idx]){
          manifest.notas[idx].images.push(imageFiles[m]);
        }
      }
    }

    return manifest;
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

  function hideUnusedLabeledItems(page, usedLabels){
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
        if (label.match(/^nota\d+_/i)){
          hideOrRemove(item);
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
    var bodyStyle = cfg.styles.body;
    var titleStyle = cfg.styles.title;
    var maxFrameExpansion = 20;
    var increment = 5;
    var expanded = 0;
    while (note.story.overflows && expanded < maxFrameExpansion){
      expandFrameHeightMM(note.textFrame, increment);
      expanded += increment;
    }
    var attempts = 0;
    while (note.story.overflows && attempts < 6){
      var res = nudgeParagraphStylePointSize(note.textFrame, bodyStyle.name, -0.25, bodyStyle.pt_min, bodyStyle.pt_max);
      if (res === null){
        break;
      }
      attempts++;
    }
    var titleAttempts = 0;
    while (note.story.overflows && titleAttempts < 3){
      var titleTarget = note.titleFrame ? note.titleFrame : note.textFrame;
      var tres = nudgeParagraphStylePointSize(titleTarget, titleStyle.name, -0.5, titleStyle.pt_min, titleStyle.pt_max);
      if (tres === null){
        break;
      }
      titleAttempts++;
    }
    if (note.story.overflows){
      note.warning = note.warning ? note.warning + ";overset" : "overset";
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
    csv.writeRow({pagina:"pagina",nota:"nota",chars:"chars",fotos:"fotos",body_pt:"body_pt",title_pt:"title_pt",overset:"overset",excedente:"excedente",warning:"warning",error:"error"});

    var pageFolders = listPageFolders(root);
    if (pageFolders.length === 0){
      alert("No se encontraron carpetas de páginas");
    }
    for (var i=0; i<pageFolders.length; i++){
      var pf = pageFolders[i];
      var page = duplicateBasePage(doc);
      if (!page){
        csv.writeRow({pagina:pf.name, nota:"-", error:"duplicate_page_failed"});
        continue;
      }
      page.name = pf.name;
      var manifest = discoverNotesAndImages(pf);
      var usedLabels = {};
      for (var w=0; w<manifest.warnings.length; w++){
        csv.writeRow({pagina:pf.name, nota:"-", warning:manifest.warnings[w]});
      }
      for (var n=0; n<manifest.notas.length; n++){
        var note = manifest.notas[n];
        try{
          var textLabel = note.id + "_texto";
          var titleLabel = note.id + "_titulo";
          var tf = findPageItemByLabel(page, textLabel);
          var titleFrame = findPageItemByLabel(page, titleLabel);
          note.textFrame = tf;
          note.titleFrame = titleFrame;
          note.titleStory = titleFrame ? titleFrame.parentStory : null;
          if (titleFrame){
            resetTextFrame(titleFrame);
          }
          var story = null;
          if (note.fileDocx && tf){
            story = placeWordIntoTextFrame(tf, note.fileDocx, cfg.import);
            note.story = story;
            if (story){
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
          }
          usedLabels[textLabel] = true;
          usedLabels[titleLabel] = true;
          for (var k=0; k<(note.images || []).length; k++){
            var imgFrameLabel = note.id + "_foto" + (k+1);
            var imgFrame = findPageItemByLabel(page, imgFrameLabel);
            if (imgFrame){
              placeImageIntoRect(imgFrame, note.images[k], cfg.images.default_object_style, cfg.images.fallback_object_style);
              usedLabels[imgFrameLabel] = true;
            }
          }
        }catch(e){
          note.error = "place_fail:" + e;
          csv.writeRow({pagina:pf.name, nota:note.id, error:note.error});
        }
      }

      for (var n2=0; n2<manifest.notas.length; n2++){
        var note2 = manifest.notas[n2];
        if (note2.story){
          resolveOversetForNote(note2, cfg);
        }
        updateNoteMetrics(note2, cfg);
      }

      hideUnusedLabeledItems(page, usedLabels);

      try{
        exportPagePDF(doc, page, File(outDir.fsName + "/" + pf.name + ".pdf"), cfg.pdf.preset_name);
      }catch(ePdf){
        csv.writeRow({pagina:pf.name, nota:"-", error:"pdf_export:" + ePdf});
      }

      for (var n3=0; n3<manifest.notas.length; n3++){
        var note3 = manifest.notas[n3];
        csv.writeRow({
          pagina: pf.name,
          nota: note3.id,
          chars: note3.chars || 0,
          fotos: (note3.images ? note3.images.length : 0),
          body_pt: note3.body_pt || cfg.styles.body.pt_base,
          title_pt: note3.title_pt || cfg.styles.title.pt_base,
          overset: note3.overset ? "1" : "0",
          excedente: note3.excedente || 0,
          warning: note3.warning || "",
          error: note3.error || ""
        });
      }
    }

    csv.close();
    alert("Listo. PDFs en /muestras y reporte.csv generado.");
  }

  main();
})();

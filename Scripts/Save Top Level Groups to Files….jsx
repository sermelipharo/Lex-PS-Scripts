/**
 * Save Top Level Groups to Files
 * Author: Lex Fradski — https://lex.mk/gh
 *
 * Photoshop ExtendScript (.jsx)
 * 
 * This script exports all visible top-level groups (Layer Sets) 
 * from the active Photoshop document into separate files.
 *
 * Features:
 *  - Choice of output format (PSD, PSB, TIFF layered, EPS, PDF, PNG, JPG, WEBP, flattened TIFF).
 *  - Destination options: same folder, subfolder, or custom location.
 *  - Duplicate group names are numbered (top→bottom or bottom→top).
 *  - Optional trim of transparent bounds.
 *  - Remembers user preferences between runs.
 *
 * File name format:
 *   <DocumentName>/<DocumentName>_<GroupName>[_NN].<ext>
 *
 * Tested on macOS, Photoshop 26.11 Beta
 */

#target photoshop
app.bringToFront();

(function () {
    if (!app.documents.length) { alert('No open document'); return; }

    var PREFS_ID = stringIDToTypeID('lex.exportTLGroups.prefs');
    var prefs = loadPrefs();

    var srcDoc = app.activeDocument;
    var baseName = stripExt(srcDoc.name);

    var groupsAll = getTopLevelVisibleGroups(srcDoc);
    if (groupsAll.length === 0) { alert('No visible top-level groups'); return; }

    // duplicates
    var nameCounts = {};
    for (var i = 0; i < groupsAll.length; i++) nameCounts[groupsAll[i].name] = (nameCounts[groupsAll[i].name] || 0) + 1;
    var hasDupes = false; for (var k in nameCounts) if (nameCounts[k] > 1) { hasDupes = true; break; }

    var dlg = null, prog = null;
    var startMs = (new Date()).getTime();

    try {
        dlg = buildDialog(prefs, hasDupes);
        if (dlg.show() !== 1) return;
    } catch (uiErr) {
        try { if (dlg) dlg.close(2); } catch(e){}
        alert('UI error: ' + uiErr);
        return;
    }

    var choice = dlg.formatList.selection.text;
    if (choice.indexOf('─') === 0) { alert('Select a format.'); return; }
    var fmt = resolveFormat(choice);

    var numberingDir = hasDupes ? (dlg.topDown.value ? 'top' : 'bottom') : 'top';
    var trimEnabled = dlg.trimBox.value;

    var destMode = dlg.dest1.value ? 1 : (dlg.dest2.value ? 2 : 3);
    var externalBase = null;
    if (destMode === 3) {
        var ptxt = dlg.pathField.text;
        if (!ptxt || !Folder(ptxt).exists) {
            var picked = Folder.selectDialog('Choose a base folder');
            if (!picked) return;
            externalBase = picked;
        } else externalBase = new Folder(ptxt);
    }

    prefs.formatLabel = fmt.label;
    prefs.destMode = destMode;
    prefs.externalPath = (destMode === 3 && externalBase) ? externalBase.fsName : (destMode === 3 ? dlg.pathField.text : '');
    prefs.numberingDir = numberingDir;
    prefs.trimEnabled = trimEnabled;
    savePrefs(prefs);

    var outFolder = resolveOutputFolder(destMode, externalBase, srcDoc, baseName);
    if (!outFolder) return;
    if (!outFolder.exists) outFolder.create();

    var groups = (numberingDir === 'top') ? groupsAll.slice() : groupsAll.slice().reverse();
    var total = groups.length;

    var counters = {};
    if (hasDupes) for (var i2 = 0; i2 < groups.length; i2++) {
        var nm = groups[i2].name;
        if (nameCounts[nm] > 1) counters[nm] = 0;
    }

    var origUnits = app.preferences.rulerUnits;
    var origDialogs = app.displayDialogs;
    app.preferences.rulerUnits = Units.PIXELS;
    app.displayDialogs = DialogModes.NO;

    try {
        for (var gi = 0; gi < total; gi++) {
            if (!prog && ((new Date()).getTime() - startMs) > 3000) {
                prog = makeProgress('Exporting groups…', total);
                prog.show();
            }
            var g = groups[gi];
            var gName = g.name;

            var suffix = '';
            if (hasDupes && nameCounts[gName] > 1) {
                counters[gName] = (counters[gName] || 0) + 1;
                suffix = '_' + pad2(counters[gName]);
            }

            var dupDoc = makeEmptyCloneDoc(srcDoc, baseName + '_' + gName + suffix);

            app.activeDocument = srcDoc;
            g.duplicate(dupDoc, ElementPlacement.INSIDE);

            app.activeDocument = dupDoc;
            if (trimEnabled) { try { dupDoc.trim(TrimType.TRANSPARENT, true, true, true, true); } catch (e) {} }

            var outFile = new File(outFolder.fsName + '/' +
                                   safeName(baseName) + '_' + safeName(gName) + suffix + '.' + fmt.ext);

            saveByFormat(dupDoc, outFile, fmt);

            dupDoc.close(SaveOptions.DONOTSAVECHANGES);

            if (prog) updateProgress(prog, gi + 1, total, gName);
        }
    } catch (err) {
        alert('Error: ' + err);
    } finally {
        app.preferences.rulerUnits = origUnits;
        app.displayDialogs = origDialogs;
        app.activeDocument = srcDoc;
        try { if (dlg) dlg.close(1); } catch(e){}
        try { if (prog) prog.close(); } catch(e){}
    }

    // ===== helpers =====
    function getTopLevelVisibleGroups(doc) {
        var res = [];
        for (var i = 0; i < doc.layerSets.length; i++) {
            var ls = doc.layerSets[i];
            if (ls.parent === doc && ls.visible) res.push(ls);
        }
        return res;
    }

    function buildDialog(saved, showDupes) {
        var d = new Window('dialog', 'Export top-level groups');
        d.orientation = 'column';
        d.alignChildren = ['fill','top'];

        d.add('statictext', undefined, 'Only top-level groups (Layer Sets) will be processed.');

        var fmtGrp = d.add('group'); fmtGrp.orientation = 'column'; fmtGrp.alignChildren = ['fill','top'];
        fmtGrp.add('statictext', undefined, 'Format:');
        var list = fmtGrp.add('dropdownlist', undefined, [
            'PSB (Large Document Format)',
            'PSD',
            'TIFF (Layered)',
            '────────────────────────',
            'EPS',
            'PDF',
            'PNG',
            'JPG',
            'WEBP',
            'TIFF (Flattened)'
        ]);
        var defaultIndex = 1;
        if (saved && saved.formatLabel) for (var i = 0; i < list.items.length; i++) if (list.items[i].text === saved.formatLabel) { defaultIndex = i; break; }
        list.selection = defaultIndex;

        var warnFlat = fmtGrp.add('statictext', undefined, 'Selected format saves a flattened copy. Layers will not be preserved.');
        var warnPdfEps = fmtGrp.add('statictext', undefined, 'EPS/PDF may not preserve original layers and text blocks (Adobe limitation).');
        setWarnColors(warnFlat); setWarnColors(warnPdfEps);

        function refreshWarnsSafe() {
            try {
                var txt = list.selection.text;
                if (txt.indexOf('─') === 0) { list.selection = 1; txt = list.selection.text; }
                var flatSet = {'PNG':1,'JPG':1,'WEBP':1,'TIFF (Flattened)':1};
                warnFlat.visible = !!flatSet[txt];
                warnPdfEps.visible = (txt === 'PDF' || txt === 'EPS');
            } catch(e) { warnFlat.visible = false; warnPdfEps.visible = false; }
        }
        list.onChange = refreshWarnsSafe;
        refreshWarnsSafe();

        var destPanel = d.add('panel', undefined, 'Destination');
        destPanel.orientation = 'column';
        destPanel.alignChildren = ['left','top'];
        destPanel.margins = 12;

        var dest1 = destPanel.add('radiobutton', undefined, 'Same folder as the working file');
        var dest2 = destPanel.add('radiobutton', undefined, 'Subfolder named after the file (in the working file folder)');
        var dest3 = destPanel.add('radiobutton', undefined, 'Create subfolder named after the file in another location');

        var chooser = destPanel.add('group'); chooser.alignChildren = ['left','center'];
        var pathField = chooser.add('edittext', undefined, (saved && saved.externalPath) ? saved.externalPath : '', {multiline:false});
        pathField.characters = 45;
        var browseBtn = chooser.add('button', undefined, 'Choose…');
        browseBtn.onClick = function () {
            try {
                var picked = Folder.selectDialog('Choose a base folder');
                if (picked) pathField.text = picked.fsName;
            } catch(e){}
        };

        function syncChooser(){ chooser.enabled = dest3.value; }
        if (saved && saved.destMode === 1) dest1.value = true;
        else if (saved && saved.destMode === 3) dest3.value = true;
        else dest2.value = true;
        syncChooser();
        dest1.onClick = dest2.onClick = dest3.onClick = syncChooser;

        // trim checkbox
        var trimBox = d.add('checkbox', undefined, 'Trim transparent bounds');
        trimBox.value = saved && typeof saved.trimEnabled !== 'undefined' ? saved.trimEnabled : true;

        var dupePanel = null, topDown = null, bottomUp = null;
        if (showDupes) {
            dupePanel = d.add('panel', undefined, 'Duplicate group names detected');
            dupePanel.orientation = 'column';
            dupePanel.alignChildren = ['left','top'];
            dupePanel.margins = 12;
            dupePanel.add('statictext', undefined, 'Numbering order:');
            var rg = dupePanel.add('group');
            topDown = rg.add('radiobutton', undefined, 'Top → Bottom');
            bottomUp = rg.add('radiobutton', undefined, 'Bottom → Top');
            if (saved && saved.numberingDir === 'bottom') bottomUp.value = true; else topDown.value = true;
            d.dupePanel = dupePanel; d.topDown = topDown; d.bottomUp = bottomUp;
        } else { d.topDown = {value:true}; d.bottomUp = {value:false}; }

        var btns = d.add('group'); btns.alignment = 'right';
        var isWindows = $.os && $.os.toLowerCase().indexOf('windows') >= 0;
        if (isWindows) {
            btns.add('button', undefined, 'OK', {name:'ok'});
            btns.add('button', undefined, 'Cancel', {name:'cancel'});
        } else {
            btns.add('button', undefined, 'Cancel', {name:'cancel'});
            btns.add('button', undefined, 'OK', {name:'ok'});
        }

        d.formatList = list;
        d.dest1 = dest1; d.dest2 = dest2; d.dest3 = dest3;
        d.pathField = pathField;
        d.trimBox = trimBox;
        return d;
    }

    // progress palette
    function makeProgress(title, total) {
        var w = new Window('palette', title);
        w.orientation = 'column';
        w.alignChildren = ['fill','top'];
        var st = w.add('statictext', undefined, 'Starting…'); st.characters = 50;
        var bar = w.add('progressbar', undefined, 0, total); bar.preferredSize.width = 400;
        var pct = w.add('statictext', undefined, '0%'); pct.alignment = 'right';
        w.onShow = function(){ try { w.center(); } catch(e){} };
        w._st = st; w._bar = bar; w._pct = pct;
        return w;
    }
    function updateProgress(w, done, total, name) {
        try {
            w._bar.value = done;
            var p = Math.floor((done/total)*100);
            w._pct.text = p + '%';
            w._st.text = 'Processing: ' + String(name);
            w.update(); app.refresh();
        } catch(e){}
    }

    function setWarnColors(st) { try { st.graphics.foregroundColor = st.graphics.newPen(st.graphics.PenType.SOLID_COLOR, [1,0,0], 1);} catch(e){} st.visible=false; }

    function resolveFormat(label) {
        switch (label) {
            case 'PSB (Large Document Format)': return {id:'PSB',  ext:'psb', asCopy:false, layered:true,  label:label};
            case 'PSD':                         return {id:'PSD',  ext:'psd', asCopy:false, layered:true,  label:label};
            case 'TIFF (Layered)':              return {id:'TIFFL',ext:'tif', asCopy:false, layered:true,  label:label};
            case 'EPS':                         return {id:'EPS',  ext:'eps', asCopy:true,  layered:false, label:label};
            case 'PDF':                         return {id:'PDF',  ext:'pdf', asCopy:true,  layered:false, label:label};
            case 'PNG':                         return {id:'PNG',  ext:'png', asCopy:true,  layered:false, label:label};
            case 'JPG':                         return {id:'JPG',  ext:'jpg', asCopy:true,  layered:false, label:label};
            case 'WEBP':                        return {id:'WEBP', ext:'webp',asCopy:true,  layered:false, label:label};
            case 'TIFF (Flattened)':            return {id:'TIFFF',ext:'tif', asCopy:true,  layered:false, label:label};
            default: return null;
        }
    }

    // helper: try to get working folder without forcing Save
function getWorkingFolder(doc) {
    // 1) normal saved/opened docs
    try { if (doc.path && Folder(doc.path).exists) return doc.path; } catch (e) {}
    // 2) some builds expose fullName for opened files
    try {
        var f = File(doc.fullName);
        if (f && f.exists) return f.parent;
    } catch (e2) {}
    // 3) no path available (e.g., Untitled) -> null
    return null;
}

function resolveOutputFolder(mode, externalBase, doc, base) {
    try {
        var workFolder = getWorkingFolder(doc);

        if (mode === 1) {
            // same folder as working file; if unknown, ask once
            if (!workFolder) {
                var picked1 = Folder.selectDialog('Working file folder is unknown. Choose a base folder:');
                if (!picked1) return null;
                return picked1; // no subfolder
            }
            return workFolder;
        }

        if (mode === 2) {
            var parent = workFolder;
            if (!parent) {
                var picked2 = Folder.selectDialog('Working file folder is unknown. Choose a base folder:');
                if (!picked2) return null;
                parent = picked2;
            }
            var f2 = new Folder(parent.fsName + '/' + safeName(base));
            if (!f2.exists) f2.create();
            return f2;
        }

        // mode === 3
        var baseF = externalBase;
        if (!baseF) {
            var txt = (typeof prefs !== 'undefined' && prefs.externalPath) ? prefs.externalPath : '';
            baseF = txt ? new Folder(txt) : null;
            if (!baseF || !baseF.exists) {
                var picked3 = Folder.selectDialog('Choose a base folder');
                if (!picked3) return null;
                baseF = picked3;
            }
        }
        var f3 = new Folder(baseF.fsName + '/' + safeName(base));
        if (!f3.exists) f3.create();
        return f3;

    } catch (e) {
        alert('Cannot resolve output folder: ' + e);
        return null;
    }
}

    function makeEmptyCloneDoc(src, newName) {
        var w = src.width, h = src.height, res = src.resolution;
        var newMode = toNewDocMode(src.mode);
        var bits = BitsPerChannelType.EIGHT; try { bits = src.bitsPerChannel; } catch (e) {}
        var prof = null; try { prof = src.colorProfileName; } catch (e) {}

        var doc = app.documents.add(w, h, res, newName, newMode, DocumentFill.TRANSPARENT);
        try { if (doc.bitsPerChannel !== bits) doc.bitsPerChannel = bits; } catch (e) {}
        try { if (prof) doc.convertProfile(prof, Intent.RELATIVECOLORIMETRIC, true, true); } catch (e) {}
        return doc;
    }

    function toNewDocMode(m) {
        switch (m) {
            case DocumentMode.RGB: return NewDocumentMode.RGB;
            case DocumentMode.CMYK: return NewDocumentMode.CMYK;
            case DocumentMode.GRAYSCALE: return NewDocumentMode.GRAYSCALE;
            case DocumentMode.LAB: return NewDocumentMode.LAB;
            default: return NewDocumentMode.RGB;
        }
    }

    function saveByFormat(doc, file, fmt) {
        app.activeDocument = doc;

        if (fmt.id === 'PSD') {
            var o = new PhotoshopSaveOptions();
            o.layers = true; o.embedColorProfile = true; o.alphaChannels = true; o.spotColors = true;
            doc.saveAs(file, o, false); return;
        }
        if (fmt.id === 'PSB') {
            var o2 = new LargeDocumentSaveOptions();
            o2.layers = true; o2.embedColorProfile = true; o2.alphaChannels = true;
            doc.saveAs(file, o2, false); return;
        }
        if (fmt.id === 'TIFFL') {
            var t = new TiffSaveOptions();
            t.layers = true; t.embedColorProfile = true;
            try { t.imageCompression = TIFFEncoding.NONE; } catch(e){}
            doc.saveAs(file, t, false); return;
        }

        // flattened / copy
        if (fmt.id === 'EPS') {
            var e = new EPSSaveOptions();
            e.embedColorProfile = true;
            try { e.preview = EPSPreview.NONE; } catch (ex) { try { e.preview = Preview.NONE; } catch (_) {} }
            e.encoding = SaveEncoding.BINARY;
            e.halftoneScreen = false; e.transferFunction = false;
            try { e.includeDocumentThumbnails = false; } catch (ex) { try { e.includeDocumentThumbNail = false; } catch (_) {} }
            try { e.vectorData = true; } catch (_) {}
            doc.saveAs(file, e, true); return;
        }
        if (fmt.id === 'PDF') {
            var p = new PDFSaveOptions();
            try { p.pDFPreset = 'High Quality Print'; } catch (e1) { try { p.pdfPreset = 'High Quality Print'; } catch(_){} }
            try { p.layers = false; } catch (_) {}
            doc.saveAs(file, p, true); return;
        }
        if (fmt.id === 'PNG') {
            try { var png = new PNGSaveOptions(); png.interlaced = false; doc.saveAs(file, png, true); }
            catch (e1) { exportSFW(doc, file, 'PNG24'); }
            return;
        }
        if (fmt.id === 'JPG') {
            ensureOpaqueBackground(doc, [255,255,255]);
            var j = new JPEGSaveOptions(); j.quality = 12; j.embedColorProfile = true; j.matte = MatteType.WHITE;
            doc.saveAs(file, j, true); return;
        }
        if (fmt.id === 'WEBP') {
            try { var w = new WebPSaveOptions(); w.lossless = false; w.quality = 80; doc.saveAs(file, w, true); }
            catch (e2) { exportSFW(doc, file, 'PNG24'); }
            return;
        }
        if (fmt.id === 'TIFFF') {
            try { if (doc.layers) doc.flatten(); } catch(e){ doc.flatten(); }
            var tf = new TiffSaveOptions();
            tf.embedColorProfile = true;
            try { tf.imageCompression = TIFFEncoding.NONE; } catch(e){}
            doc.saveAs(file, tf, true); return;
        }
        throw new Error('Unknown format id: ' + fmt.id);
    }

    function exportSFW(doc, file, kind) {
        var opts = new ExportOptionsSaveForWeb();
        if (kind === 'PNG24') {
            opts.format = SaveDocumentType.PNG;
            opts.PNG8 = false;
            opts.transparency = true;
            opts.interlaced = false;
        }
        app.activeDocument.exportDocument(file, ExportType.SAVEFORWEB, opts);
    }

    function ensureOpaqueBackground(doc, rgb) {
        var lay = doc.artLayers.add();
        lay.name = '_bg_flatten_helper';
        lay.move(doc, ElementPlacement.PLACEATEND);
        lay.kind = LayerKind.NORMAL;
        app.foregroundColor = solidRGB(rgb[0], rgb[1], rgb[2]);
        doc.selection.selectAll();
        doc.selection.fill(app.foregroundColor);
        doc.selection.deselect();
        doc.flatten();
    }

    function solidRGB(r,g,b) { var c = new SolidColor(); c.rgb.red=r; c.rgb.green=g; c.rgb.blue=b; return c; }
    function safeName(s){ s=(s==null)?'':String(s); s=s.replace(/[\\\/:\*\?"<>|]/g,'_').replace(/\s+/g,' '); s=s.replace(/^\s+|\s+$/g,''); return s||'untitled'; }
    function stripExt(n){ var i=n.lastIndexOf('.'); return i>0?n.substring(0,i):n; }
    function pad2(n){ return (n<10?'0':'')+n; }

    function loadPrefs(){ var o={}; try{ var d=app.getCustomOptions(PREFS_ID);
        var sid=function(s){return stringIDToTypeID(s);};
        if(d.hasKey(sid('formatLabel'))) o.formatLabel=d.getString(sid('formatLabel'));
        if(d.hasKey(sid('destMode'))) o.destMode=d.getInteger(sid('destMode'));
        if(d.hasKey(sid('externalPath'))) o.externalPath=d.getString(sid('externalPath'));
        if(d.hasKey(sid('numberingDir'))) o.numberingDir=d.getString(sid('numberingDir'));
        if(d.hasKey(sid('trimEnabled'))) o.trimEnabled=d.getBoolean(sid('trimEnabled'));
    }catch(e){} return o; }
    function savePrefs(p){ try{ var d=new ActionDescriptor();
        var sid=function(s){return stringIDToTypeID(s);};
        if(p.formatLabel) d.putString(sid('formatLabel'), p.formatLabel);
        if(p.destMode) d.putInteger(sid('destMode'), p.destMode);
        d.putString(sid('externalPath'), p.externalPath || '');
        d.putString(sid('numberingDir'), p.numberingDir || 'top');
        d.putBoolean(sid('trimEnabled'), !!p.trimEnabled);
        app.putCustomOptions(PREFS_ID, d, true);
    }catch(e){} }
})();
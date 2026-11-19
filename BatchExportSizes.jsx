#target illustrator
(function () {
    if (app.documents.length === 0) { alert('Open your A0 or US Letter Illustrator file first.'); return; }
    var doc = app.activeDocument;
    var sizeLayers = [], allTop = [], i, j;
    for (i = 0; i < doc.layers.length; i++) {
        var l = doc.layers[i];
        allTop.push({layer:l, visible:l.visible, locked:l.locked});
        var m = /^Size-(\d+)$/i.exec(l.name);
        if (m && Number(m[1]) >= 33 && Number(m[1]) <= 44) sizeLayers.push({layer:l, size:Number(m[1])});
    }
    if (sizeLayers.length !== 12) {
        alert('Expected 12 top-level layers Size-33 through Size-44; found ' + sizeLayers.length + '. No changes made.'); return;
    }
    var format = prompt('Export format label (A0 or Letter):', doc.artboards.length === 1 ? 'A0' : 'Letter');
    if (format === null) return;
    format = format.replace(/^\s+|\s+$/g, '');
    if (!/^(A0|Letter)$/i.test(format)) { alert('Enter A0 or Letter.'); return; }
    format = /^a0$/i.test(format) ? 'A0' : 'Letter';
    var code = prompt('Pattern number for PDF filenames:', 'SS-111626');
    if (code === null) return;
    code = code.replace(/^\s+|\s+$/g, '');
    if (!/^[A-Za-z0-9_-]+$/.test(code)) { alert('Pattern number may only contain letters, numbers, underscores and hyphens.'); return; }
    var folder = Folder.selectDialog('Choose a folder for the ' + format + ' PDFs');
    if (!folder) return;
    var exported = [], errors = [], options = new PDFSaveOptions();
    options.preserveEditability = false;
    options.viewAfterSaving = false;
    // All Illustrator artboards are exported as separate pages in a multipage PDF.
    options.artboardRange = '1-' + doc.artboards.length;
    try {
        for (i = 0; i < allTop.length; i++) allTop[i].layer.locked = false;
        for (i = 33; i <= 44; i++) {
            for (j = 0; j < allTop.length; j++) {
                // Hide ALL top-level layers except the active size to prevent
                // accidental duplicate master artwork or stray visible layers.
                allTop[j].layer.visible = (allTop[j].layer.name.toLowerCase() === ('size-' + i));
            }
            app.redraw();
            var target = new File(folder.fsName + '/' + code + '-' + i + '-' + format + '.pdf');
            try {
                doc.saveAs(target, options);
                exported.push(target.name);
            } catch (e) { errors.push('Size-' + i + ': ' + e); }
        }
    } finally {
        for (i = 0; i < allTop.length; i++) {
            try { allTop[i].layer.locked = false; allTop[i].layer.visible = allTop[i].visible; } catch(e) {}
        }
        for (i = 0; i < allTop.length; i++) {
            try { allTop[i].layer.locked = allTop[i].locked; } catch(e) {}
        }
        app.redraw();
    }
    alert('Exported ' + exported.length + ' PDF(s) to:\n' + folder.fsName +
          (errors.length ? '\n\nErrors:\n' + errors.join('\n') : '') +
          '\n\nCheck PDF page sizes and visible artwork before uploading to your shop.');
})();

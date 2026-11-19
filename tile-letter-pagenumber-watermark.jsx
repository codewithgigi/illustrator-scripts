#target illustrator
(function () {
    if (app.documents.length === 0) { alert('Open your Illustrator document first.'); return; }
    var doc = app.activeDocument;
    var PT_W = 612, PT_H = 792; // US Letter portrait, points
    var SOURCE_ARTBOARD = doc.artboards.getActiveArtboardIndex();
    var bounds = doc.artboards[SOURCE_ARTBOARD].artboardRect; // left, top, right, bottom
    var left = bounds[0], top = bounds[1];
    var width = bounds[2] - bounds[0], height = bounds[1] - bounds[3];
    var cols = Math.ceil(width / PT_W - 0.000001);
    var rows = Math.ceil(height / PT_H - 0.000001);
    var count = cols * rows;
    if (count > 100) { alert('This layout needs ' + count + ' artboards; Illustrator supports at most 100. Reduce the layout or divide it.'); return; }
    if (!confirm('Create ' + count + ' US Letter artboards (' + cols + ' columns x ' + rows + ' rows) from the active artboard top-left?\n\nThis replaces the document artboards, but does not move pattern artwork. Save a backup first.')) return;

    var layoutName = 'Page Layout';
    var layout;
    try { layout = doc.layers.getByName(layoutName); } catch(e) { layout = doc.layers.add(); layout.name = layoutName; }
    if (layout.locked) { alert('Unlock the existing Page Layout layer and rerun.'); return; }
    // Only delete objects on the dedicated layout layer, never source artwork.
    while (layout.pageItems.length > 0) { layout.pageItems[0].remove(); }
    while (layout.layers.length > 0) { layout.layers[0].remove(); }
    layout.visible = true;

    // Illustrator artboards are listed in reading order: left to right, top to bottom.
    // Reuse first artboard, delete the others, then add the tiled artboards.
    for (var i = doc.artboards.length - 1; i >= 1; i--) doc.artboards.remove(i);
    var black = new RGBColor(); black.red = 80; black.green = 80; black.blue = 80;
    var gray = new RGBColor(); gray.red = 175; gray.green = 175; gray.blue = 175;
    function addLabel(label, x, y, fontSize, color) {
        var t = layout.textFrames.add();
        t.contents = label;
        t.position = [x, y];
        t.textRange.characterAttributes.size = fontSize;
        t.textRange.characterAttributes.fillColor = color;
        return t;
    }
    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var n = r * cols + c + 1;
            var x = left + c * PT_W;
            var y = top - r * PT_H;
            var rect = [x, y, x + PT_W, y - PT_H];
            if (n === 1) doc.artboards[0].artboardRect = rect;
            else doc.artboards.add(rect);
            // Place labels within a conservative 0.35-inch printer-safe zone.
            var number = String(n);
            var p = addLabel(number, x + PT_W / 2, y - 29, 42, black);
            p.position = [x + (PT_W - p.width) / 2, y - 29];
            var wm = addLabel('\u00a9 2026 StyledStitch', x + PT_W / 2, y - PT_H + 27, 7, gray);
            wm.position = [x + (PT_W - wm.width) / 2, y - PT_H + 27];
        }
    }
    doc.artboards.setActiveArtboardIndex(0);
    alert('Created ' + count + ' US Letter artboards, with page numbers and copyright footers.\n\nIMPORTANT: Export each size separately by hiding all other sizes and master layers. In PDF options choose All artboards, 100% / Actual Size when printing, and no Fit to Page.');
})();

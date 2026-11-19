/* StyledStitch 00 — Overlapping US Letter artboards
   Run on a NEW, BLANK Illustrator document before adding artwork or marks.
   8.5 x 11 in artboards, 6.5 in horizontal / 9 in vertical advance.
   Physical artboard overlap = 2 in; printable-region overlap = 1 in
   when using 0.5 in margins.
*/
#target illustrator
(function () {
    if (app.documents.length === 0) { alert('Open a blank Illustrator document first.'); return; }
    var doc = app.activeDocument;
    var cols = parseInt(prompt('Number of columns?', '2'), 10);
    if (isNaN(cols) || cols < 1 || cols > 20) return;
    var rows = parseInt(prompt('Number of rows?', '2'), 10);
    if (isNaN(rows) || rows < 1 || rows > 20 || cols * rows > 100) { alert('Use 1–20 rows, maximum 100 artboards.'); return; }
    if (doc.artboards.length !== 1 || doc.pageItems.length !== 0) {
        alert('For safety, run this only on a NEW BLANK document with exactly one artboard and no artwork.');
        return;
    }
    if (!confirm('Create ' + (cols * rows) + ' overlapping US Letter artboards? This will reposition the blank starting artboard.')) return;
    var inch = 72;
    var w = 8.5 * inch, h = 11 * inch;
    var dx = 6.5 * inch, dy = 9 * inch;
    // Keep the existing artboard top-left as the anchor.
    var original = doc.artboards[0].artboardRect;
    var left = original[0], top = original[1];
    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var n = r * cols + c;
            var x = left + c * dx, y = top - r * dy;
            var bounds = [x, y, x + w, y - h];
            if (n === 0) doc.artboards[0].artboardRect = bounds;
            else doc.artboards.add(bounds);
            doc.artboards[n].name = String(n + 1);
        }
    }
    doc.artboards.setActiveArtboardIndex(0);
    alert('Created ' + (cols * rows) + ' overlapping US Letter artboards.\n\nArtboard spacing: 6.5 in right, 9 in down.\nArtboard overlap: 2 in; printable-area overlap: 1 in (with 0.5 in margins).\n\nNEXT: Do not run page-number or diamond scripts yet. Overlapping artboards can cause marks to appear on multiple exported pages; we will test/fix that separately.');
})();

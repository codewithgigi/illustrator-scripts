/* StyledStitch — Script 02: Page Numbers
   Adobe Illustrator ExtendScript (.jsx)
   Adds simple printed page numbers to existing artboards.
   Does NOT add, resize, move, or rename artboards.
*/
#target illustrator
(function () {
    if (app.documents.length === 0) {
        alert('Open your Illustrator template first.');
        return;
    }
    var doc = app.activeDocument;
    var layerName = '02 - PAGE NUMBERS';
    var layer;
    var i;
    for (i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            layer = doc.layers[i];
            break;
        }
    }
    if (!layer) {
        layer = doc.layers.add();
        layer.name = layerName;
    } else {
        if (!confirm('The page-number layer already exists. Replace its contents?')) return;
        layer.locked = false;
        layer.visible = true;
        while (layer.pageItems.length > 0) {
            layer.pageItems[0].remove();
        }
    }
    layer.locked = false;
    layer.visible = true;

    var black = new RGBColor();
    black.red = 0; black.green = 0; black.blue = 0;
    var PT = 72;
    var margin = 0.5 * PT;
    var fontSize = 10;
    var offsetAbovePrintableBottom = 0.16 * PT;

    for (i = 0; i < doc.artboards.length; i++) {
        var bounds = doc.artboards[i].artboardRect;
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        var t = layer.textFrames.add();
        t.contents = String(i + 1);
        t.textRange.characterAttributes.size = fontSize;
        t.textRange.characterAttributes.fillColor = black;
        t.textRange.paragraphAttributes.justification = Justification.CENTER;
        // Center of page, just above the lower printable boundary.
        t.position = [(left + right) / 2, bottom + margin + offsetAbovePrintableBottom];
    }
    doc.activeLayer = layer;
    alert('Added page numbers 1–' + doc.artboards.length + ' on the "' + layerName + '" layer. Artboards were not changed.');
})();

/* StyledStitch - Script 03: 0.5-inch margin guides
   Adobe Illustrator ExtendScript (.jsx)
   Adds editable, non-export-intended rectangles to each existing artboard.
   Does not change artboards or any other layers.
*/
#target illustrator
(function () {
    if (app.documents.length === 0) {
        alert('Open your tiled Illustrator document first.');
        return;
    }
    var doc = app.activeDocument;
    var layerName = '03 - MARGIN GUIDES (HIDE BEFORE EXPORT)';
    var margin = 0.5 * 72; // Illustrator uses points
    var layer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            layer = doc.layers[i];
            break;
        }
    }
    if (layer !== null) {
        if (!confirm('Margin guide layer already exists. Replace its contents?')) return;
        layer.locked = false;
        layer.visible = true;
        while (layer.pageItems.length > 0) layer.pageItems[0].remove();
    } else {
        layer = doc.layers.add();
        layer.name = layerName;
    }
    var color = new RGBColor();
    color.red = 90;
    color.green = 155;
    color.blue = 195;
    var count = 0;
    for (var a = 0; a < doc.artboards.length; a++) {
        var b = doc.artboards[a].artboardRect; // left, top, right, bottom
        var left = b[0], top = b[1], width = b[2] - b[0], height = b[1] - b[3];
        if (width <= 2 * margin || height <= 2 * margin) continue;
        var guide = layer.pathItems.rectangle(
            top - margin,
            left + margin,
            width - 2 * margin,
            height - 2 * margin
        );
        guide.filled = false;
        guide.stroked = true;
        guide.strokeColor = color;
        guide.strokeWidth = 0.5;
        count++;
    }
    layer.visible = true;
    doc.activeLayer = layer;
    alert('Added 0.5-inch margin rectangles to ' + count + ' artboard(s).\n\nHide the "' + layerName + '" layer before exporting your pattern PDF.');
})();

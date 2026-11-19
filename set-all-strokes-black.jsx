#target illustrator
(function () {
    if (app.documents.length === 0) {
        alert('Open an Illustrator document first.');
        return;
    }
    var doc = app.activeDocument;
    var changed = 0, skipped = 0;
    var black = new RGBColor();
    black.red = 0; black.green = 0; black.blue = 0;

    // Document-wide pageItems includes artwork inside groups and sublayers.
    // Work on paths and text only, so group containers are not modified.
    for (var i = 0; i < doc.pageItems.length; i++) {
        var item = doc.pageItems[i];
        try {
            if (item.typename === 'PathItem') {
                if (item.stroked) {
                    item.strokeColor = black;
                    item.strokeWidth = 1;
                    changed++;
                }
            } else if (item.typename === 'TextFrame') {
                // Apply to text only when its characters already have strokes.
                var chars = item.textRange.characters;
                for (var j = 0; j < chars.length; j++) {
                    var attr = chars[j].characterAttributes;
                    if (attr.strokeWeight > 0 && attr.strokeColor.typename !== 'NoColor') {
                        attr.strokeColor = black;
                        attr.strokeWeight = 1;
                        changed++;
                    }
                }
            }
        } catch (e) { skipped++; }
    }
    alert('Finished. Updated ' + changed + ' stroked paths/text characters to black, 1 pt.\nSkipped: ' + skipped + '.\nObjects without strokes were unchanged.');
})();

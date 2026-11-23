var doc = app.activeDocument;

// 1. Prompt for the source layer name (flexible, case-insensitive)
var sourceInput = prompt("Enter the name of the markings layer to copy:", "pattrnmarkings");

if (sourceInput) {
    var sourceLayer = null;
    var searchName = sourceInput.toLowerCase();

    // Find the matching top-level layer (ignoring uppercase/lowercase differences)
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name.toLowerCase() === searchName) {
            sourceLayer = doc.layers[i];
            break;
        }
    }

    if (!sourceLayer) {
        alert("Could not find a top-level layer named '" + sourceInput + "'.");
    } else if (sourceLayer.locked || !sourceLayer.visible) {
        alert("Please unlock and make '" + sourceLayer.name + "' visible before running the script.");
    } else {
        // 2. Find all top-level layers starting with "size-"
        var sizeLayers = [];
        for (var j = 0; j < doc.layers.length; j++) {
            var lyr = doc.layers[j];
            if (lyr !== sourceLayer && lyr.name.toLowerCase().indexOf("size-") === 0) {
                sizeLayers.push(lyr);
            }
        }

        if (sizeLayers.length === 0) {
            alert("No layers starting with 'size-' were found.");
        } else {
            var copiedCount = 0;

            // 3. Helper function to duplicate items or sublayers cleanly into a target layer
            function copyContents(fromContainer, toLayer) {
                // Duplicate sublayers recursively (in reverse to preserve stacking order)
                for (var s = fromContainer.layers.length - 1; s >= 0; s--) {
                    var srcSub = fromContainer.layers[s];
                    if (srcSub.locked || !srcSub.visible) continue;

                    var newSub = toLayer.layers.add();
                    newSub.name = srcSub.name;
                    copyContents(srcSub, newSub);
                }

                // Duplicate all pageItems (text, grainline paths, groups) in reverse order
                for (var p = fromContainer.pageItems.length - 1; p >= 0; p--) {
                    var item = fromContainer.pageItems[p];
                    if (!item.locked && !item.hidden) {
                        var dup = item.duplicate(toLayer, ElementPlacement.PLACEATBEGINNING);
                        dup.name = item.name;
                    }
                }
            }

            // 4. Copy into every unlocked, visible size-XX layer
            for (var k = 0; k < sizeLayers.length; k++) {
                var targetSizeLayer = sizeLayers[k];
                if (!targetSizeLayer.locked && targetSizeLayer.visible) {
                    // Create a matching sublayer inside each size-XX layer to keep markings organized
                    var markingsSub = targetSizeLayer.layers.add();
                    markingsSub.name = sourceLayer.name;

                    copyContents(sourceLayer, markingsSub);
                    copiedCount++;
                }
            }

            alert("Copied '" + sourceLayer.name + "' into " + copiedCount + " size layers.");
        }
    }
}

var doc = app.activeDocument;

// Prompt for the starting and ending numbers of the range
var startSize = parseInt(prompt("Enter the starting size number:", "35"), 10);
var endSize = parseInt(prompt("Enter the ending size number:", "51"), 10);

if (!isNaN(startSize) && !isNaN(endSize) && startSize <= endSize) {
    // 1. Pre-create buckets for every number in the range (35 through 51)
    var buckets = {};
    for (var n = startSize; n <= endSize; n++) {
        buckets[n.toString()] = [];
    }

    // Regex that strictly checks for "-<number>" at the very end of the name
    var trailingRegex = /-(\d+)$/;

    // 2. Scan the document recursively from top to bottom
    function scanContainer(container) {
        // Check sublayers
        if (container.layers) {
            for (var i = 0; i < container.layers.length; i++) {
                var sub = container.layers[i];
                // Skip destination layers if the script is run a second time
                if (sub.name.indexOf("size-") === 0) continue;

                var match = sub.name.match(trailingRegex);
                if (match && buckets.hasOwnProperty(match[1])) {
                    buckets[match[1]].push(sub);
                } else {
                    scanContainer(sub);
                }
            }
        }

        // Check pageItems (paths, groups, compound paths, text, etc.)
        if (container.pageItems) {
            for (var j = 0; j < container.pageItems.length; j++) {
                var item = container.pageItems[j];
                var match = item.name.match(trailingRegex);

                if (match && buckets.hasOwnProperty(match[1])) {
                    buckets[match[1]].push(item);
                } else if (item.typename === "GroupItem") {
                    // Only look inside a group if the group itself didn't match
                    scanContainer(item);
                }
            }
        }
    }

    scanContainer(doc);

    // 3. Create layers size-35 through size-51 and move matching items
    var totalMoved = 0;

    // Loop backwards from 51 down to 35 so size-35 sits at the top of the Layers panel
    for (var s = endSize; s >= startSize; s--) {
        var sizeKey = s.toString();
        var targetLayerName = "size-" + sizeKey;
        var targetLayer;

        try {
            targetLayer = doc.layers.getByName(targetLayerName);
        } catch (e) {
            targetLayer = doc.layers.add();
            targetLayer.name = targetLayerName;
        }

        // Move items in reverse order to preserve their original top-to-bottom stacking
        var items = buckets[sizeKey];
        for (var k = items.length - 1; k >= 0; k--) {
            var obj = items[k];
            if (!obj.locked && !obj.hidden) {
                obj.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                totalMoved++;
            }
        }
    }

    alert("Done! Created size-" + startSize + " through size-" + endSize + " and moved " + totalMoved + " items.");
} else {
    alert("Invalid number range.");
}

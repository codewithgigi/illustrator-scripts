#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("Open your Illustrator pattern document first.");
        return;
    }

    var doc = app.activeDocument;
    var MIN_SIZE = 33;
    var MAX_SIZE = 44;
    var moved = 0;
    var copiedNeckbands = 0;
    var errors = [];
    var plans = [];

    function findTopLayer(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name.toLowerCase() === name.toLowerCase()) {
                return doc.layers[i];
            }
        }
        return null;
    }

    function findChildLayer(parent, name) {
        for (var i = 0; i < parent.layers.length; i++) {
            if (parent.layers[i].name === name) return parent.layers[i];
        }
        return null;
    }

    function parseGroup(name) {
        var match = name.match(/(?:^|[^0-9])(3[3-9]|4[0-4])$/);
        if (!match) return null;
        var size = parseInt(match[1], 10);
        var pieceMatch = name.match(/^(front|back|sleeve)(?=[^a-z]|\d|$)/i);
        if (!pieceMatch) return null;
        var piece = pieceMatch[1].toLowerCase();
        return { size: size, piece: piece };
    }

    var patterns = findTopLayer("patterns");
    var notches = findTopLayer("notches");
    if (!patterns || !notches) {
        alert("Expected top-level layers named 'patterns' and 'notches'. Nothing changed.");
        return;
    }

    function collect(source) {
        // Only direct children: moving a group preserves everything inside it.
        for (var i = 0; i < source.groupItems.length; i++) {
            var group = source.groupItems[i];
            if (group.parent !== source) continue;
            var info = parseGroup(group.name);
            if (info) {
                plans.push({ item: group, size: info.size, piece: info.piece });
            }
        }
    }

    collect(patterns);
    collect(notches);

    if (plans.length === 0) {
        alert("No matching named groups found. Nothing changed.");
        return;
    }

    var neckband = null;
    for (var n = 0; n < patterns.groupItems.length; n++) {
        var candidate = patterns.groupItems[n];
        if (candidate.parent === patterns &&
            candidate.name.toLowerCase() === "neckband-allsizes") {
            neckband = candidate;
            break;
        }
    }

    var confirmed = confirm(
        "Found " + plans.length + " groups to organize into Size-33 through Size-44.\n" +
        "The neckband will be copied into every size, and the original kept.\n\n" +
        "Run this on a backup copy. Continue?"
    );
    if (!confirmed) return;

    // Unlock only the source layers required for moving the groups.
    patterns.locked = false;
    notches.locked = false;

    for (var j = 0; j < plans.length; j++) {
        var plan = plans[j];
        try {
            var sizeName = "Size-" + plan.size;
            var sizeLayer = findTopLayer(sizeName);
            if (!sizeLayer) {
                sizeLayer = doc.layers.add();
                sizeLayer.name = sizeName;
            }
            sizeLayer.locked = false;
            sizeLayer.visible = true;

            var pieceName = plan.piece.charAt(0).toUpperCase() +
                plan.piece.substring(1) + "-" + plan.size;
            var pieceLayer = findChildLayer(sizeLayer, pieceName);
            if (!pieceLayer) {
                pieceLayer = sizeLayer.layers.add();
                pieceLayer.name = pieceName;
            }
            pieceLayer.locked = false;
            pieceLayer.visible = true;

            plan.item.locked = false;
            plan.item.move(pieceLayer, ElementPlacement.PLACEATEND);
            moved++;
        } catch (e) {
            errors.push(plan.item.name + ": " + e.message);
        }
    }

    // Duplicate the master neckband into each size's own sublayer.
    // Re-running the script does not create duplicate neckbands.
    if (neckband) {
        for (var size = MIN_SIZE; size <= MAX_SIZE; size++) {
            try {
                var sizeLayer = findTopLayer("Size-" + size);
                if (!sizeLayer) {
                    sizeLayer = doc.layers.add();
                    sizeLayer.name = "Size-" + size;
                }
                sizeLayer.locked = false;
                sizeLayer.visible = true;
                var neckName = "Neckband-" + size;
                var neckLayer = findChildLayer(sizeLayer, neckName);
                if (!neckLayer) {
                    neckLayer = sizeLayer.layers.add();
                    neckLayer.name = neckName;
                }
                neckLayer.locked = false;
                neckLayer.visible = true;
                var alreadyCopied = false;
                for (var k = 0; k < neckLayer.groupItems.length; k++) {
                    if (neckLayer.groupItems[k].parent === neckLayer &&
                        neckLayer.groupItems[k].name === neckName) {
                        alreadyCopied = true;
                        break;
                    }
                }
                if (!alreadyCopied) {
                    var copy = neckband.duplicate(neckLayer, ElementPlacement.PLACEATEND);
                    copy.name = neckName;
                    copiedNeckbands++;
                }
            } catch (e) {
                errors.push("Neckband-" + size + ": " + e.message);
            }
        }
    } else {
        errors.push("Master group neckband-allsizes not found in patterns.");
    }

    var message = "Moved " + moved + " of " + plans.length + " named groups.\n" +
        "Copied " + copiedNeckbands + " neckbands. Master retained in patterns.";
    if (errors.length) {
        message += "\n\nFailed moves (first 10):\n" + errors.slice(0, 10).join("\n");
    }
    alert(message);
})();

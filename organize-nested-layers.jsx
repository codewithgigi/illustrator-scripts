
#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("Open an Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;
    var MIN_SIZE = 33;
    var MAX_SIZE = 44;

    var pieces = {};
    var notches = {};
    var sizeLayers = {};

    var movedPieces = 0;
    var mergedNotches = 0;
    var unmatchedNotches = [];

    function getSize(name) {
        var match = name.match(/(\d+)$/);
        if (!match) return null;

        var size = parseInt(match[1], 10);

        return (size >= MIN_SIZE && size <= MAX_SIZE)
            ? size : null;
    }

    function isNotch(name) {
        return /notch/i.test(name);
    }

    function getPieceName(name) {
        // Remove trailing size number.
        var result = name.replace(/[\s_-]*\d+$/, "");

        // Remove the word "notch" or "notches".
        result = result.replace(/notches?/ig, "");

        // Normalize separators.
        result = result.replace(/[\s_-]+/g, " ");

        return result.replace(/^\s+|\s+$/g, "").toLowerCase();
    }

    function collect(parent) {
        for (var i = parent.layers.length - 1; i >= 0; i--) {
            var layer = parent.layers[i];

            // Ignore previously organized size layers.
            if (/^Size-\d+$/i.test(layer.name)) {
                continue;
            }

            var size = getSize(layer.name);

            if (size !== null) {
                var key = size + "|" + getPieceName(layer.name);

                if (isNotch(layer.name)) {
                    if (!notches[key]) notches[key] = [];
                    notches[key].push(layer);
                } else {
                    if (!pieces[key]) pieces[key] = [];
                    pieces[key].push(layer);
                }
            } else {
                collect(layer);
            }
        }
    }

    function unlockTree(layer) {
        layer.locked = false;

        for (var i = 0; i < layer.layers.length; i++) {
            unlockTree(layer.layers[i]);
        }

        for (var j = 0; j < layer.pageItems.length; j++) {
            try {
                layer.pageItems[j].locked = false;
            } catch (e) {}
        }
    }

    function moveContents(source, destination) {
        unlockTree(source);

        // Move direct sublayers.
        while (source.layers.length > 0) {
            source.layers[0].move(
                destination,
                ElementPlacement.PLACEATEND
            );
        }

        // Move direct artwork.
        // pageItems can include nested artwork, so use
        // only items whose immediate parent is source.
        for (var i = source.pageItems.length - 1; i >= 0; i--) {
            var item = source.pageItems[i];

            if (item.parent === source) {
                try {
                    item.locked = false;
                    item.move(
                        destination,
                        ElementPlacement.PLACEATEND
                    );
                } catch (e) {}
            }
        }
    }

    // 1. Collect pattern pieces and notch layers.
    collect(doc);

    // 2. Create size layers.
    for (var size = MIN_SIZE; size <= MAX_SIZE; size++) {
        var name = "Size-" + size;
        var destination = null;

        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                destination = doc.layers[i];
                break;
            }
        }

        if (!destination) {
            destination = doc.layers.add();
            destination.name = name;
        }

        destination.locked = false;
        destination.visible = true;
        sizeLayers[size] = destination;
    }

    // 3. Move pattern piece layers into their size layers.
    var destinations = {};

    for (var key in pieces) {
        if (!pieces.hasOwnProperty(key)) continue;

        var size = parseInt(key.split("|")[0], 10);
        var list = pieces[key];

        for (var i = list.length - 1; i >= 0; i--) {
            var layer = list[i];

            unlockTree(layer);

            layer.move(
                sizeLayers[size],
                ElementPlacement.PLACEATEND
            );

            // First matching piece becomes the destination.
            if (!destinations[key]) {
                destinations[key] = layer;
            }

            movedPieces++;
        }
    }

    // 4. Merge notch artwork into matching pattern pieces.
    for (var key in notches) {
        if (!notches.hasOwnProperty(key)) continue;

        var notchList = notches[key];
        var target = destinations[key];

        if (!target) {
            for (var i = 0; i < notchList.length; i++) {
                unmatchedNotches.push(notchList[i].name);
            }
            continue;
        }

        for (var i = notchList.length - 1; i >= 0; i--) {
            var notchLayer = notchList[i];

            moveContents(notchLayer, target);

            // Only remove the empty notch layer.
            if (
                notchLayer.pageItems.length === 0 &&
                notchLayer.layers.length === 0
            ) {
                notchLayer.remove();
                mergedNotches++;
            }
        }
    }

    var message =
        "Pattern layers organized!\n\n" +
        "Pattern piece layers moved: " + movedPieces + "\n" +
        "Notch layers merged: " + mergedNotches;

    if (unmatchedNotches.length > 0) {
        message +=
            "\n\nUnmatched notch layers left unchanged:\n" +
            unmatchedNotches.join("\n");
    }

    alert(message);

})();

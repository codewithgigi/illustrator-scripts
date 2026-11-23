var doc = app.activeDocument;

// 1. Prompt for the ease/addition number (supports decimals like 0, 2, 2.7)
var easeInput = prompt("Enter the number (in inches) to add for 'Finished' measurements (e.g., 0 or 2.7):", "0");

if (easeInput !== null) {
    var easeVal = parseFloat(easeInput);

    if (isNaN(easeVal)) {
        alert("Please enter a valid number.");
    } else {
        // Helper to round cleanly (removes unnecessary trailing zeros like 35.0 -> 35)
        function formatNum(num) {
            return Math.round(num * 100) / 100;
        }

        var updatedFrames = 0;

        // 2. Recursive function to scan all TextFrames inside a layer or group
        function processTextInContainer(container, fullInches, fullCm, finInches, finCm) {
            if (container.textFrames) {
                for (var t = 0; t < container.textFrames.length; t++) {
                    var tf = container.textFrames[t];
                    if (tf.locked || tf.hidden) continue;

                    var originalText = tf.contents;
                    var newText = originalText;

                    var fullStr = fullInches + '" / ' + fullCm + " cm";
                    var finStr = finInches + '" / ' + finCm + " cm";


                    // Replace "full hip:" or "full bust:" (and anything after it on that same line)
                    newText = newText.replace(/(full\s+(?:hip|bust)\s*:)[^\r\n]*/gi, "$1 " + fullStr);

                    // Replace "finished hip:" or "finished bust:" (and anything after it on that same line)
                    newText = newText.replace(/(finished\s+(?:hip|bust)\s*:)[^\r\n]*/gi, "$1 " + finStr);

                    if (newText !== originalText) {
                        tf.contents = newText;
                        updatedFrames++;
                    }
                }
            }

            // Recurse into sublayers
            if (container.layers) {
                for (var s = 0; s < container.layers.length; s++) {
                    if (!container.layers[s].locked && container.layers[s].visible) {
                        processTextInContainer(container.layers[s], fullInches, fullCm, finInches, finCm);
                    }
                }
            }

            // Recurse into groups
            if (container.groupItems) {
                for (var g = 0; g < container.groupItems.length; g++) {
                    if (!container.groupItems[g].locked && !container.groupItems[g].hidden) {
                        processTextInContainer(container.groupItems[g], fullInches, fullCm, finInches, finCm);
                    }
                }
            }
        }

        // 3. Loop through all top-level "size-XX" layers
        for (var i = 0; i < doc.layers.length; i++) {
            var lyr = doc.layers[i];
            if (lyr.locked || !lyr.visible) continue;

            // Match layer names like "size-35" or "size-35.5"
            var match = lyr.name.match(/^size-(\d+(?:\.\d+)?)/i);
            if (match) {
                var sizeNum = parseFloat(match[1]);

                var fullIn = formatNum(sizeNum);
                var fullCm = formatNum(sizeNum * 2.54);

                var finIn = formatNum(sizeNum + easeVal);
                var finCm = formatNum((sizeNum + easeVal) * 2.54);

                processTextInContainer(lyr, fullIn, fullCm, finIn, finCm);
            }
        }

        alert("Done! Updated " + updatedFrames + " text frames across your size layers.");
    }
}

var doc = app.activeDocument;

// 1. Choose the destination folder for your PDFs
var exportFolder = Folder.selectDialog("Select a folder to save your size PDF(s):");

if (exportFolder) {
    // 2. Ask for a specific size OR leave blank for all sizes
    var targetSizeInput = prompt(
        "Enter a SINGLE size number to export (e.g., 35),\nor leave BLANK to export ALL sizes:",
        ""
    );

    if (targetSizeInput !== null) {
        // 3. Ask for an optional file prefix
        var filePrefix = prompt("Enter a prefix for the PDF filenames (or leave blank):", "");

        if (filePrefix !== null) {
            var allSizeLayers = [];
            var layersToExport = [];
            var originalVisibility = [];

            // Clean up user input in case they typed "size-35" instead of just "35"
            var cleanTarget = targetSizeInput.replace(/^\s+|\s+$/g, "").replace(/^size-/i, "");

            // 4. Record original visibility and collect matching size layers
            for (var i = 0; i < doc.layers.length; i++) {
                var lyr = doc.layers[i];
                originalVisibility.push(lyr.visible);

                if (lyr.name.toLowerCase().indexOf("size-") === 0) {
                    allSizeLayers.push(lyr);

                    // If blank, queue all sizes; otherwise only queue the requested size
                    if (cleanTarget === "" || lyr.name.toLowerCase() === ("size-" + cleanTarget).toLowerCase()) {
                        layersToExport.push(lyr);
                    }
                }
            }

            if (layersToExport.length === 0) {
                alert("Could not find a layer matching 'size-" + cleanTarget + "'.");
            } else {
                // 5. Configure PDF export options
                var pdfOpts = new PDFSaveOptions();
                pdfOpts.createAcrobatLayers = false;
                pdfOpts.saveMultipleArtboards = true;
                pdfOpts.artboardRange = "1-" + doc.artboards.length;
                pdfOpts.preserveEditability = false;
                pdfOpts.viewAfterSaving = false;

                var exportedCount = 0;

                // 6. Loop through only the queued size layer(s)
                for (var s = 0; s < layersToExport.length; s++) {
                    var activeSizeLayer = layersToExport[s];

                    // Hide ALL size-XX layers first so only the target size shows
                    for (var k = 0; k < allSizeLayers.length; k++) {
                        allSizeLayers[k].visible = false;
                    }

                    // Turn on ONLY the active size layer
                    activeSizeLayer.visible = true;

                    // Save PDF (overwrites the existing PDF for that size if it's already in the folder)
                    var fileName = filePrefix + activeSizeLayer.name + ".pdf";
                    var destFile = new File(exportFolder.fsName + "/" + fileName);

                    doc.saveAs(destFile, pdfOpts);
                    exportedCount++;
                }

                // 7. Restore all layers to their original visibility state
                for (var r = 0; r < doc.layers.length; r++) {
                    doc.layers[r].visible = originalVisibility[r];
                }

                alert("Done! Exported " + exportedCount + " PDF(s) to:\n" + exportFolder.fsName);
            }
        }
    }
}

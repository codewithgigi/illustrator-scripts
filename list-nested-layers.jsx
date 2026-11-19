
#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("Open your pattern file first.");
        return;
    }

    var doc = app.activeDocument;
    var output = [];

    function inspect(parent, depth) {
        for (var i = 0; i < parent.layers.length; i++) {
            var layer = parent.layers[i];

            var indent = "";
            for (var j = 0; j < depth; j++) {
                indent += "    ";
            }

            output.push(
                indent + layer.name +
                " | Artwork: " + layer.pageItems.length +
                " | Sublayers: " + layer.layers.length
            );

            inspect(layer, depth + 1);
        }
    }

    inspect(doc, 0);

    var file = new File(Folder.desktop + "/illustrator-layers.txt");

    if (file.open("w")) {
        file.write(output.join("\n"));
        file.close();

        alert(
            "Layer report saved to your Desktop:\n" +
            "illustrator-layers.txt\n\n" +
            "No changes were made to your document."
        );
    } else {
        alert("Unable to save layer report.");
    }
})();

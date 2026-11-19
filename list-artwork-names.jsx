
#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("Open your Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;
    var output = [];

    function inspect(layer) {
        output.push("\nLAYER: " + layer.name);

        for (var i = 0; i < layer.pageItems.length; i++) {
            var item = layer.pageItems[i];

            output.push(
                "Name: " + (item.name || "[unnamed]") +
                " | Type: " + item.typename +
                " | Parent: " + item.parent.typename
            );
        }
    }

    inspect(doc.layers.getByName("patterns"));
    inspect(doc.layers.getByName("notches"));

    var file = new File(
        Folder.desktop + "/illustrator-objects.txt"
    );

    if (file.open("w")) {
        file.write(output.join("\n"));
        file.close();

        alert(
            "Object report saved to Desktop:\n" +
            "illustrator-objects.txt"
        );
    } else {
        alert("Could not save report.");
    }
})();

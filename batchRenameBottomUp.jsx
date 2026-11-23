var doc = app.activeDocument;
var sel = doc.selection;

if (sel.length > 0) {
    var baseName = prompt("Enter base name (e.g., Path-35-):", "Path-35-");
    var startNum = parseInt(prompt("Enter starting number:", "35"), 10);

    if (baseName !== null && !isNaN(startNum)) {
        var count = 0;
        for (var i = sel.length - 1; i >= 0; i--) {
            var item = sel[i];
            var newName = baseName + (startNum + count);
            item.name = newName;

            // If it's a Compound Path, also rename its inner paths so they match
            if (item.typename === "CompoundPathItem") {
                for (var p = 0; p < item.pathItems.length; p++) {
                    item.pathItems[p].name = newName;
                }
            }
            count++;
        }
    }
} else {
    alert("Please select at least one object first.");
}

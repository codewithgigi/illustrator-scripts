// StylePattern_BW_Recursive.jsx
// Styles patterns in monochrome black using unique dashes based on layer suffixes.

function main() {
    if (app.documents.length === 0) {
        alert("Please open a document first.");
        return;
    }

    var doc = app.activeDocument;

    // MONOCHROME BLACK STYLES (Optimized for grayscale printing)
    // weight = thickness, dash = [dash, gap, dash, gap]
    var sizeStyles = {
        "xs":  { weight: 1.0, dash: [] },                  // Fine Solid Line
        "s":   { weight: 1.5, dash: [6, 4] },              // Standard Dashed
        "m":   { weight: 1.5, dash: [2, 3] },              // Dotted Line
        "l":   { weight: 2.0, dash: [12, 4, 2, 4] },       // Classic Dash-Dot
        "xl":  { weight: 2.0, dash: [14, 5] },             // Long Dash
        "xxl": { weight: 2.5, dash: [6, 3, 2, 3, 2, 3] }   // Heavy Dash-Dot-Dot
    };

    // Create a pure black RGB color object
    var blackColor = new RGBColor();
    blackColor.red = 0;
    blackColor.green = 0;
    blackColor.blue = 0;

    // Recursive function to style paths inside groups, compound paths, etc.
    function applyStyleToElement(element, style) {
        if (element.typename === "PathItem") {
            element.stroked = true;
            element.strokeWidth = style.weight;
            element.strokeColor = blackColor;
            element.strokeDashes = style.dash;
        } 
        else if (element.typename === "CompoundPathItem") {
            // Drill down into compound paths
            for (var p = 0; p < element.pathItems.length; p++) {
                applyStyleToElement(element.pathItems[p], style);
            }
        } 
        else if (element.typename === "GroupItem") {
            // Drill down into nested groups
            for (var g = 0; g < element.pageItems.length; g++) {
                applyStyleToElement(element.pageItems[g], style);
            }
        }
    }

    var stylizedCount = 0;

    // Loop through every layer in the document
    for (var l = 0; l < doc.layers.length; l++) {
        var currentLayer = doc.layers[l];
        var layerNameLower = currentLayer.name.toLowerCase();

        for (var sizeKey in sizeStyles) {
            if (sizeStyles.hasOwnProperty(sizeKey)) {
                var suffix = "-" + sizeKey;

                // Check if the layer name ends with the size suffix
                if (layerNameLower.indexOf(suffix, layerNameLower.length - suffix.length) !== -1) {
                    var currentStyle = sizeStyles[sizeKey];

                    // Process all top-level items inside this layer recursively
                    for (var i = 0; i < currentLayer.pageItems.length; i++) {
                        applyStyleToElement(currentLayer.pageItems[i], currentStyle);
                    }
                    
                    stylizedCount++;
                    break;
                }
            }
        }
    }

    if (stylizedCount > 0) {
        alert("Success! Stylized paths across " + stylizedCount + " size layers in high-contrast Black & White.");
    } else {
        alert("No layers changed. Please ensure your layer names end exactly with a hyphen and the size (e.g., 'back-xl', 'front-xs').");
    }
}

main();

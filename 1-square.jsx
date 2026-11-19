/* StyledStitch - Script 01: 2 x 2 inch test square
   Adds a calibration square to artboard 1 without modifying artboards.
   Adobe Illustrator ExtendScript (JSX).
*/
#target illustrator
(function () {
  if (app.documents.length === 0) {
    alert('Open your Illustrator document with artboards first.');
    return;
  }
  var doc = app.activeDocument;
  var P = 72;
  var board = doc.artboards[0].artboardRect;
  var left = board[0], top = board[1], right = board[2], bottom = board[3];
  var margin = 0.5 * P;
  var size = 2 * P;
  var inset = 0.25 * P;
  var squareLeft = left + margin + inset;
  var squareTop = top - margin - 0.45 * P;
  if (squareLeft + size > right - margin || squareTop - size < bottom + margin) {
    alert('Artboard 1 is too small for the test square at the planned location.');
    return;
  }
  var layer = doc.layers.add();
  layer.name = '01 - TEST SQUARE (2 x 2 in)';
  var black = new RGBColor();
  black.red = 0; black.green = 0; black.blue = 0;
  var square = layer.pathItems.rectangle(squareTop, squareLeft, size, size);
  square.name = '2 x 2 in calibration square';
  square.filled = false;
  square.stroked = true;
  square.strokeColor = black;
  square.strokeWidth = 0.5;
  var text = layer.textFrames.add();
  text.contents = '2 in x 2 in TEST SQUARE';
  text.position = [squareLeft, squareTop + 0.20 * P];
  text.textRange.characterAttributes.size = 8;
  text.textRange.characterAttributes.fillColor = black;
  doc.activeLayer = layer;
  alert('Added a 2 x 2 inch test square to Artboard 1.\nMeasure between the CENTERLINES of the printed square edges.\nArtboards were not changed.');
})();

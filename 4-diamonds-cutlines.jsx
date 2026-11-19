#target illustrator
/* StyledStitch: shared cut/fold lines and registration diamonds.
   Run on overlapping US Letter artboards (6.5in x 9in advances).
   Creates one NEW layer; never deletes or changes existing layers.
   IMPORTANT: Hide prior diamond/fold layers before exporting.
   Trim/fold the LEFT edge of incoming right-hand pages and the TOP
   edge of incoming lower pages, at the dashed shared-coordinate line.
*/
(function () {
    if (app.documents.length === 0) { alert('Open the tiled Illustrator document first.'); return; }
    var doc = app.activeDocument;
    var PT = 72, eps = 2;
    var W = 8.5*PT, H = 11*PT, STEP_X = 6.5*PT, STEP_Y = 9*PT;
    var M = .5*PT, DIAMOND = .25*PT;
    var ab = [], i, j;
    for (i=0; i<doc.artboards.length; i++) {
        var r=doc.artboards[i].artboardRect;
        if (Math.abs((r[2]-r[0])-W)>eps || Math.abs((r[1]-r[3])-H)>eps) {
            alert('Artboard '+(i+1)+' is not US Letter. Nothing changed.'); return;
        }
        ab.push({l:r[0],t:r[1],r:r[2],b:r[3]});
    }
    var layer=doc.layers.add();
    layer.name='07 - COORDINATED CUT FOLD + DIAMONDS';
    var ink=new GrayColor(); ink.gray=100;
    function line(x1,y1,x2,y2) {
        var p=layer.pathItems.add();
        p.setEntirePath([[x1,y1],[x2,y2]]);
        p.filled=false; p.stroked=true; p.strokeColor=ink;
        p.strokeWidth=.5; p.strokeDashes=[5,4];
        return p;
    }
    function diamond(cx,cy) {
        var h=DIAMOND/2;
        var p=layer.pathItems.add();
        p.setEntirePath([[cx,cy+h],[cx+h,cy],[cx,cy-h],[cx-h,cy]]);
        p.closed=true; p.filled=false; p.stroked=true;
        p.strokeColor=ink; p.strokeWidth=.75;
        return p;
    }
    var vertical=0, horizontal=0;
    for(i=0;i<ab.length;i++) for(j=i+1;j<ab.length;j++) {
        var a=ab[i], b=ab[j], dx=Math.abs(a.l-b.l), dy=Math.abs(a.t-b.t);
        if(dy<eps && Math.abs(dx-STEP_X)<eps) {
            var left=a.l<b.l?a:b, right=a.l<b.l?b:a;
            // Intersection of BOTH 0.5-inch printable regions.
            var xLo=Math.max(left.l+M,right.l+M);
            var xHi=Math.min(left.r-M,right.r-M);
            var yTop=Math.min(left.t-M,right.t-M);
            var yBot=Math.max(left.b+M,right.b+M);
            if(xHi>xLo && yTop>yBot) {
                var x=(xLo+xHi)/2, y=(yTop+yBot)/2;
                line(x,yTop,x,yBot); diamond(x,y); vertical++;
            }
        } else if(dx<eps && Math.abs(dy-STEP_Y)<eps) {
            var top=a.t>b.t?a:b, bottom=a.t>b.t?b:a;
            var yyTop=Math.min(top.t-M,bottom.t-M);
            var yyBot=Math.max(top.b+M,bottom.b+M);
            var xxLo=Math.max(top.l+M,bottom.l+M);
            var xxHi=Math.min(top.r-M,bottom.r-M);
            if(xxHi>xxLo && yyTop>yyBot) {
                var xx=(xxLo+xxHi)/2, yy=(yyTop+yyBot)/2;
                line(xxLo,yy,xxHi,yy); diamond(xx,yy); horizontal++;
            }
        }
    }
    if(vertical+horizontal===0) {
        layer.remove();
        alert('No neighboring artboards found at 6.5in / 9in spacing. Nothing created.');
    } else {
        alert('Created '+vertical+' vertical and '+horizontal+' horizontal shared cut/fold lines with one 1/4-inch diamond per line.\n\nHide ALL previous registration and fold layers before PDF export.\n\nTest-print two pages at 100% and verify alignment before using for a pattern.');
    }
})();

#target illustrator
(function () {
    if (app.documents.length === 0) { alert('Open an Illustrator document first.'); return; }
    var doc = app.activeDocument;
    function norm(s) { return s.toLowerCase().replace(/[\s_-]+/g, ''); }
    function findTop(name) {
        for (var i = 0; i < doc.layers.length; i++)
            if (norm(doc.layers[i].name) === norm(name)) return doc.layers[i];
        return null;
    }
    function findChild(parent, name) {
        for (var i = 0; i < parent.layers.length; i++)
            if (norm(parent.layers[i].name) === norm(name)) return parent.layers[i];
        return null;
    }
    var master = findTop('pattern information');
    if (!master) { alert('Cannot find top-level "pattern information" layer.'); return; }
    var masterLock = master.locked, masterVisible = master.visible;
    var total = 0, sizesDone = 0, errors = [], skipped = [];
    try {
        master.locked = false;
        master.visible = true;
        for (var size = 33; size <= 44; size++) {
            var label = 'Size-' + size;
            var dest = findTop(label);
            if (!dest) { errors.push(label + ': size layer missing'); continue; }
            var destLock = dest.locked, destVisible = dest.visible;
            try {
                dest.locked = false;
                dest.visible = true;
                var sub = findChild(dest, 'Pattern Information');
                if (sub && (sub.pageItems.length > 0 || sub.layers.length > 0)) {
                    skipped.push(label + ' (already populated)');
                    continue;
                }
                if (!sub) { sub = dest.layers.add(); sub.name = 'Pattern Information'; }
                sub.locked = false;
                sub.visible = true;
                var copied = 0;
                // Duplicate only direct children; duplicating a group includes its contents.
                // Snapshot references first, because Illustrator collections are live.
                var items = [];
                for (var i = 0; i < master.pageItems.length; i++) {
                    if (master.pageItems[i].parent === master) items.push(master.pageItems[i]);
                }
                for (var j = items.length - 1; j >= 0; j--) {
                    var item = items[j];
                    var locked = item.locked, hidden = item.hidden;
                    try {
                        item.locked = false;
                        item.hidden = false;
                        item.duplicate(sub, ElementPlacement.PLACEATEND);
                        copied++;
                    } catch (e) {
                        errors.push(label + ' / ' + (item.name || item.typename) + ': ' + e);
                    } finally {
                        try { item.locked = locked; item.hidden = hidden; } catch (ignore) {}
                    }
                }
                for (var k = master.layers.length - 1; k >= 0; k--) {
                    try { master.layers[k].duplicate(sub, ElementPlacement.PLACEATEND); copied++; }
                    catch (e2) { errors.push(label + ' / nested layer: ' + e2); }
                }
                if (copied > 0) { sizesDone++; total += copied; }
                else errors.push(label + ': no artwork copied');
            } catch (e3) { errors.push(label + ': ' + e3); }
            finally {
                try { dest.locked = destLock; dest.visible = destVisible; } catch (ignore2) {}
            }
        }
    } finally {
        master.locked = masterLock;
        master.visible = masterVisible;
    }
    alert('Pattern information copied to ' + sizesDone + ' of 12 size layers.\n' +
          'Total top-level objects copied: ' + total +
          (skipped.length ? '\nSkipped: ' + skipped.join(', ') : '') +
          (errors.length ? '\nIssues:\n' + errors.join('\n').substring(0, 2000) : ''));
})();

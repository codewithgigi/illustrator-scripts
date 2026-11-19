#target illustrator
(function () {
    if (app.documents.length === 0) { alert('Open your Illustrator document first.'); return; }
    var doc = app.activeDocument;
    var master = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name.toLowerCase().replace(/[\s_-]/g, '') === 'patternmarkings') {
            master = doc.layers[i]; break;
        }
    }
    if (!master) { alert('Cannot find the pattern-markings master layer.'); return; }
    var count = 0, errors = [], skipped = [], originalMasterLock = master.locked;
    var states = [];
    try {
        master.locked = false;
        for (var size = 33; size <= 44; size++) {
            var dest = null, label = 'Size-' + size;
            for (var a = 0; a < doc.layers.length; a++) {
                if (doc.layers[a].name.toLowerCase() === label.toLowerCase()) { dest = doc.layers[a]; break; }
            }
            if (!dest) { errors.push(label + ': missing size layer'); continue; }
            var oldLock = dest.locked, oldVisible = dest.visible;
            states.push({layer: dest, locked: oldLock, visible: oldVisible});
            try {
                dest.locked = false; dest.visible = true;
                var existing = null;
                for (var b = 0; b < dest.layers.length; b++) {
                    if (dest.layers[b].name.toLowerCase().replace(/[\s_-]/g, '') === 'patternmarkings') { existing = dest.layers[b]; break; }
                }
                if (existing) { skipped.push(label); continue; }
                var sub = dest.layers.add(); sub.name = 'Pattern Markings';
                // Copy direct artwork only. pageItems also includes descendants of groups.
                var copied = 0;
                for (var c = master.pageItems.length - 1; c >= 0; c--) {
                    var item = master.pageItems[c];
                    if (item.parent !== master) continue;
                    var itemLocked = item.locked, itemHidden = item.hidden;
                    try {
                        item.locked = false; item.hidden = false;
                        item.duplicate(sub, ElementPlacement.PLACEATEND);
                        copied++;
                    } catch (e) {
                        errors.push(label + ' / ' + (item.name || item.typename) + ': ' + e);
                    } finally {
                        try { item.locked = itemLocked; item.hidden = itemHidden; } catch (ignore) {}
                    }
                }
                // Copy any nested layers, if present.
                for (var d = master.layers.length - 1; d >= 0; d--) {
                    try { master.layers[d].duplicate(sub, ElementPlacement.PLACEATEND); copied++; }
                    catch (e2) { errors.push(label + ' / nested layer: ' + e2); }
                }
                if (copied === 0) {
                    try { sub.remove(); } catch (ignore2) {}
                    errors.push(label + ': no objects copied');
                } else count++;
            } catch (err) { errors.push(label + ': ' + err); }
        }
    } finally {
        master.locked = originalMasterLock;
        for (var s = 0; s < states.length; s++) {
            try { states[s].layer.locked = states[s].locked; states[s].layer.visible = states[s].visible; } catch (ignore3) {}
        }
    }
    alert('Markings copied into ' + count + ' size layers.' +
          (skipped.length ? '\nAlready had markings: ' + skipped.join(', ') : '') +
          (errors.length ? '\nProblems:\n' + errors.join('\n').substring(0, 1400) : ''));
})();

#target illustrator
/* StyledStitch universal measurement labels - use on a backup. */
(function () {
 if (!app.documents.length) { alert('Open a document first.'); return; }
 var doc=app.activeDocument, FIRST=33, LAST=44;
 function trim(x){return String(x).replace(/^\s+|\s+$/g,'');}
 function num(x){x=trim(x);if(!x||!/^[-+]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(x))return null;return parseFloat(x);}
 function fmt(x){return String(Math.round(x*100)/100);}
 function norm(x){return String(x).toLowerCase().replace(/[^a-z0-9]/g,'');}
 function findLayer(parent,name){for(var i=0;i<parent.layers.length;i++)if(norm(parent.layers[i].name)===norm(name))return parent.layers[i];return null;}
 var w=new Window('dialog','StyledStitch - Measurement Labels');w.orientation='column';w.alignChildren='fill';
 w.add('statictext',undefined,'Choose measurements and ease in inches:');
 function option(label,defaultChecked,defaultEase){var g=w.add('group');var c=g.add('checkbox',undefined,label);c.value=defaultChecked;g.add('statictext',undefined,'Ease:');var e=g.add('edittext',undefined,defaultEase);e.characters=8;return {check:c,ease:e};}
 var b=option('Full bust',true,'2.7'),wa=option('Waist',false,'1'),h=option('Full hip',false,'2');
 var dg=w.add('group'),dart=dg.add('checkbox',undefined,'Waist darts optional');dg.add('statictext',undefined,'Total dart intake:');var di=dg.add('edittext',undefined,'2');di.characters=8;
 w.add('statictext',undefined,'Which measurement does the Size-XX layer number represent?');
 var primary=w.add('dropdownlist',undefined,['Full bust','Waist','Full hip']);primary.selection=0;
 w.add('statictext',undefined,'For selected OTHER measurements, enter 12 body values (sizes 33-44):');
 function mapping(label){var g=w.add('group');g.add('statictext',undefined,label);var t=g.add('edittext',undefined,'');t.characters=48;return t;}
 var bm=mapping('Bust:'),wm=mapping('Waist:'),hm=mapping('Hip:');
 var buttons=w.add('group');buttons.alignment='right';buttons.add('button',undefined,'Cancel',{name:'cancel'});buttons.add('button',undefined,'Apply',{name:'ok'});
 if(w.show()!=1)return;
 var c={useB:b.check.value,useW:wa.check.value,useH:h.check.value,be:num(b.ease.text),we:num(wa.ease.text),he:num(h.ease.text),dart:dart.value,di:num(di.text),primary:primary.selection.index};
 if(!c.useB&&!c.useW&&!c.useH){alert('Select a measurement.');return;}
 if((c.useB&&c.be===null)||(c.useW&&c.we===null)||(c.useH&&c.he===null)||(c.useW&&c.dart&&(c.di===null||c.di<0))){alert('Enter valid ease/dart values.');return;}
 if(c.dart&&!c.useW){alert('Select Waist to use dart intake.');return;}
 function parseMap(raw,needed,label){if(!needed)return null;var a=raw.split(','),v=[];if(a.length!==12){alert(label+': provide 12 values, one for each size 33-44.');return false;}for(var i=0;i<12;i++){var n=num(a[i]);if(n===null||n<=0){alert(label+': invalid value for Size-'+(33+i));return false;}v.push(n);}return v;}
 c.bmap=parseMap(bm.text,c.useB&&c.primary!==0,'Bust');if(c.bmap===false)return;
 c.wmap=parseMap(wm.text,c.useW&&c.primary!==1,'Waist');if(c.wmap===false)return;
 c.hmap=parseMap(hm.text,c.useH&&c.primary!==2,'Hip');if(c.hmap===false)return;
 function value(which,size){return c.primary===which?size:(which===0?c.bmap:which===1?c.wmap:c.hmap)[size-FIRST];}
 function lines(size){var a=[];if(c.useB){var n=value(0,size);a.push('FULL BUST: '+fmt(n)+'"');a.push('FINISHED BUST: '+fmt(n+c.be)+'"');}
 if(c.useW){var n=value(1,size),f=n+c.we;a.push('WAIST: '+fmt(n)+'"');a.push('FINISHED WAIST: '+fmt(f)+'"'+(c.dart?' ('+fmt(f+c.di)+'" WITHOUT DARTS)':''));}
 if(c.useH){var n=value(2,size);a.push('FULL HIP: '+fmt(n)+'"');a.push('FINISHED HIP: '+fmt(n+c.he)+'"');}return a;}
 // The prior working script matched SS-<digits> anywhere within a text line.
 var id=/SS-\d+/i;
 function isGenerated(line){return /^\s*(FULL BUST|FINISHED BUST|WAIST|FINISHED WAIST|FULL HIP|FINISHED HIP)\s*:/i.test(line);}
 function transform(old,size){if(!id.test(old))return null;var sep=old.indexOf('\r')>=0?'\r':'\n',src=old.split(/\r\n|\r|\n/),clean=[],at=-1;
 for(var i=0;i<src.length;i++){var line=src[i];if(isGenerated(line))continue;
 // Also remove combined measurement lines from the earlier universal version.
 if(/^\s*(FULL BUST|WAIST|FULL HIP)\s*:/i.test(line))continue;
 clean.push(line);if(at<0&&id.test(line))at=clean.length-1;}
 if(at<0)return null;var ins=lines(size);for(var j=0;j<ins.length;j++)clean.splice(at+1+j,0,ins[j]);return clean.join(sep);}
 var changed=0,matched=0,errors=[],missing=[],unchanged=0;
 for(var s=FIRST;s<=LAST;s++){
  var layer=findLayer(doc,'Size-'+s);if(!layer){missing.push('Size-'+s);continue;}
  var marks=findLayer(layer,'Pattern Markings');if(!marks){missing.push('Size-'+s+' / Pattern Markings');continue;}
  var oldL=layer.locked,oldM=marks.locked,oldV=layer.visible,oldMV=marks.visible;
  try{
   layer.locked=false;marks.locked=false;layer.visible=true;marks.visible=true;
   // Same collection access as the previously working script.
   for(var i=0;i<marks.textFrames.length;i++){
    var tf=marks.textFrames[i];try{
     var old=tf.contents,updated=transform(old,s);if(updated===null)continue;matched++;
     var oldLock=tf.locked,oldHidden=tf.hidden;
     tf.locked=false;tf.hidden=false;
     try{if(updated!==old){tf.contents=updated;changed++;}else unchanged++;}
     finally{tf.hidden=oldHidden;tf.locked=oldLock;}
    }catch(e){errors.push('Size-'+s+' text '+i+': '+e);}
   }
  }catch(e){errors.push('Size-'+s+': '+e);}
  finally{marks.locked=oldM;layer.locked=oldL;marks.visible=oldMV;layer.visible=oldV;}
 }
 alert('Measurement labels updated: '+changed+'\nMatching SS labels found: '+matched+'\nUnchanged: '+unchanged+'\n'+(missing.length?'\nMissing layers:\n'+missing.join('\n'):'')+(errors.length?'\nErrors:\n'+errors.join('\n').substring(0,1600):'')+(matched===0?'\nNo matching editable text found. Check that labels are in Pattern Markings and contain SS- followed by digits.':''));
})();

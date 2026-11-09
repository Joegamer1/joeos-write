'use strict';
let bridge,meta={},selected=null,renderToken=0,dirtyFormat=new Set(),pending=Promise.resolve(),renderPending=Promise.resolve();
const $=id=>document.getElementById(id);
const localFontFaces=new Map();
let activeDocumentKey=null;
const call=(name,...args)=>new Promise(resolve=>bridge[name](...args,resolve));
function queue(name,...args){pending=pending.then(()=>call(name,...args));return pending;}
function notify(text){$('banner').textContent=text;}
function zoom(){const page=$('render').querySelector('section.docx');$('render').style.zoom=$('zoom').value==='Fit width'&&page?Math.min(1,($('scroll').clientWidth-56)/page.offsetWidth):parseInt($('zoom').value)/100||1;correctWordSpacing($('render'));}
function updateStatus(raw){const s=JSON.parse(raw);$('filename').textContent=s.name;$('saved').textContent=s.dirty?'UNSAVED CHANGES':'SAVED LOCALLY';$('summary').textContent=s.summary;$('changes').textContent=s.changes||'No changes';$('undo').disabled=!s.undo;$('redo').disabled=!s.redo;}
function choose(el){
 selected=el.dataset.joeosId;
 const m=meta[selected];$('textvalue').value=m?.text??el.textContent;$('textvalue').disabled=!m?.editable;$('applytext').disabled=!m?.editable;
 showPageMeasurements(el.closest('section.docx'));
 $('selectionHint').textContent=m?.editable?'Type at the cursor. Enter starts a paragraph; Shift+Enter inserts a line break.':'This element is retained unchanged. It is not editable in this version.';
 const style=getComputedStyle(el);$('font').value=(el.dataset.joeosDocumentFont||style.fontFamily).split(',')[0].replaceAll('"','');$('size').value=Number(el.dataset.joeosDocumentSize||parseFloat(style.fontSize)*.75).toFixed(2).replace(/\.?0+$/,'');
 const rgb=style.color.match(/\d+/g);if(rgb?.length>=3)$('color').value='#'+rgb.slice(0,3).map(x=>(+x).toString(16).padStart(2,'0')).join('');
 $('bold').classList.toggle('active',+style.fontWeight>=600||style.fontWeight==='bold');$('italic').classList.toggle('active',style.fontStyle==='italic');$('underline').classList.toggle('active',style.textDecorationLine.includes('underline'));dirtyFormat.clear();
 for(const id of ['bold','italic','underline','font','size','color','applyformat'])$(id).disabled=!m?.editable;
}
const textEditVersions=new Map();
async function edit(id,text,before=null,after=null,fromTyping=false){
 const version=(textEditVersions.get(id)||0)+1; textEditVersions.set(id,version);
 const ok=before?await queue('typeText',id,text,JSON.stringify(before),JSON.stringify(after)):await queue('editText',id,text);
 if(ok&&textEditVersions.get(id)===version&&meta[id]){
  meta[id].text=text;
  // The browser already applied a native keystroke. Replacing that live DOM
  // on acknowledgment loses its caret and can overwrite later keystrokes.
  if(!fromTyping)document.querySelectorAll('[data-joeos-id="'+id+'"]').forEach(n=>{if(!n.dataset.joeosFragment&&n.textContent!==text)n.textContent=text;});
 }
 return ok;
}
// Reconcile the prepared layout without clearing the live editing surface.
// Existing page/paragraph/text nodes survive when their structure is unchanged.
function reconcilePreview(target,source){
 if(target.nodeType!==1){if(target.data!==source.data)target.data=source.data;return;}
 for(const attr of [...target.attributes])if(!source.hasAttribute(attr.name))target.removeAttribute(attr.name);
 for(const attr of source.attributes)if(target.getAttribute(attr.name)!==attr.value)target.setAttribute(attr.name,attr.value);
 const wanted=[...source.childNodes];
 for(let i=0;i<wanted.length;i++){
  const fresh=wanted[i],old=target.childNodes[i];
  if(!old){target.appendChild(fresh.cloneNode(true));continue;}
  if(old.nodeType!==fresh.nodeType||old.nodeName!==fresh.nodeName)target.replaceChild(fresh.cloneNode(true),old);
  else reconcilePreview(old,fresh);
 }
 while(target.childNodes.length>wanted.length)target.lastChild.remove();
}
async function receive(raw){
 const payload=JSON.parse(raw),sameDocument=payload.documentKey===activeDocumentKey;
 const restoreBefore=sameDocument?window.typingBookmark?.():null;typingBusy=true;
 if(!sameDocument){typingMarks={};typingActions=[];}activeDocumentKey=payload.documentKey;
 const liveRoot=$('render'),liveStyles=$('docstyles');
 const renderRoot=sameDocument?liveRoot.cloneNode(false):liveRoot,renderStyles=sameDocument?document.createElement('div'):liveStyles;
 if(sameDocument){renderRoot.removeAttribute('contenteditable');renderRoot.style.position='fixed';renderRoot.style.left='-100000px';renderRoot.style.top='0';renderRoot.style.width=liveRoot.clientWidth+'px';renderRoot.style.visibility='hidden';renderRoot.setAttribute('aria-hidden','true');document.body.append(renderRoot,renderStyles);}
 const token=++renderToken;if(!sameDocument)document.body.classList.add('busy');selected=null;meta=Object.fromEntries(payload.texts.map(x=>[x.id,x]));
 try{
 const fontLoads=await Promise.all((payload.localFontFaces||[]).map(async face=>{
  const key=JSON.stringify([face.family,face.weight,face.style]);
  try{
   if(!localFontFaces.has(key)){
    const font=await new FontFace(face.family,`url(data:font/ttf;base64,${face.data})`,{weight:face.weight,style:face.style}).load();
    document.fonts.add(font);localFontFaces.set(key,font);
   }
   return {family:face.family,sourceFamily:face.sourceFamily,weight:face.weight,style:face.style,loaded:true};
  }catch(e){return {family:face.family,loaded:false,error:String(e)};}
 }));
 const failedFonts=[...new Set(fontLoads.filter(f=>!f.loaded).map(f=>f.family))];
 if(failedFonts.length)payload.notice+=' Could not load local fonts: '+failedFonts.join(', ')+'. Display may use substitutes.';
 const bytes=Uint8Array.from(atob(payload.data),c=>c.charCodeAt(0));
 await docx.renderAsync(bytes,renderRoot,renderStyles,{renderAltChunks:false,renderChanges:false,renderComments:false,ignoreLastRenderedPageBreak:false,ignoreWidth:false,ignoreHeight:false,ignoreFonts:false,renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true,useBase64URL:true,experimental:true,debug:false});
 if(token!==renderToken)return;
 applyDocumentTypography(renderRoot,payload.layout);
 applyDocumentLayout(renderRoot,payload.layout);
 for(const el of renderRoot.querySelectorAll('[data-joeos-id]'))el.removeAttribute('contenteditable');
 await document.fonts.ready;correctWordSpacing(renderRoot);refreshJoeosTabs(renderRoot);
 window.paginationReport=paginatePreview(renderRoot,payload.layout);
 applyDocumentLayout(renderRoot,payload.layout,true);correctWordSpacing(renderRoot);refreshJoeosTabs(renderRoot);
 if(sameDocument){
  const scrollTop=$('scroll').scrollTop,scrollLeft=$('scroll').scrollLeft;
  // Only children are reconciled: root focus, dimensions and scroll container stay live.
  const prepared=renderRoot.cloneNode(false);prepared.replaceChildren(...renderRoot.childNodes);
  for(const attr of [...prepared.attributes])prepared.removeAttribute(attr.name);
  for(const attr of liveRoot.attributes)prepared.setAttribute(attr.name,attr.value);
  reconcilePreview(liveRoot,prepared);
  if(liveStyles.innerHTML!==renderStyles.innerHTML)liveStyles.replaceChildren(...renderStyles.childNodes);
  renderRoot.remove();renderStyles.remove();$('scroll').scrollTop=scrollTop;$('scroll').scrollLeft=scrollLeft;
 }
 const fontChoices=[...new Set([...payload.fonts,...payload.documentFonts])].sort((a,b)=>a.localeCompare(b));
 for(const id of ['font'])$(id).replaceChildren(...fontChoices.map(f=>{const o=document.createElement('option');o.value=f;o.textContent=f+(payload.fonts.includes(f)?'':' (not installed)');return o;}));
 notify(payload.notice);$('textvalue').value='';$('selectionHint').textContent='Click text on the page to edit it directly. Existing formatting stays attached.';
 installTyping(payload,restoreBefore);
 await document.fonts.ready;zoom();showPageMeasurements($('render').querySelector('section.docx'));const overfull=[...document.querySelectorAll('section.docx')].some(p=>parseFloat(getComputedStyle(p).height)>parseFloat(getComputedStyle(p).minHeight)+1);if(overfull)notify(payload.notice+' Some content exceeds the paper height; some oversized or complex objects cannot yet be divided across pages.');await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));bridge.rendered(document.querySelectorAll('section.docx').length,document.querySelectorAll('[data-joeos-id]').length);
 const ctx=document.createElement('canvas').getContext('2d'),fontWidths={};
 for(const name of new Set(['Aptos','Aptos Serif','Arial','Times New Roman','Noto Serif','Courier New',...(payload.documentFonts||[]),...(payload.localFontFaces||[]).map(f=>f.family)])){
  ctx.font=`32px ${JSON.stringify(name)}`;fontWidths[name]=ctx.measureText('Hamburgefontsiv AV fi 0123456789').width;
 }
 const fontStyleProbes={};
 for(const face of fontLoads.filter(f=>f.loaded)){
  ctx.font=`${face.style} ${face.weight} 32px ${JSON.stringify(face.family)}`;
  fontStyleProbes[`${face.family}|${face.weight}|${face.style}`]={latin:ctx.measureText('Hamburgefontsiv AV fi 0123456789').width,thai:ctx.measureText('ภาษาไทย กำลังทดสอบ สระและวรรณยุกต์').width,narrow:ctx.measureText('iiiiiiii').width,wide:ctx.measureText('WWWWWWWW').width};
 }
 bridge.renderDiagnostics(JSON.stringify({fontWidths,fontStyleProbes,fontResolutions:payload.fontResolutions,localFontLoads:fontLoads,fonts:[...new Set([...document.querySelectorAll('[data-joeos-id]')].map(e=>getComputedStyle(e).fontFamily))],
  sizes:[...new Set([...document.querySelectorAll('[data-joeos-id]')].map(e=>getComputedStyle(e).fontSize))],
  zoom:getComputedStyle($('render')).zoom,referenceControlsPresent:!!$('appearance')}));
 }catch(e){notify('Page display could not be completed. The original package is still preserved. '+e.message);bridge.renderFailed(String(e));}
 finally{if(sameDocument){renderRoot.remove();renderStyles.remove();}document.body.classList.remove('busy');window.typingRendered?.();}
}
new QWebChannel(qt.webChannelTransport,channel=>{
 bridge=channel.objects.bridge;bridge.documentChanged.connect(raw=>{renderPending=renderPending.then(()=>receive(raw));});bridge.statusChanged.connect(updateStatus);bridge.message.connect(notify);
 for(const [id,method] of [['new','newDocument'],['open','openDocument'],['save','saveDocument'],['saveas','saveAs'],['undo','undo'],['redo','redo'],['help','help']])$(id).onclick=()=>queue(method);
 $('reflow').onclick=()=>queue('refreshPreview');
 $('add').onclick=()=>queue('addParagraph',selected||'');
 $('applytext').onclick=async()=>{if(selected&&await edit(selected,$('textvalue').value)){notify('Text updated; other document elements are retained.');queue('refreshPreview');}};
 for(const id of ['bold','italic','underline']){$(id).onmousedown=e=>e.preventDefault();$(id).onclick=()=>setTypingFormat(id,!$(id).classList.contains('active'));}
 for(const id of ['font','size','color'])$(id).addEventListener('input',()=>dirtyFormat.add(id));
 $('applyformat').onclick=async()=>{if(!selected)return;const id=selected,values=[...dirtyFormat].map(key=>[key,JSON.stringify(key==='size'?+$('size').value:$(key).value)]);for(const [key,value] of values)setTypingFormat(key,JSON.parse(value));};
 $('zoom').onchange=zoom;window.addEventListener('resize',zoom);
 $('guides').onchange=()=>document.body.classList.toggle('margin-guides',$('guides').checked);
 document.addEventListener('click',e=>{if(e.target.closest('#render a'))e.preventDefault();});
 document.addEventListener('keydown',e=>{if(e.ctrlKey||e.metaKey){const k=e.key.toLowerCase();if(e.defaultPrevented)return;const methods={s:e.shiftKey?'saveAs':'saveDocument',o:'openDocument',n:'newDocument',z:e.shiftKey?'redo':'undo',y:'redo'};if(methods[k]){e.preventDefault();queue(methods[k]);}}});
 bridge.ready();
});

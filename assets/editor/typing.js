'use strict';
// A single native editing surface. OOXML run IDs are bookkeeping, never focus
// targets. Native cursor navigation, selection, clipboard and IME stay native.
let typingBusy=false, typingActions=[], typingMarks={}, typingComposing=false, typingReflowTimer, nativeBefore=null;
const writing=()=>document.getElementById('render');
const textElement=node=>(node?.nodeType===3?node.parentElement:node)?.closest?.('[data-joeos-id]');

function typingPoint(node,offset,bias='end'){
 const root=writing();if(!node||!(node===root||root.contains(node)))return null;
 let el=textElement(node),pos=0;
 if(el){const r=document.createRange();r.selectNodeContents(el);r.setEnd(node,offset);pos=r.toString().length;}
 else{
  const els=[...root.querySelectorAll('[data-joeos-id]')];
  for(const candidate of els){const r=document.createRange();r.selectNodeContents(candidate);const compare=r.comparePoint(node,offset);if(compare<=0){if(!(compare<0&&bias==='start'&&el)){el=candidate;pos=0;}break;}el=candidate;pos=candidate.textContent.length;}
 }
 if(!el)return null;
 if(el.dataset.caretPlaceholder&&el.textContent.startsWith('\u200b'))pos=Math.max(0,pos-1);
 const original=meta[el.dataset.joeosId]?.text||'';
 // Imports may trim insignificant XML whitespace for display.
 if(el.dataset.joeosSpace!=='preserve'&&original!==el.textContent){
  const trimmed=original.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g,'');
  if(trimmed===el.textContent)pos+=original.length-original.replace(/^[ \t\r\n]+/,'').length;
 }
 if(el.dataset.joeosFragment){
  for(const other of root.querySelectorAll(`[data-joeos-id="${el.dataset.joeosId}"]`)){
   if(other===el)break;pos+=other.textContent.length;
  }
 }
 return {id:el.dataset.joeosId,offset:pos};
}
function typingRange(range){
 const s=window.getSelection();range=range||(s.rangeCount?s.getRangeAt(0):null);
 if(!range)return null;const collapsed=range.startContainer===range.endContainer&&range.startOffset===range.endOffset;
 const start=typingPoint(range.startContainer,range.startOffset,collapsed?'end':'start'),end=typingPoint(range.endContainer,range.endOffset);
 return start&&end?{start,end}:null;
}
function typingBookmark(){return writing().contains(document.activeElement)?typingRange():null;}
function domPoint(point){
 const elements=[...writing().querySelectorAll(`[data-joeos-id="${point.id}"]`)];let el=elements[0];if(!el)return null;
 let offset=point.offset;
 if(el.dataset.joeosFragment)for(const candidate of elements){el=candidate;if(offset<=el.textContent.length)break;offset-=el.textContent.length;}
 const original=meta[point.id]?.text||'';
 if(el.dataset.joeosSpace!=='preserve'&&original.trim()===el.textContent)offset-=original.length-original.replace(/^[ \t\r\n]+/,'').length;
 if(el.dataset.caretPlaceholder&&el.textContent.startsWith('\u200b'))offset++;
 offset=Math.max(0,Math.min(offset,el.textContent.length));
 const walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT);let node,last;
 while(node=walker.nextNode()){last=node;if(offset<=node.length)return [node,offset];offset-=node.length;}
 return last?[last,last.length]:[el,0];
}
function restoreTyping(bookmark){
 if(!bookmark)return;
 for(const point of [bookmark.start||bookmark,bookmark.end||bookmark]){
  const el=writing().querySelector(`[data-joeos-id="${point.id}"]`);
  if(el&&!el.textContent){el.dataset.caretPlaceholder='true';el.textContent='\u200b';}
 }
 const start=domPoint(bookmark.start||bookmark),end=domPoint(bookmark.end||bookmark);
 if(!start||!end)return;
 const r=document.createRange();r.setStart(...start);r.setEnd(...end);
 writing().focus({preventScroll:true});const s=window.getSelection();s.removeAllRanges();s.addRange(r);
 const current=textElement(start[0]);if(current)choose(current);
 textElement(start[0])?.scrollIntoView({block:'nearest',inline:'nearest'});
}
async function commitNativeText(){
 if(typingComposing)return;
 const seen=new Set();let changed=false;
 for(const el of writing().querySelectorAll('[data-joeos-id]')){
  const id=el.dataset.joeosId,m=meta[id];if(!m?.editable||seen.has(id))continue;seen.add(id);
  const group=el.dataset.joeosFragment?[...writing().querySelectorAll(`[data-joeos-id="${id}"]`)]:[el];
  const baseline=group.map(e=>e.dataset.typingBaseline).join(''),text=group.map(e=>e.dataset.caretPlaceholder?e.textContent.replace(/^\u200b/,''):e.textContent).join('');
  if(text!==baseline){
   changed=true;for(const e of group){e.dataset.typingBaseline=e.dataset.caretPlaceholder?e.textContent.replace(/^\u200b/,''):e.textContent;e.dataset.joeosSpace='preserve';}
   if(!await edit(id,text,nativeBefore,typingRange(),true)){if(el.textContent===text){el.textContent=m.text;el.dataset.typingBaseline=m.text;}}
  }
 }
 if(changed){clearTimeout(typingReflowTimer);typingReflowTimer=setTimeout(()=>{if(!typingBusy&&!typingComposing&&writing().contains(document.activeElement))queue('refreshPreview');},900);}
}
async function rangeAction(text,action='replace',marks=null,explicit=null){
 if(typingBusy){typingActions.push({text,action,marks});return;}
 const range=explicit||typingRange();if(!range)return;
 clearTimeout(typingReflowTimer);
 typingBusy=true;
 const appliedMarks=marks||typingMarks;typingMarks={};
 await commitNativeText();
 const ok=await queue('editRange',JSON.stringify({...range,text,action,marks:appliedMarks}));
 if(!ok){typingBusy=false;typingActions=[];return;}
}
function typingRendered(){
 typingBusy=false;
 if(typingActions.length){typingBusy=true;const next=typingActions.shift();
  // Keystrokes arriving during layout stay ordered, but need not each rebuild a transaction.
  while(typeof next.text==='string'&&!/[\t\n\v\f]/.test(next.text)&&typingActions.length){
   const following=typingActions[0];
   if(following.action!==next.action||JSON.stringify(following.marks)!==JSON.stringify(next.marks)||typeof following.text!=='string'||/[\t\n\v\f]/.test(following.text))break;
   next.text+=typingActions.shift().text;
  }
  setTimeout(()=>{
  typingBusy=false;
  if(next.deleteType){
   const s=getSelection();if(s.isCollapsed)s.modify('extend',next.deleteType.includes('Backward')?'backward':'forward',next.deleteType.includes('Word')?'word':'character');
   rangeAction('');
  }else if(next.command)queue(next.command);
  else rangeAction(next.text,next.action,next.marks);
 },0);}
}
function setTypingFormat(key,value){
 const range=typingRange();if(!range)return;
 if(range.start.id===range.end.id&&range.start.offset===range.end.offset){typingMarks[key]=value;$(key)?.classList.toggle('active',!!value);}
 else rangeAction('', 'format', {[key]:value},range);
}
function tableTab(backward){
 const s=getSelection(),el=s.anchorNode?.nodeType===3?s.anchorNode.parentElement:s.anchorNode;
 const cell=el?.closest?.('td,th');if(!cell)return false;
 const cells=[...cell.closest('table').querySelectorAll('td,th')],next=cells[cells.indexOf(cell)+(backward?-1:1)];
 const span=next?.querySelector('[data-joeos-id]');
 if(span){const last=[...next.querySelectorAll('[data-joeos-id]')].at(-1);restoreTyping({start:{id:span.dataset.joeosId,offset:0},end:{id:last.dataset.joeosId,offset:last.textContent.length}});}
 else notify('End of table. Adding table rows with Tab is not supported yet.');
 return true;
}
function installTyping(payload,bookmark){
 const root=writing();
 for(const entry of payload.emptyParagraphs||[]){
  const p=root.querySelector(`[data-joeos-paragraph="${entry.paragraph}"]`);if(!p||p.querySelector('[data-joeos-id]'))continue;
  const span=document.createElement('span');span.dataset.joeosId=entry.id;p.appendChild(span);meta[entry.id]=entry;
 }
 root.contentEditable='true';root.setAttribute('role','textbox');root.setAttribute('aria-label','Document');root.setAttribute('aria-multiline','true');root.spellcheck=true;
 for(const el of root.querySelectorAll('[data-joeos-id]')){
  const m=meta[el.dataset.joeosId];el.removeAttribute('tabindex');el.removeAttribute('contenteditable');el.dataset.typingBaseline=el.textContent;
  if(!m?.editable)el.contentEditable='false';
 }
 for(const p of root.querySelectorAll('p'))p.classList.toggle('joeos-empty-line',!p.textContent&&!p.querySelector('img,svg,canvas'));
 for(const el of root.querySelectorAll('img,svg,canvas'))el.contentEditable='false';
 root.oninput=()=>commitNativeText();
 root.oncompositionstart=()=>{typingComposing=true;};root.oncompositionend=()=>{typingComposing=false;commitNativeText();};
 root.onbeforeinput=e=>{
  nativeBefore=typingRange();
  if(e.isComposing)return;
  if(typingBusy){e.preventDefault();if(e.data){typingActions.push({text:e.data,marks:typingMarks});typingMarks={};}else if(e.inputType.startsWith('delete'))typingActions.push({deleteType:e.inputType});return;}
  if(e.inputType==='insertParagraph'||e.inputType==='insertLineBreak'){e.preventDefault();rangeAction(e.inputType==='insertParagraph'?'\n':'\v');return;}
  if(e.inputType==='historyUndo'||e.inputType==='historyRedo'){e.preventDefault();queue(e.inputType==='historyUndo'?'undo':'redo');return;}
  const target=e.getTargetRanges?.()[0],range=typingRange(target);
  if(!range){e.preventDefault();return;}
  const nativeNode=textElement(target?.startContainer||getSelection().anchorNode);
  const single=!nativeNode?.dataset.caretPlaceholder&&range.start.id===range.end.id&&nativeNode?.dataset.joeosId===range.start.id&&meta[range.start.id]?.editable&&!range.start.id.startsWith('e')&&(!target||textElement(target.endContainer)===nativeNode);
  const replacesWhole=range.start.offset===0&&range.end.offset===(meta[range.end.id]?.text.length||0)&&range.start.offset!==range.end.offset;
  if(single&&!replacesWhole&&!Object.keys(typingMarks).length&&(e.inputType.startsWith('delete')||['insertText','insertReplacementText'].includes(e.inputType)))return;
  e.preventDefault();
  if(e.inputType.startsWith('delete'))rangeAction('', 'replace',null,range);
  else if(e.inputType==='insertText'||e.inputType==='insertReplacementText')rangeAction(e.data||'','replace',null,range);
 };
 const clipboardText=()=>{const s=getSelection();if(!s.rangeCount)return '';const fragment=s.getRangeAt(0).cloneContents();for(const el of fragment.querySelectorAll('[data-caret-placeholder]'))el.textContent=el.textContent.replace(/^\u200b/,'');const holder=document.createElement('div');holder.style.cssText='position:fixed;left:-100000px;white-space:pre-wrap';holder.setAttribute('aria-hidden','true');holder.append(fragment);document.body.append(holder);const text=holder.innerText;holder.remove();return text;};
 root.oncopy=e=>{if(!getSelection().isCollapsed&&writing().querySelector('[data-caret-placeholder]')){e.preventDefault();e.clipboardData.setData('text/plain',clipboardText());}};
 root.oncut=e=>{if(!getSelection().isCollapsed){e.preventDefault();e.clipboardData.setData('text/plain',clipboardText());rangeAction('');}};
 root.onpaste=e=>{e.preventDefault();rangeAction(e.clipboardData.getData('text/plain').replace(/\r\n?/g,'\n'));};
 root.ondrop=e=>{e.preventDefault();notify('Paste text with Ctrl+V. Moving embedded objects by dragging is not supported yet.');};
 root.onkeydown=e=>{
  if(e.isComposing)return;
  const selection=getSelection(),anchor=textElement(selection.anchorNode);
  if(selection.isCollapsed&&anchor?.dataset.caretPlaceholder&&anchor.textContent==='\u200b'&&['Backspace','ArrowLeft'].includes(e.key))selection.collapse(anchor.firstChild,0);
  if(e.key==='Tab'){e.preventDefault();if(!e.ctrlKey&&tableTab(e.shiftKey))return;if(e.shiftKey)rangeAction('outdent','paragraph');else rangeAction('\t');return;}
  if(e.key==='Enter'){e.preventDefault();if(e.ctrlKey&&e.shiftKey){notify('Column breaks are not supported yet. No page break was inserted.');return;}rangeAction(e.ctrlKey?'\f':e.shiftKey?'\v':'\n');return;}
  if(e.key==='Escape'){document.getElementById('findPanel')?.remove();return;}
  if(e.key==='F3'){e.preventDefault();findNext(e.shiftKey);return;}
  if(e.key==='F12'){e.preventDefault();queue('saveAs');return;}
  if(e.ctrlKey&&!e.altKey){
   const key=e.key.toLowerCase();
   const paragraphCommands={l:'left',e:'center',r:'right',j:'both',m:e.shiftKey?'outdent':'indent',t:e.shiftKey?'unhang':'hang','1':'single','2':'double','5':'onehalf'};
   if(paragraphCommands[key]&&(!e.shiftKey||['m','t'].includes(key))){e.preventDefault();rangeAction(paragraphCommands[key],'paragraph');return;}
   if(key==='f'||key==='h'){e.preventDefault();openFind(key==='h');return;}
   if(key===' '&&e.shiftKey){e.preventDefault();rangeAction('\u00a0');return;}
   if(key==='-'){e.preventDefault();rangeAction(e.shiftKey?'\u2011':'\u00ad');return;}
   if(key==='['||key===']'||(e.shiftKey&&(key==='<'||key==='>'))){
    e.preventDefault();const point=typingRange()?.start,el=point&&writing().querySelector(`[data-joeos-id="${point.id}"]`);
    let size=typingMarks.size||parseFloat(el&&getComputedStyle(el).fontSize)*.75||11;
    const increase=key===']'||key==='>';
    if(e.shiftKey){const sizes=[8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72];size=increase?(sizes.find(n=>n>size)||size+2):([...sizes].reverse().find(n=>n<size)||size-1);}
    else size+=increase?1:-1;
    setTypingFormat('size',Math.max(1,Math.min(400,size)));return;
   }
   if(key==='home'||key==='end'){
    e.preventDefault();const els=[...writing().querySelectorAll('[data-joeos-id]')],el=key==='home'?els[0]:els.at(-1);
    if(el){const point=domPoint({id:el.dataset.joeosId,offset:key==='home'?0:el.textContent.length}),sel=getSelection();if(e.shiftKey)sel.extend(...point);else sel.collapse(...point);el.scrollIntoView({block:'nearest'});}return;
   }
   if(typingBusy&&['z','y'].includes(key)){e.preventDefault();typingActions.push({command:key==='y'||e.shiftKey?'redo':'undo'});return;}
   if(['b','i','u'].includes(key)){
    e.preventDefault();const property={b:'bold',i:'italic',u:'underline'}[key];
    const point=typingRange()?.start,el=point&&writing().querySelector(`[data-joeos-id="${point.id}"]`),style=el&&getComputedStyle(el);
    const current=property in typingMarks?typingMarks[property]:property==='bold'?Number(style?.fontWeight)>=600:property==='italic'?style?.fontStyle==='italic':style?.textDecorationLine.includes('underline');
    setTypingFormat(property,!current);
   }
   if(key==='a'){e.preventDefault();const r=document.createRange();r.selectNodeContents(root);getSelection().removeAllRanges();getSelection().addRange(r);}
  }
 };
 const first=root.querySelector('[data-joeos-id]:not([contenteditable="false"])');
 restoreTyping(payload.caret||bookmark||(first?{id:first.dataset.joeosId,offset:0}:null));
}
document.addEventListener('selectionchange',()=>{
 const selection=getSelection(),el=textElement(selection.anchorNode);
 if(el&&writing().contains(el))choose(el);
});


let findText='';
function openFind(replace=false){
 document.getElementById('findPanel')?.remove();
 const panel=document.createElement('section');panel.id='findPanel';panel.setAttribute('aria-label','Find and replace');
 panel.innerHTML='<label>Find <input id="findText" aria-label="Find text"></label><button id="findNext">Next</button><button id="findPrevious">Previous</button>'+(replace?'<label>Replace with <input id="replacementText" aria-label="Replacement text"></label><button id="replaceMatch">Replace</button>':'')+'<button id="closeFind">Close</button><p id="findStatus" role="status"></p>';
 document.querySelector('aside').prepend(panel);
 const input=panel.querySelector('#findText');input.value=findText;
 input.oninput=()=>{findText=input.value;};
 input.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();findNext(e.shiftKey);}if(e.key==='Escape'){panel.remove();writing().focus();}};
 for(const [id,back] of [['findNext',false],['findPrevious',true]]){panel.querySelector('#'+id).onmousedown=e=>e.preventDefault();panel.querySelector('#'+id).onclick=()=>findNext(back);}
 panel.querySelector('#closeFind').onclick=()=>{panel.remove();writing().focus();};
 if(replace){const button=panel.querySelector('#replaceMatch');button.onmousedown=e=>e.preventDefault();button.onclick=()=>{const r=typingRange();if(r&&getSelection().toString().toLocaleLowerCase()===findText.toLocaleLowerCase())rangeAction(panel.querySelector('#replacementText').value);else findNext();};}
 input.focus();input.select();
}
function findNext(back=false){
 if(!findText){openFind();return;}
 const walker=document.createTreeWalker(writing(),NodeFilter.SHOW_TEXT),nodes=[];let text='',n;
 while(n=walker.nextNode()){nodes.push({node:n,start:text.length});text+=n.textContent;}
 const sel=getSelection();let from=back?text.length:0;
 if(sel.rangeCount&&writing().contains(sel.anchorNode)){
  const r=sel.getRangeAt(0),node=back?r.startContainer:r.endContainer,offset=back?r.startOffset:r.endOffset,entry=nodes.find(x=>x.node===node);if(entry)from=entry.start+offset;
 }
 const hay=text.toLocaleLowerCase(),needle=findText.toLocaleLowerCase();
 let at=back?hay.lastIndexOf(needle,from-needle.length):hay.indexOf(needle,from);
 if(at<0)at=back?hay.lastIndexOf(needle):hay.indexOf(needle);
 const status=document.getElementById('findStatus');if(at<0){if(status)status.textContent='No matches';return;}
 const point=offset=>{const entry=nodes.find(x=>offset>=x.start&&offset<x.start+x.node.length)||nodes.at(-1);return [entry.node,offset-entry.start];};
 const range=document.createRange();range.setStart(...point(at));range.setEnd(...point(at+findText.length));writing().focus();sel.removeAllRanges();sel.addRange(range);range.startContainer.parentElement.scrollIntoView({block:'nearest'});if(status)status.textContent='Match selected';
}

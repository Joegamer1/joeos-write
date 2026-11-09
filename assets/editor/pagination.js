'use strict';
// Only the temporary view is paginated. Never emit page breaks into OOXML.
function paginatePreview(root,layout){
 const report={addedPages:0,splitParagraphs:0,splitTables:0,unsupported:0};
 const sourcePages=[...root.querySelectorAll('section.docx')];
 const zoom=Number(getComputedStyle(root).zoom)||1;
 const flow=p=>layout?.paragraphs[p.dataset.joeosParagraph]?.flow||{};
 function splitParagraph(p,bottom){
  if(p.querySelector('img,svg,table,br')||flow(p).keepLines)return null;
  const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT),chars=[];let n;
  while(n=walker.nextNode())for(let i=0;i<n.length;i++)chars.push([n,i]);
  if(chars.length<2||chars.length>100000)return null;
  const lines=[];
  for(let i=0;i<chars.length;i++){
   const [n,at]=chars[i],r=document.createRange();r.setStart(n,at);r.setEnd(n,at+1);
   const box=r.getBoundingClientRect();let line=lines.at(-1);
   if(!line||Math.abs(box.top-line.top)>2*zoom){line={top:box.top,bottom:box.bottom,start:i};lines.push(line);}
   else line.bottom=Math.max(line.bottom,box.bottom);
  }
  let count=lines.filter(l=>l.bottom<=bottom+0.5).length;
  if(flow(p).widowControl!==false){if(lines.length-count===1)count--;if(count<2)return null;}
  if(count<1||count>=lines.length)return null;
  const [node,offset]=chars[lines[count].start],range=document.createRange();
  range.setStart(node,offset);range.setEnd(p,p.childNodes.length);
  const backup=p.cloneNode(true),tail=p.cloneNode(false);tail.appendChild(range.extractContents());
  p.style.marginBottom='0px';tail.style.marginTop='0px';tail.style.textIndent='0px';
  const excess=p.getBoundingClientRect().bottom-bottom;
  if(excess>0.5){p.replaceChildren(...backup.childNodes);p.style.cssText=backup.style.cssText;return splitParagraph(p,bottom-excess-1);}
  // Split source segments are edited as one original segment in the sidebar.
  const left=new Set([...p.querySelectorAll('[data-joeos-id]')].map(e=>e.dataset.joeosId));
  for(const el of tail.querySelectorAll('[data-joeos-id]'))if(left.has(el.dataset.joeosId)){
   el.dataset.joeosFragment='true';for(const other of p.querySelectorAll('[data-joeos-id="'+el.dataset.joeosId+'"]'))other.dataset.joeosFragment='true';
  }
  report.splitParagraphs++;return tail;
 }
 for(const initial of sourcePages){
  const articles=[...initial.children].filter(e=>e.tagName==='ARTICLE');
  // Multi-column/continuous sections and footnotes require a separate layout pass.
  if(articles.length!==1||Number(getComputedStyle(articles[0]).columnCount)>1||[...initial.children].some(e=>e.tagName==='OL')){report.unsupported++;continue;}
  let page=initial,article=articles[0];
  if(!(parseFloat(getComputedStyle(page).minHeight)>0)){
   page.style.width='612pt';page.style.minHeight='792pt';page.style.padding='72pt';
  }
  let iterations=0;
  while(iterations++<200){
   const ps=getComputedStyle(page),rect=page.getBoundingClientRect();
   const footer=page.querySelector(':scope > footer');
   const footerReserve=footer?Math.max(0,footer.getBoundingClientRect().height+parseFloat(getComputedStyle(footer).marginBottom)):0;
   const bottom=rect.top+(parseFloat(ps.minHeight)-parseFloat(ps.paddingBottom)-footerReserve)*zoom;
   const children=[...article.children];
   let at=children.findIndex(e=>e.getBoundingClientRect().bottom+Math.max(0,parseFloat(getComputedStyle(e).marginBottom))*zoom>bottom+0.5);
   if(at<0)break;
   const over=children[at];let tail=null;
   if(over.tagName==='P')tail=splitParagraph(over,bottom);
   if(!tail&&over.tagName==='TABLE'&&!over.querySelector('[rowspan]:not([rowspan="1"])')){
    const rows=[...over.rows],cut=rows.findIndex(r=>r.getBoundingClientRect().bottom>bottom+1);
    const headers=[];for(const row of rows){if(!row.dataset.joeosTableHeader)break;headers.push(row);}
    if(cut>headers.length&&cut>0){
     tail=over.cloneNode(false);for(const col of over.querySelectorAll(':scope > colgroup'))tail.appendChild(col.cloneNode(true));
     const body=document.createElement('tbody');tail.appendChild(body);for(const row of headers){const copy=row.cloneNode(true);copy.dataset.joeosRepeated='true';body.appendChild(copy);}
     rows.slice(cut).forEach(r=>body.appendChild(r));report.splitTables++;
    }
   }
   if(!tail){
    while(at>0&&flow(children[at-1]).keepNext)at--;
    if(at===0){
     // A keep-with-next chain longer than a page must be relaxed to make progress.
     at=children.indexOf(over);
     if(at===0){report.unsupported++;break;}
    }
   }
   const next=page.cloneNode(false);next.dataset.joeosContinuation='true';
   const nextIndex=[...root.querySelectorAll('section.docx')].indexOf(page)+1;
   const decorations=initial.joeosContinuationTemplates?.[nextIndex%2]||page;
   for(const el of decorations.children)if(el.tagName==='HEADER')next.appendChild(el.cloneNode(true));
   const nextArticle=article.cloneNode(false);next.appendChild(nextArticle);
   if(tail){nextArticle.appendChild(tail);at=children.indexOf(over)+1;}
   children.slice(at).forEach(e=>nextArticle.appendChild(e));
   for(const el of decorations.children)if(el.tagName==='FOOTER')next.appendChild(el.cloneNode(true));
   page.after(next);
   for(const decor of next.children)if(decor.tagName==='HEADER'||decor.tagName==='FOOTER'){applyDocumentTypography(decor,layout);applyDocumentLayout(decor,layout);correctWordSpacing(decor);}
   page=next;article=nextArticle;report.addedPages++;
  }
  if(iterations>=200)report.unsupported++;
 }
 return report;
}

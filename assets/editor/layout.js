'use strict';
// Preview corrections only. No value here is written into the original DOCX.
const naturalFontRatios=new Map();
function applyDocumentLayout(root,layout,pagesOnly=false){
 if(!layout)return;
 [...root.querySelectorAll('section.docx')].forEach((page,index)=>{
  const section=layout.sections[page.dataset.joeosSection];if(!section)return;
  if(section.widthPt!==null)page.style.width=section.widthPt+'pt';
  if(section.heightPt!==null)page.style.minHeight=section.heightPt+'pt';
  const m={...section.marginsPt},settings=layout.settings||{};
  if(settings.mirrorMargins&&index%2===1)[m.left,m.right]=[m.right,m.left];
  const gutterEdge=settings.gutterAtTop?'top':settings.mirrorMargins&&index%2===1?'right':'left';
  if(m.gutter)m[gutterEdge]=(m[gutterEdge]||0)+m.gutter;
  for(const edge of ['top','bottom','left','right'])if(m[edge]!==undefined)page.style.setProperty('padding-'+edge,Math.abs(m[edge])+'pt');
 });
 if(pagesOnly)return;
 for(const p of root.querySelectorAll('[data-joeos-paragraph]')){
  const metrics=layout.paragraphs[p.dataset.joeosParagraph];if(!metrics)continue;
  p.dataset.joeosStyle=metrics.style||'';
  if(metrics.beforePt!==undefined)p.style.marginTop=metrics.beforePt+'pt';
  if(metrics.afterPt!==undefined)p.style.marginBottom=metrics.afterPt+'pt';
  const ind=metrics.indentsPt||{};
  if(ind.start!==undefined||ind.left!==undefined)p.style.marginInlineStart=(ind.start??ind.left)+'pt';
  if(ind.end!==undefined||ind.right!==undefined)p.style.marginInlineEnd=(ind.end??ind.right)+'pt';
  if(ind.hanging!==undefined)p.style.textIndent=-ind.hanging+'pt';
  else if(ind.firstLine!==undefined)p.style.textIndent=ind.firstLine+'pt';
  if(metrics.contextualSpacing!==null)p.style.setProperty('--joeos-context-spacing',metrics.contextualSpacing?'1':'0');
  if(metrics.lineRule){
   p.style.setProperty('--joeos-line-auto',metrics.lineRule==='auto'?String(metrics.line):'0');
   p.style.setProperty('--joeos-line-min',metrics.lineRule==='atLeast'?metrics.line+'pt':'0pt');
   p.style.lineHeight=metrics.lineRule==='exact'?metrics.line+'pt':'normal';
  }
 }
}
function applyDocumentTypography(root,layout){
 // Resolve the document's own kerning threshold after style inheritance and
 // font-size resolution. No per-document substitutions or PDF conventions.
 for(const element of root.querySelectorAll('p,span')){
  const style=getComputedStyle(element),size=parseFloat(style.fontSize)*.75;
  const threshold=style.getPropertyValue('--joeos-kern-pt').trim();
  element.style.fontKerning=threshold!==''&&size>=Number(threshold)?'normal':'none';
  const ligatures=layout?.settings?.openTypeFeatures?style.getPropertyValue('--joeos-ligatures').trim().toLowerCase():'';
  const enabled=name=>ligatures==='all'||ligatures.includes(name);
  // CSS common-ligatures combines liga/clig, while Word controls them separately.
  // Contextual ligatures (clig) are distinct from contextual alternates (calt).
  element.style.fontVariantLigatures='none';
  const alternates=layout?.settings?.openTypeFeatures&&style.getPropertyValue('--joeos-context-alternates').trim()==='1';
  element.style.fontFeatureSettings=[['liga',enabled('standard')],['clig',enabled('contextual')],
   ['hlig',enabled('historical')],['dlig',enabled('discretional')],['calt',alternates]]
   .map(([tag,on])=>`"${tag}" ${on?1:0}`).join(', ');
 }
}
function naturalFontRatio(style){
 const key=style.fontFamily+'|'+style.fontWeight+'|'+style.fontStyle;
 if(!naturalFontRatios.has(key)){
  const context=document.createElement('canvas').getContext('2d');
  context.font=style.fontStyle+' '+style.fontWeight+' 1000px '+style.fontFamily;
  const metrics=context.measureText('Ag');
  naturalFontRatios.set(key,(metrics.fontBoundingBoxAscent+metrics.fontBoundingBoxDescent)/1000);
 }
 return naturalFontRatios.get(key);
}
function correctWordSpacing(root){
 const paragraphs=[...root.querySelectorAll('p')];
 const styleKey=p=>p ? p.dataset.joeosStyle??[...p.classList].filter(c=>!c.startsWith('docx-num-')).sort().join(' ') : null;
 for(const p of paragraphs){
  // The browser's paragraph strut must not impose the Normal-style size on
  // a paragraph whose actual runs are all smaller (e.g. compact résumé lines).
  if(p.dataset.joeosBaseSize===undefined)p.dataset.joeosBaseSize=p.style.fontSize;
  p.style.fontSize=p.dataset.joeosBaseSize;
  const textSizes=[...p.querySelectorAll('[data-joeos-id]')].filter(e=>e.textContent.length).map(e=>parseFloat(getComputedStyle(e).fontSize));
  if(textSizes.length)p.style.fontSize=Math.max(...textSizes)+'px';
  const s=getComputedStyle(p);
  const automatic=Number(s.getPropertyValue('--joeos-line-auto'));
  if(automatic>0){
   for(const element of [p,...p.querySelectorAll('span:not([data-joeos-id])')]){
    element.style.lineHeight=String(automatic*naturalFontRatio(getComputedStyle(element)));
   }
  }
  const minimum=s.getPropertyValue('--joeos-line-min').trim();
  if(parseFloat(minimum)>0){
   // Measure normal browser leading for the paragraph's actual fonts. "At least"
   // means a floor, never "font height plus the requested spacing".
   const probe=document.createElement('div');
   probe.style.cssText='position:absolute;left:-100000px;top:0;width:max-content;visibility:hidden;line-height:normal;';
   let natural=0;
   for(const element of [p,...p.querySelectorAll('span')]){
    const font=getComputedStyle(element);probe.style.fontFamily=font.fontFamily;probe.style.fontSize=font.fontSize;
    probe.style.fontWeight=font.fontWeight;probe.style.fontStyle=font.fontStyle;probe.textContent='Ag';
    document.body.appendChild(probe);natural=Math.max(natural,probe.getBoundingClientRect().height);
   }
   probe.style.height=minimum;const floor=probe.getBoundingClientRect().height;probe.remove();
   p.style.lineHeight=Math.max(natural,floor)+'px';
  }
  if(s.getPropertyValue('--joeos-context-spacing').trim()==='1'){
   if(p.previousElementSibling?.tagName==='P'&&styleKey(p.previousElementSibling)===styleKey(p))p.style.marginTop='0px';
   if(p.nextElementSibling?.tagName==='P'&&styleKey(p.nextElementSibling)===styleKey(p))p.style.marginBottom='0px';
  }
  for(const edge of ['Top','Bottom','Left','Right']){
   const space=s.getPropertyValue('--joeos-border-space-'+edge.toLowerCase()).trim();
   if(!space||s['border'+edge+'Style']==='none')continue;
   p.style['padding'+edge]=space;
   if(edge==='Left'||edge==='Right'){
    // CSS puts borders inside a paragraph; Word extends them beyond the text.
    const key='joeosOriginal'+edge;
    const margin=p.dataset[key]??(p.dataset[key]=s['margin'+edge]),border=s['border'+edge+'Width'];
    p.style['margin'+edge]=`calc(${margin} - ${space} - ${border})`;
   }
  }
 }
 for(const page of root.querySelectorAll('section.docx')){
  const s=getComputedStyle(page);
  page.style.setProperty('--joeos-page-height',parseFloat(s.minHeight)>0?s.minHeight:s.height);
  for(const edge of ['top','bottom','left','right'])page.style.setProperty('--joeos-margin-'+edge,s.getPropertyValue('padding-'+edge));
 }
}
function showPageMeasurements(page){
 if(!page)return;
 const s=getComputedStyle(page),inch=value=>(parseFloat(value)/96).toFixed(3).replace(/0+$/,'').replace(/\.$/,'');
 document.getElementById('pageMeasurements').textContent=
  `${inch(s.width)} × ${inch(s.minHeight)} in\nMargins: left ${inch(s.paddingLeft)}, right ${inch(s.paddingRight)}\nTop ${inch(s.paddingTop)}, bottom ${inch(s.paddingBottom)} in`;
}

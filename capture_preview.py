"""Local visual QA capture. Outputs can contain document text; use a private folder."""
import json,sys,time,os
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer
from joeos_app import Editor,digest

source,out=map(Path,sys.argv[1:3]);out.mkdir(parents=True,exist_ok=True)
app=QApplication([]);w=Editor(str(source));w.show();start=time.monotonic();started=False
setup="""(()=>{
 const sheet=document.createElement('style');sheet.textContent='body>header,body>nav,body>footer,#banner,aside{display:none!important}body,main{overflow:visible!important}.workspace{display:block!important}#render .docx-wrapper{padding:0!important}#render section.docx{margin:0!important;box-shadow:none!important}';sheet.textContent+='#render [data-joeos-id]{outline:none!important;background:transparent!important;caret-color:transparent!important}';sheet.textContent+=EXTRA_CSS;document.head.appendChild(sheet);
 document.getElementById('zoom').value='100%';zoom();
 correctWordSpacing(document.getElementById('render'));
 return true;
})()""".replace('EXTRA_CSS',json.dumps(os.environ.get('JOEOS_CAPTURE_CSS','')))
measure="""JSON.stringify((()=>{
 const page=document.querySelector('section.docx'),origin=page.getBoundingClientRect(),ctx=document.createElement('canvas').getContext('2d'),chars=[],paragraphs=[];
 for(const p of page.querySelectorAll('p')){
  const r=p.getBoundingClientRect(),s=getComputedStyle(p);paragraphs.push({text:p.textContent,x:(r.left-origin.left)*.75,y:(r.top-origin.top)*.75,h:r.height*.75,font:s.fontFamily,size:parseFloat(s.fontSize)*.75,line:s.lineHeight,marginTop:s.marginTop,marginBottom:s.marginBottom});
 }
 const walker=document.createTreeWalker(page,NodeFilter.SHOW_TEXT);let node;
 while(node=walker.nextNode()){
  const s=getComputedStyle(node.parentElement);ctx.font=s.fontStyle+' '+s.fontWeight+' '+s.fontSize+' '+s.fontFamily;const m=ctx.measureText('Ag');
  for(let i=0;i<node.length;i++){
   const range=document.createRange();range.setStart(node,i);range.setEnd(node,i+1);const r=range.getBoundingClientRect();
   chars.push({text:node.data[i],paragraph:[...page.querySelectorAll('p')].indexOf(node.parentElement.closest('p')),x:(r.left-origin.left)*.75,y:(r.top-origin.top)*.75,width:r.width*.75,height:r.height*.75,baseline:(r.top-origin.top+m.fontBoundingBoxAscent)*.75,font:s.fontFamily,weight:s.fontWeight,color:s.color,size:parseFloat(s.fontSize)*.75});
  }
 }
 return {page:{width:origin.width*.75,height:origin.height*.75},paragraphs,chars};
})())"""
def got(value):
 (out/'preview-metrics.json').write_text(value,encoding='utf-8');w.view.grab().save(str(out/'preview.png'));app.exit(0)
def measure_later():w.page.runJavaScript(measure,got)
def ready(_):
 w.setMinimumSize(816,1056);w.view.setFixedSize(816,1056);w.adjustSize();QTimer.singleShot(800,measure_later)
def tick():
 global started
 if w.render_error:print(w.render_error);app.exit(1)
 if time.monotonic()-start>30:print('Capture timed out');app.exit(1)
 if w.render_result and not started:started=True;w.page.runJavaScript(setup,ready)
t=QTimer();t.timeout.connect(tick);t.start(200)
sys.exit(app.exec())

"""Assert preview geometry against explicit OOXML units, independently of Word."""
import io,json,sys,time
from pathlib import Path
from docx import Document
from docx.shared import Inches,Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer
from PySide6.QtGui import QFont,QFontMetricsF
from joeos_app import Editor
from preserved_docx import PreservedDocx

d=Document();sec=d.sections[0];sec.page_height=Inches(5);sec.page_width=Inches(6);sec.top_margin=sec.bottom_margin=sec.left_margin=sec.right_margin=Inches(.5)
sec.different_first_page_header_footer=True
sec.first_page_header.paragraphs[0].text='FIRST HEADER'
sec.header.paragraphs[0].text='DEFAULT HEADER'
sec.even_page_header.paragraphs[0].text='EVEN HEADER'
d.settings.element.append(OxmlElement('w:evenAndOddHeaders'))
d.styles['Normal'].font.name='Cambria';d.styles['Normal'].font.size=Pt(12)
p=d.add_paragraph(' '.join('word'+str(i) for i in range(320)))
for i in range(15):d.add_paragraph('Paragraph '+str(i)+' with independent spacing.')
t=d.add_table(rows=18,cols=2)
t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
for i,row in enumerate(t.rows):
 for j,cell in enumerate(row.cells):cell.text='Cell '+str(i)+'-'+str(j)
b=io.BytesIO();d.save(b);raw=b.getvalue()
app=QApplication([]);w=Editor();w.model=PreservedDocx(raw);w.show();start=time.monotonic();done=False

script="""JSON.stringify((()=>{
 const pages=[...document.querySelectorAll('section.docx')];
 return {sizes:pages.map(p=>({height:getComputedStyle(p).height,min:getComputedStyle(p).minHeight,article:p.querySelector('article').getBoundingClientRect().height,header:p.querySelector('header')?.getBoundingClientRect().height,lastMargin:getComputedStyle(p.querySelector('article').lastElementChild).marginBottom})),headers:pages.map(p=>p.querySelector('header')?.textContent),pages:pages.length,report:paginationReport,fonts:[...document.getElementById('font').options].map(o=>o.value),segments:[...document.querySelectorAll('[data-joeos-id]')].filter(e=>!e.closest('[data-joeos-repeated]')).map(e=>({id:e.dataset.joeosId,text:e.textContent,fragment:!!e.dataset.joeosFragment,editable:e.contentEditable})),overfull:pages.filter(p=>parseFloat(getComputedStyle(p).height)>parseFloat(getComputedStyle(p).minHeight)+1).length};
})())"""
def check(value):
 try:
  r=json.loads(value);assert r['pages']>3,r;assert r['overfull']==0,r['sizes'];assert r['report']['splitParagraphs']>0 and r['report']['splitTables']>0,r
  for item in [m for m in w.model.metadata() if m['part']=='word/document.xml']:assert ''.join(x['text'] for x in r['segments'] if x['id']==item['id'])==item['text'],item['id']
  assert all(x['editable']=='inherit' for x in r['segments'] if x['fragment'])
  assert r['headers'][0]=='FIRST HEADER' and r['headers'][1]=='EVEN HEADER' and r['headers'][2]=='DEFAULT HEADER',r['headers']
  assert len(r['fonts'])>30 and all(f in r['fonts'] for f in ['Arial','Times New Roman','Cambria']),r['fonts']
  assert w.model.bytes()==raw
  Path('audit-output/pagination-test.json').write_text(json.dumps({'pages':r['pages'],'report':r['report'],'overfull':r['overfull'],'all_text_preserved':True,'font_count':len(r['fonts']),'unchanged_save_identical':True},indent=2))
  w.page.runJavaScript("(()=>{const e=[...document.querySelectorAll('article [data-joeos-id]')].find(e=>!e.dataset.joeosFragment);restoreTyping({start:{id:e.dataset.joeosId,offset:0},end:{id:e.dataset.joeosId,offset:e.textContent.length}});document.getElementById('font').value='Arial';document.getElementById('font').dispatchEvent(new Event('input'));document.getElementById('applyformat').click();})()");QTimer.singleShot(700,check_font)
 except Exception as e:__import__('traceback').print_exc();print(repr(e));app.exit(1)
def check_font():
 try:
  assert b'w:ascii="Arial"' in w.model.parts['word/document.xml']
  from test_preservation import parts
  a,b=parts(raw),parts(w.model.bytes());assert [n for n in a if a[n]!=b[n]]==['word/document.xml']
  w.dirty=False;print('Pagination, text coverage, full font list and actual dropdown font edit passed.');app.exit(0)
 except Exception as e:__import__('traceback').print_exc();print(repr(e));w.dirty=False;app.exit(1)
def tick():
 global done
 if time.monotonic()-start>30:print('Timeout');app.exit(1)
 if w.render_error:print(w.render_error);app.exit(1)
 if w.render_result and not done:done=True;w.page.runJavaScript(script,check)
timer=QTimer();timer.timeout.connect(tick);timer.start(250)
sys.exit(app.exec())

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

d=Document();s=d.sections[0];s.page_width=Inches(11);s.page_height=Inches(8.5)
s.left_margin=Inches(.75);s.right_margin=Inches(1.25);s.top_margin=Inches(.6);s.bottom_margin=Inches(.8);s.gutter=Inches(.2)
d.styles['Normal'].font.name='Arial';d.styles['Normal'].font.size=Pt(12)
mirror=OxmlElement('w:mirrorMargins');d.settings.element.append(mirror)
p=d.add_paragraph('Precise automatic spacing');f=p.paragraph_format;f.left_indent=Inches(.25);f.right_indent=Inches(.125);f.first_line_indent=Inches(-.125);f.space_before=Pt(0);f.space_after=Pt(6)
spacing=p._p.get_or_add_pPr().get_or_add_spacing();spacing.set(qn('w:line'),'245');spacing.attrib.pop(qn('w:lineRule'),None)
p=d.add_paragraph('Minimum line spacing');f=p.paragraph_format;f.space_before=Pt(12);f.space_after=Pt(0)
spacing=p._p.get_or_add_pPr().get_or_add_spacing();spacing.set(qn('w:line'),'400');spacing.set(qn('w:lineRule'),'atLeast')
p=d.add_paragraph('Border with three-point text clearance');p.paragraph_format.left_indent=Inches(.5)
borders=OxmlElement('w:pBdr')
for side in ['left','bottom']:
 b=OxmlElement('w:'+side)
 for k,v in {'val':'single','sz':'8','space':'3','color':'174E50'}.items():b.set(qn('w:'+k),v)
 borders.append(b)
p._p.get_or_add_pPr().append(borders)
for text in ['Contextual spacing one','Contextual spacing two']:
 p=d.add_paragraph(text);p.paragraph_format.space_before=Pt(10);p.paragraph_format.space_after=Pt(10);p._p.get_or_add_pPr().append(OxmlElement('w:contextualSpacing'))
d.add_page_break();second=d.add_paragraph('Mirrored second page')
char=d.styles.add_style('JoeTestCharacter',WD_STYLE_TYPE.CHARACTER);char.font.name='Courier New'
second.add_run(' Character override',style=char)
second.add_run(' Direct override',style=char).font.name='Times New Roman'
b=io.BytesIO();d.save(b);raw=b.getvalue()
app=QApplication([]);w=Editor();w.model=PreservedDocx(raw);w.show();start=time.monotonic();done=False

script="""JSON.stringify((()=>{
 const pages=[...document.querySelectorAll('section.docx')],p=[...pages[0].querySelectorAll('article p')],css=e=>getComputedStyle(e),n=v=>parseFloat(v),v=e=>{const s=css(e);return {width:n(s.width),height:n(s.minHeight),left:n(s.paddingLeft),right:n(s.paddingRight),top:n(s.paddingTop),bottom:n(s.paddingBottom)}};
 const a=css(p[0]),b=css(p[1]),c=css(p[2]),z=n(css(document.querySelector('#render')).zoom)||1;
 return {characterFonts:[...pages[1].querySelectorAll('[data-joeos-id]')].map(e=>css(e).fontFamily),fonts:[a.fontFamily,...[...p[0].querySelectorAll('span')].map(e=>css(e).fontFamily)],pages:pages.map(v),indent:{left:n(a.marginLeft),right:n(a.marginRight),first:n(a.textIndent)},lineRatio:n(a.lineHeight)/n(a.fontSize),minimum:n(b.lineHeight),gap:(p[1].getBoundingClientRect().top-p[0].getBoundingClientRect().bottom)/z,border:{left:n(c.paddingLeft),bottom:n(c.paddingBottom),textOrigin:n(c.marginLeft)+n(c.borderLeftWidth)+n(c.paddingLeft)},context:{after:css(p[3]).marginBottom,before:css(p[4]).marginTop}};
})())"""
def check(value):
 try:
  r=json.loads(value);Path('audit-output/layout-metrics.json').write_text(json.dumps(r,indent=2))
  def near(a,b):assert abs(a-b)<.15,(a,b)
  assert len(r['pages'])==2
  assert r['fonts']==['Arial','Arial','Arial'],r['fonts']
  assert r['characterFonts']==['Arial','\"Courier New\"','\"Times New Roman\"'],r['characterFonts']
  for page in r['pages']:near(page['width'],1056);near(page['height'],816);near(page['top'],57.6);near(page['bottom'],76.8)
  near(r['pages'][0]['left'],91.2);near(r['pages'][0]['right'],120)
  near(r['pages'][1]['left'],120);near(r['pages'][1]['right'],91.2)
  near(r['indent']['left'],24);near(r['indent']['right'],12);near(r['indent']['first'],-12)
  font=QFont('Arial');font.setPixelSize(1000)
  expected=QFontMetricsF(font).height()/1000*245/240
  assert abs(r['lineRatio']-expected)<.003,(r['lineRatio'],expected)
  near(r['minimum'],400/20*96/72);near(r['gap'],16)
  near(r['border']['left'],4);near(r['border']['bottom'],4);near(r['border']['textOrigin'],48)
  assert r['context']=={'after':'0px','before':'0px'}
  assert w.model.bytes()==raw
  w.page.runJavaScript("document.getElementById('guides').click()")
  QTimer.singleShot(500,finish)
 except Exception as e:print(repr(e),flush=True);app.exit(1)
def finish():w.grab().save('audit-output/layout-guides.png');print('Geometry and spacing assertions passed.',flush=True);app.exit(0)
def tick():
 global done
 if w.render_error:print(w.render_error);app.exit(1)
 if time.monotonic()-start>30:print('Timed out');app.exit(1)
 if w.render_result and not done:done=True;w.page.runJavaScript(script,check)
t=QTimer();t.timeout.connect(tick);t.start(200)
sys.exit(app.exec())

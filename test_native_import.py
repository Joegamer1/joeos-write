"""Regression: imports use document formatting even with legacy preferences present."""
import io,json,sys,time,tempfile,os,re
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication
from joeos_app import Editor,digest

def prop(parent,name,value):
    el=OxmlElement('w:'+name);el.set(qn('w:val'),str(value));parent.append(el)

folder=tempfile.TemporaryDirectory(dir='.tmp');root=Path(folder.name).resolve()
d=Document();d.styles['Normal'].font.name='Arial'
style=d.styles.add_style('Inherited typography',WD_STYLE_TYPE.PARAGRAPH)
style.base_style=d.styles['Normal'];style.font.size=Pt(14)
prop(style.element.get_or_add_rPr(),'kern',28)
prop(style.element.get_or_add_rPr(),'spacing',20)
lig=etree.SubElement(style.element.get_or_add_rPr(),'{http://schemas.microsoft.com/office/word/2010/wordml}ligatures')
lig.set('{http://schemas.microsoft.com/office/word/2010/wordml}val','standardContextual')
for size,spacing in [(14,None),(11,None),(14,-10)]:
    r=d.add_paragraph(style=style).add_run('AV typography');r.font.size=Pt(size)
    if spacing is not None:prop(r._r.get_or_add_rPr(),'spacing',spacing)
d.add_paragraph('No inherited kerning')
for mode,value in [('default','  default  '),('preserve','  preserve  '),('default','\u00a0nonbreaking\u00a0'),('inherit','  inherited  ')]:
    r=d.add_paragraph().add_run(value);t=r._r.find(qn('w:t'))
    t.attrib.pop(qn('xml:space'),None)
    if mode=='preserve':t.set(qn('xml:space'),'preserve')
    if mode=='inherit':r._r.set(qn('xml:space'),'preserve')
out=io.BytesIO();d.save(out);raw=out.getvalue();source=root/'source.docx';source.write_bytes(raw)
os.environ['LOCALAPPDATA']=str(root)
legacy=root/'JoeOS/Write/preview-profiles.json';legacy.parent.mkdir(parents=True)
legacy.write_text(json.dumps({digest(raw):{'fontSubstitutions':{'Arial':'Noto Serif'},'pdfSpacing':True}}))
app=QApplication([]);w=Editor(source);w.show();start=time.monotonic();waiting=False;stage=0

def checked(value):
    global waiting,stage
    try:
        result=json.loads(value);assert not result['toggle'];rows=result['rows']
        assert all('Arial' in r['font'] for r in rows),rows
        assert [r['kern'] for r in rows[:4]]==['normal','none','normal','none'],rows
        features={tag:int(value or 1) for tag,value in re.findall(r'"(\w+)"(?: (\d+))?',rows[0]['features'])}
        assert features=={'liga':1,'clig':1,'calt':0,'hlig':0,'dlig':0},rows
        assert rows[3]['ligatures']=='none' or 'no-common-ligatures' in rows[3]['ligatures'],rows
        assert [r['text'] for r in rows[4:]]==['default','  preserve  ','\u00a0nonbreaking\u00a0','  inherited  '],rows
        assert abs(float(rows[0]['spacing'].replace('px',''))-4/3)<.01,rows
        assert abs(float(rows[2]['spacing'].replace('px',''))+2/3)<.01,rows
        assert w.model.bytes()==raw and source.read_bytes()==raw
        if stage==0:
            stage=1;waiting=False;w.open_document(str(source));return
        w.model.save(root/'saved.docx');assert (root/'saved.docx').read_bytes()==raw
        Path('audit-output/document-import.json').write_text(json.dumps(result,indent=2))
        print('Native import, inherited kerning, signed character spacing, legacy preference isolation, reopen and unchanged save passed.');app.exit(0)
    except Exception as e:print(repr(e));app.exit(1)

def tick():
    global waiting
    if time.monotonic()-start>30 or w.render_error:print(w.render_error or 'Timeout');app.exit(1);return
    if waiting or not w.render_result:return
    waiting=True
    w.page.runJavaScript("JSON.stringify({toggle:!!document.getElementById('appearance'),rows:[...document.querySelectorAll('[data-joeos-id]')].map(e=>{let s=getComputedStyle(e);return {font:s.fontFamily,kern:s.fontKerning,spacing:s.letterSpacing,ligatures:s.fontVariantLigatures,features:s.fontFeatureSettings,text:e.textContent}})})",checked)
t=QTimer();t.timeout.connect(tick);t.start(300)
code=app.exec();folder.cleanup();sys.exit(code)

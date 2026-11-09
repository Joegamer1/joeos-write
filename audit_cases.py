"""Non-destructive compatibility probes. Findings are observations, not passing claims."""
import os, json
os.environ['QT_QPA_PLATFORM']='offscreen'
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QTextCursor
from docx_io import load_docx,save_docx
from app import Editor
app=QApplication([])
root=Path('audit-output');root.mkdir(exist_ok=True)
findings={}
def cycle(name,d):
    p=root/(name+'.docx');d.save(p);q,w=load_docx(p);save_docx(q,root/(name+'-saved.docx'))
    return q,Document(root/(name+'-saved.docx')),w
d=Document();s=d.sections[0];s.page_width=Inches(8.27);s.left_margin=Inches(.5)
s.header.paragraphs[0].text='HEADER SENTINEL';s.footer.paragraphs[0].text='FOOTER SENTINEL'
p=d.add_paragraph('Spaced');p.paragraph_format.space_after=Pt(24);p.paragraph_format.left_indent=Inches(.4)
q,r,w=cycle('geometry',d)
findings['geometry']={'width_before':s.page_width.inches,'width_after':r.sections[0].page_width.inches,'margin_before':.5,'margin_after':r.sections[0].left_margin.inches,'header_after':r.sections[0].header.paragraphs[0].text,'footer_after':r.sections[0].footer.paragraphs[0].text,'space_after':str(r.paragraphs[0].paragraph_format.space_after),'indent_after':str(r.paragraphs[0].paragraph_format.left_indent)}
d=Document();base=d.styles['Normal'];base.font.name='Arial';base.font.size=Pt(15)
style=d.styles.add_style('Inherited',1);style.base_style=base
d.add_paragraph('Inherited font',style='Inherited')
q,r,w=cycle('inheritance',d);f=q.begin().begin().fragment().charFormat()
findings['inheritance']={'font':f.fontFamilies(),'size':f.fontPointSize()}
d=Document()
for t in ['one','two','three']:d.add_paragraph(t,style='List Number')
q,r,w=cycle('numbering',d);b=q.begin();labels=[]
while b.isValid():
    if b.textList():labels.append(b.textList().itemText(b))
    b=b.next()
findings['numbering']=labels
d=Document();t=d.add_table(rows=2,cols=2);t.cell(0,0).merge(t.cell(0,1)).text='MERGED'
t.cell(1,0).text='Outer';nested=t.cell(1,1).add_table(rows=1,cols=1);nested.cell(0,0).text='NESTED'
q,r,w=cycle('tables',d);findings['tables']={'editor_text':q.toPlainText(),'saved_cells':[[c.text for c in row.cells] for row in r.tables[0].rows]}
d=Document();d.add_paragraph('Before');control=OxmlElement('w:sdt');body=OxmlElement('w:sdtContent');p=OxmlElement('w:p');run=OxmlElement('w:r');text=OxmlElement('w:t');text.text='CONTROL SENTINEL';run.append(text);p.append(run);body.append(p);control.append(body);d.element.body.insert(1,control)
d.add_paragraph('After');q,r,w=cycle('control',d);findings['control']={'text':q.toPlainText(),'warnings':w}
d=Document();p=d.add_paragraph('Before');p.add_run().add_break(__import__("docx.enum.text",fromlist=["WD_BREAK"]).WD_BREAK.PAGE);p.add_run('After');q,r,w=cycle('pagebreak',d)
findings['pagebreak']={'saved_page_break': 'w:type="page"' in r.element.xml,'warnings':w}
d=Document();p=d.add_paragraph();r=p.add_run('Power2');r.font.superscript=True;r.font.highlight_color=7
q,r,w=cycle('runfeatures',d);findings['runfeatures']={'superscript':r.paragraphs[0].runs[0].font.superscript,'highlight':str(r.paragraphs[0].runs[0].font.highlight_color)}
w=Editor();w.editor.setHtml('<p><span style="font-size:10.5pt">Half point</span></p>');c=w.editor.textCursor();c.select(QTextCursor.Document);w.editor.setTextCursor(c);w.sync_format(c.charFormat());before=c.charFormat().fontPointSize();display=w.sizebox.currentText();w.set_size();findings['size_control']={'before':before,'display':display,'after':w.editor.textCursor().charFormat().fontPointSize()};w.editor.document().setModified(False)
(root/'results.json').write_text(json.dumps(findings,indent=2),encoding='utf-8')
print(json.dumps(findings,indent=2))


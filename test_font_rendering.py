"""Installed-font regression and a public fixture for actual desktop Open tests."""
import base64,io,json,sys,time
from pathlib import Path
from fontTools.ttLib import TTFont
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer
from docx import Document
from docx.shared import Pt
from local_fonts import LocalFonts
from preserved_docx import PreservedDocx
from joeos_app import Editor

families=['Courier','AngsanaUPC','Bierstadt','Times New Roman (Headings CS)']
app=QApplication([])
loader=LocalFonts();faces=loader.faces(families+['Not an installed JoeOS test font'])
for family in families:
 selected=[f for f in faces if f['family']==family]
 assert {(f['weight'],f['style']) for f in selected}=={('400','normal'),('400','italic'),('700','normal'),('700','italic')},family
 for face in selected:
  font=TTFont(io.BytesIO(base64.b64decode(face['data'])))
  assert font['name'].getDebugName(1)==face['sourceFamily']
  if family=='AngsanaUPC':
   assert all(ord(c) in font.getBestCmap() for c in '\u0e01\u0e33\u0e25\u0e31\u0e07'), 'Missing Thai glyphs'
  font.close()
assert not any(f['family'].startswith('Not an installed') for f in faces)
# Exact outline family must take precedence over compatibility substitutions.
loader.index.insert(0,('Bierstadt',400,'normal',Path('C:/Windows/Fonts/arial.ttf'),None))
assert loader.faces(['Bierstadt'])[0]['sourceFamily']=='Bierstadt'
assert not loader.resolutions[0]['compatibility']

path=Path('audit-output/font-desktop-check.docx').resolve()
d=Document();d.styles['Normal'].font.name='Arial';d.styles['Normal'].font.size=Pt(10)
d.add_paragraph('JoeOS desktop font verification - original DOCX import')
for family in families:
 for bold,italic in [(False,False),(True,False),(False,True),(True,True)]:
  p=d.add_paragraph();p.paragraph_format.space_after=Pt(4)
  r=p.add_run(f'{family} | '+('bold ' if bold else '')+('italic ' if italic else '')+'iiii WWWW 0123456789 AV fi')
  if family=='AngsanaUPC':r.add_text(' \u0e20\u0e32\u0e29\u0e32\u0e44\u0e17\u0e22 \u0e01\u0e33\u0e25\u0e31\u0e07\u0e17\u0e14\u0e2a\u0e2d\u0e1a')
  r.font.name=family;r.font.size=Pt(12);r.bold=bold;r.italic=italic
  for script in ['eastAsia','cs']:r._r.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'+script,family)
d.save(path);raw=path.read_bytes();assert PreservedDocx(raw).bytes()==raw
w=Editor();w.show();started=time.monotonic();opened=False;waiting=False

def checked(raw_result):
 try:
  result=json.loads(raw_result)
  assert result['loaded']==16,result
  assert abs(result['courierNarrow']-result['courierWide'])<.001,result
  assert abs(result['angsana']-result['times'])>1,result
  assert w.model.bytes()==raw and path.read_bytes()==raw
  Path('audit-output/font-rendering-test.json').write_text(json.dumps(result,indent=2))
  print(json.dumps(result),flush=True);app.exit(0)
 except Exception as e:
  print(repr(e),flush=True);app.exit(1)

def tick():
 global opened,waiting
 if time.monotonic()-started>60:app.exit(1);return
 if waiting or not w.render_result:return
 if not opened:w.open_document(str(path));opened=True;return
 waiting=True
 w.page.runJavaScript("JSON.stringify((()=>{const c=document.createElement('canvas').getContext('2d');const width=(f,t)=>{c.font='32px '+JSON.stringify(f);return c.measureText(t).width;};return {loaded:[...document.fonts].filter(f=>['Courier','AngsanaUPC','Bierstadt','Times New Roman (Headings CS)'].includes(f.family)&&f.status==='loaded').length,courierNarrow:width('Courier','iiiiiiii'),courierWide:width('Courier','WWWWWWWW'),angsana:width('AngsanaUPC','Hamburgefontsiv AV fi 0123456789'),times:width('Times New Roman','Hamburgefontsiv AV fi 0123456789'),banner:document.getElementById('banner').textContent};})())",checked)
t=QTimer();t.timeout.connect(tick);t.start(400)
sys.exit(app.exec())

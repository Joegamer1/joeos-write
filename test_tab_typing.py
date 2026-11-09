"""Real Qt keyboard events exercise the continuous embedded editor."""
import io,json,time,sys
from pathlib import Path
from docx import Document
from PySide6.QtCore import Qt,QEventLoop,QTimer
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication
from joeos_app import Editor
from preserved_docx import PreservedDocx,W
from lxml import etree
app=QApplication([])
d=Document();p=d.add_paragraph();p.add_run('Hello ').bold=True;p.add_run('world').italic=True;d.add_paragraph('Next paragraph')
path=Path('audit-output/tab-check.docx').resolve();d.save(path);original=path.read_bytes()
w=Editor(str(path));w.show();w.activateWindow()
def spin(seconds=.2):
 until=time.monotonic()+seconds
 while time.monotonic()<until:app.processEvents();time.sleep(.005)
def js(code):
 result=[];w.page.runJavaScript(code,lambda x:result.append(x))
 until=time.monotonic()+8
 while not result and time.monotonic()<until:spin(.01)
 assert result,'JS timed out';return result[0]
def ready():
 until=time.monotonic()+15
 while time.monotonic()<until:
  spin(.05)
  if w.render_result and js("!document.body.classList.contains('busy')&&!typingBusy"):return
 raise AssertionError('Render timed out '+str(w.render_error))
def key(k,mods=Qt.NoModifier):QTest.keyClick(w.view.focusProxy(),k,mods);spin(.15)
def type_text(t):QTest.keyClicks(w.view.focusProxy(),t,delay=12);spin(.3)
def document_text():return ''.join(etree.fromstring(w.model.parts['word/document.xml']).itertext()) if not w.model.text_edits else ''.join(etree.fromstring(PreservedDocx(w.model.bytes()).parts['word/document.xml']).itertext())
try:
 ready()
 js("restoreTyping({id:'t2',offset:0});true")
 key(Qt.Key_Tab);ready()
 measurement=json.loads(js("JSON.stringify((()=>{let t=document.querySelector('.docx-tab-stop'),p=t.closest('p');return {width:t.getBoundingClientRect().width,x:t.getBoundingClientRect().right-p.getBoundingClientRect().left,caret:typingRange()}})())"))
 assert abs(measurement['width']-48)<.2,measurement
 assert abs(measurement['x']-48)<.2,measurement
 type_text('Indented');ready();spin(1)
 saved=Path('audit-output/tab-saved.docx');saved.write_bytes(w.model.bytes())
 assert Document(saved).paragraphs[1].text=='\tIndentedNext paragraph',Document(saved).paragraphs[1].text
 assert path.read_bytes()==original
 Path('audit-output/tab-test.json').write_text(json.dumps(measurement,indent=2))
 print('PASS actual editor Tab: 48px half-inch stop, typing after tab, DOCX reopen, original unchanged',flush=True)
finally:w.dirty=False;w.close();spin(.1)

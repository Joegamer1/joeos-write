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
path=Path('audit-output/space-caret-check.docx').resolve();d.save(path);original=path.read_bytes()
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
 js("restoreTyping({id:'t2',offset:14});const originalType=bridge.typeText;bridge.typeText=(...args)=>setTimeout(()=>originalType(...args),100);true")
 QTest.keyClicks(w.view.focusProxy(),' word ',delay=12)
 spin(2);ready()
 result=json.loads(js("JSON.stringify({text:document.querySelector('[data-joeos-id=\"t2\"]').textContent,caret:typingRange()})"))
 expected='Next paragraph word '
 assert result['text']==expected,result
 assert result['caret']['start']=={'id':'t2','offset':len(expected)},result
 QTest.keyClicks(w.view.focusProxy(),'next ',delay=12);spin(2);ready()
 result=json.loads(js("JSON.stringify({text:document.querySelector('[data-joeos-id=\"t2\"]').textContent,caret:typingRange()})"))
 assert result['text']==expected+'next ',result
 assert result['caret']['start']=={'id':'t2','offset':len(expected+'next ')},result
 Path('audit-output/space-caret-test.json').write_text(json.dumps(result,indent=2))
 print('PASS delayed edit acknowledgments: spaces and cursor stay after typed words',flush=True)
except Exception as e:
 print(repr(e),flush=True);sys.exit(1)
finally:w.dirty=False;w.close();spin(.1)

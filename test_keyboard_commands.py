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
path=Path('audit-output/keyboard-check.docx').resolve();d.save(path);original=path.read_bytes()
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
 ready();js("restoreTyping({id:'t2',offset:4});true")
 key(Qt.Key_E,Qt.ControlModifier);ready()
 def props():return etree.fromstring(w.model.parts['word/document.xml']).xpath('//w:body/w:p[2]/w:pPr',namespaces={'w':W})[0]
 assert props().find('{'+W+'}jc').get('{'+W+'}val')=='center'
 key(Qt.Key_M,Qt.ControlModifier);ready();assert props().find('{'+W+'}ind').get('{'+W+'}left')=='720'
 key(Qt.Key_M,Qt.ControlModifier|Qt.ShiftModifier);ready();assert props().find('{'+W+'}ind').get('{'+W+'}left')=='0'
 key(Qt.Key_2,Qt.ControlModifier);ready();assert props().find('{'+W+'}spacing').get('{'+W+'}line')=='480'
 key(Qt.Key_Z,Qt.ControlModifier);ready();assert props().find('{'+W+'}spacing') is None
 key(Qt.Key_End,Qt.ControlModifier);assert js('typingRange().end.offset')==14
 key(Qt.Key_Home,Qt.ControlModifier|Qt.ShiftModifier);assert js('getSelection().toString().includes("Hello")')
 key(Qt.Key_Home,Qt.ControlModifier);assert js('typingRange().start.offset')==0
 key(Qt.Key_Right,Qt.ControlModifier|Qt.ShiftModifier);assert js('getSelection().toString().trim()')=='Hello'
 js("restoreTyping({id:'t2',offset:0});true")
 key(Qt.Key_Right,Qt.ControlModifier);assert js('typingRange().start.offset')==5
 key(Qt.Key_Delete,Qt.ControlModifier);ready();assert 'paragraph' not in Document(io.BytesIO(w.model.bytes())).paragraphs[1].text
 key(Qt.Key_Z,Qt.ControlModifier);ready()
 key(Qt.Key_F,Qt.ControlModifier);type_text('paragraph');key(Qt.Key_Return);assert js('getSelection().toString()')=='paragraph'
 key(Qt.Key_Escape)
 key(Qt.Key_H,Qt.ControlModifier);js("document.getElementById('replacementText').value='section';findText='paragraph';findNext();true")
 js("document.getElementById('replaceMatch').click();true");ready();assert Document(io.BytesIO(w.model.bytes())).paragraphs[1].text=='Next section'
 key(Qt.Key_Escape)
 js("restoreTyping({id:'t0',offset:0});true")
 key(Qt.Key_Space,Qt.ControlModifier|Qt.ShiftModifier);ready()
 assert Document(io.BytesIO(w.model.bytes())).paragraphs[0].text.startswith('\u00a0')
 saved=Path('audit-output/keyboard-saved.docx');saved.write_bytes(w.model.bytes());assert Document(saved).paragraphs[1].text=='Next section'
 assert path.read_bytes()==original
 print('PASS keyboard alignment, indent, spacing, undo, movement/selection, Ctrl+Delete, find/replace, NBSP, saved DOCX',flush=True)
finally:w.dirty=False;w.close();spin(.1)

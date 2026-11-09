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
path=Path('audit-output/typing-check.docx').resolve();d.save(path);original=path.read_bytes()
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
 ready();assert w.model.bytes()==original;assert js('getSelection().isCollapsed')
 assert js("document.querySelectorAll('#render [data-joeos-id][contenteditable=\"plaintext-only\"]').length")==0
 js("restoreTyping({id:'t0',offset:3});true")
 type_text('X');assert 'HelXlo world' in document_text(),document_text()
 key(Qt.Key_Return);ready();type_text('new');assert 'new' in document_text(),document_text()
 key(Qt.Key_Tab);ready();type_text('tab');assert 'tab' in document_text(),document_text()
 key(Qt.Key_Return,Qt.ShiftModifier);ready();type_text('line')
 key(Qt.Key_B,Qt.ControlModifier);type_text('B');ready()
 root=etree.fromstring(PreservedDocx(w.model.bytes()).parts['word/document.xml']);ns={'w':W}
 assert root.xpath('count(//w:tab)',namespaces=ns)==1
 assert root.xpath('count(//w:br)',namespaces=ns)==1
 assert len(root.xpath('//w:body/w:p',namespaces=ns))==3
 assert root.xpath('//w:r[w:rPr/w:b[@w:val="0"]]/w:t[text()="B"]',namespaces=ns),'Ctrl+B should toggle inherited bold off for new text only'
 # Cursor moves across original run boundaries and selection can span them.
 js("const spans=[...document.querySelectorAll('#render [data-joeos-id]')];const e=spans.find(e=>e.textContent==='world');restoreTyping({id:e.dataset.joeosId,offset:0});true")
 key(Qt.Key_Left);key(Qt.Key_Right);type_text('Z');assert 'Z' in document_text()
 key(Qt.Key_Z,Qt.ControlModifier);ready();key(Qt.Key_Y,Qt.ControlModifier);ready()
 # Boundary deletion must join paragraphs rather than trap the cursor.
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({id:'t2',offset:0});true")
 key(Qt.Key_Backspace);ready()
 assert len(etree.fromstring(PreservedDocx(w.model.bytes()).parts['word/document.xml']).xpath('//w:body/w:p',namespaces=ns))==1,'Backspace did not join paragraphs'
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({start:{id:'t0',offset:3},end:{id:'t1',offset:2}});true")
 type_text('Q');ready();assert document_text()=='HelQrldNext paragraph',document_text()
 key(Qt.Key_Backspace,Qt.ControlModifier);ready()
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({id:'t0',offset:3});true")
 key(Qt.Key_Tab);ready();key(Qt.Key_Backspace);ready()
 assert not etree.fromstring(PreservedDocx(w.model.bytes()).parts['word/document.xml']).xpath('//w:tab',namespaces=ns),'Tab cannot be deleted'
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({start:{id:'t1',offset:0},end:{id:'t1',offset:5}});true")
 key(Qt.Key_B,Qt.ControlModifier);ready()
 assert js("getSelection().toString()")=='world','Formatting lost selected word'
 key(Qt.Key_I,Qt.ControlModifier);ready()
 root=etree.fromstring(PreservedDocx(w.model.bytes()).parts['word/document.xml'])
 assert root.xpath('//w:r[w:rPr/w:b[@w:val="1"] and w:rPr/w:i[@w:val="0"]]/w:t[text()="world"]',namespaces=ns),'Selected-word bold/italic failed'
 key(Qt.Key_Z,Qt.ControlModifier);ready();assert js("getSelection().toString()")=='world','Undo lost selection'
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({id:'t2',offset:14});true")
 key(Qt.Key_B,Qt.ControlModifier);type_text('Bold words');ready()
 key(Qt.Key_Return);ready();type_text('More');ready()
 key(Qt.Key_B,Qt.ControlModifier);type_text('plain');ready()
 root=etree.fromstring(PreservedDocx(w.model.bytes()).parts['word/document.xml'])
 bold=''.join(root.xpath('//w:r[w:rPr/w:b[not(@w:val) or @w:val="1"]]/w:t/text()',namespaces=ns))
 normal=''.join(root.xpath('//w:r[w:rPr/w:b[@w:val="0"]]/w:t/text()',namespaces=ns))
 assert 'Bold words' in bold and 'More' in bold and 'plain' in normal,(bold,normal)
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({start:{id:'t1',offset:0},end:{id:'t1',offset:5}});true")
 type_text('replacement');ready();assert document_text()=='Hello replacementNext paragraph',document_text()
 key(Qt.Key_End);key(Qt.Key_Left,Qt.ControlModifier|Qt.ShiftModifier)
 assert len(js("getSelection().toString()"))>1,'Word selection stopped at empty bookkeeping span'
 w.model.save('audit-output/typing-saved.docx')
 w.dirty=False;w._load(str(path));w.publish();ready();assert js('getSelection().isCollapsed'),'New document inherited old selection'
 w.model.save('audit-output/typing-saved.docx');saved=Path('audit-output/typing-saved.docx').read_bytes();assert PreservedDocx(saved)
 assert path.read_bytes()==original
 Path('audit-output/typing-test.json').write_text(json.dumps({'native_keys':True,'continuous_surface':True,'enter_tab_shift_enter':True,'ctrl_b':True,'undo_redo':True,'source_unchanged':True,'docx_reopened':True},indent=2))
 print('PASS native typing, Enter, Tab, Shift+Enter, Ctrl+B, arrows, undo/redo, DOCX reopen',flush=True)
except Exception as e:
 print(repr(e),flush=True)
 print(js("JSON.stringify({text:document.getElementById('render').innerText,sel:typingRange(),busy:typingBusy,banner:document.getElementById('banner').textContent})"),flush=True)
 sys.exit(1)
finally:w.dirty=False;w.close();spin(.1)

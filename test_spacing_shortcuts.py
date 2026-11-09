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
path=Path('audit-output/spacing-shortcuts-check.docx').resolve();d.save(path);original=path.read_bytes()
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
 js("window.livePage=document.querySelector('#render section.docx');window.liveParagraph=document.querySelector('#render p');window.badFrames=[];window.watchFrames=true;function checkFrame(){if(!watchFrames)return;if(!livePage.isConnected||getComputedStyle(document.getElementById('render')).opacity!=='1'||!document.querySelector('#render section.docx'))badFrames.push('blank/faded/replaced');requestAnimationFrame(checkFrame)};checkFrame();restoreTyping({id:'t2',offset:14});true")
 def root():return etree.fromstring(PreservedDocx(w.model.bytes()).parts['word/document.xml'])
 def attr(element,name):return root().xpath('//w:body/w:p[2]/w:pPr/w:'+element+'/@w:'+name,namespaces={'w':W})
 measurements=[]
 for i in range(2):
  key(Qt.Key_Tab);ready()
  m=json.loads(js("JSON.stringify((()=>{let r=getSelection().getRangeAt(0).getBoundingClientRect(),t=[...document.querySelectorAll('#render .docx-tab-stop')].at(-1);return {caret:r.x,tabEnd:t.getBoundingClientRect().right,height:r.height}})())"))
  assert m['height']>0 and abs(m['caret']-m['tabEnd'])<.2,m
  measurements.append(m)
 assert abs(measurements[1]['caret']-measurements[0]['caret']-48)<.2,measurements
 key(Qt.Key_Backspace);ready();assert len(root().xpath('//w:tab',namespaces={'w':W}))==1
 key(Qt.Key_Backspace);ready();assert not root().xpath('//w:tab',namespaces={'w':W})
 for k,mods,value in [(Qt.Key_M,Qt.ControlModifier,'720'),(Qt.Key_M,Qt.ControlModifier,'1440'),(Qt.Key_M,Qt.ControlModifier|Qt.ShiftModifier,'720'),(Qt.Key_Tab,Qt.ShiftModifier,'0')]:
  key(k,mods);ready();assert attr('ind','left')==[value],(k,attr('ind','left'))
 for k,mods,value in [(Qt.Key_T,Qt.ControlModifier,'720'),(Qt.Key_T,Qt.ControlModifier,'1440'),(Qt.Key_T,Qt.ControlModifier|Qt.ShiftModifier,'720')]:
  key(k,mods);ready();assert attr('ind','hanging')==[value]
 for k,value in [(Qt.Key_1,'240'),(Qt.Key_2,'480'),(Qt.Key_5,'360')]:
  key(k,Qt.ControlModifier);ready();assert attr('spacing','line')==[value]
 for k,value in [(Qt.Key_L,'left'),(Qt.Key_E,'center'),(Qt.Key_R,'right'),(Qt.Key_J,'both')]:
  key(k,Qt.ControlModifier);ready();assert attr('jc','val')==[value]
 key(Qt.Key_Z,Qt.ControlModifier);ready();assert attr('jc','val')==['right']
 key(Qt.Key_Y,Qt.ControlModifier);ready();assert attr('jc','val')==['both']
 key(Qt.Key_Return,Qt.ControlModifier|Qt.ShiftModifier);ready();assert not root().xpath('//w:br',namespaces={'w':W})
 key(Qt.Key_Return,Qt.ShiftModifier);ready();assert len(root().xpath('//w:br[not(@w:type)]',namespaces={'w':W}))==1
 key(Qt.Key_Return,Qt.ControlModifier);ready();assert len(root().xpath('//w:br[@w:type="page"]',namespaces={'w':W}))==1
 key(Qt.Key_Return);ready();assert len(root().xpath('//w:body/w:p',namespaces={'w':W}))==3
 type_text('After breaks');ready();spin(1)
 assert 'After breaks' in Document(io.BytesIO(w.model.bytes())).paragraphs[-1].text
 assert '\u200b' not in ''.join(root().itertext())
 report=json.loads(js("JSON.stringify({badFrames,firstParagraphRetained:liveParagraph===document.querySelector('#render p')})"));js('watchFrames=false;true')
 assert report['badFrames']==[],report
 assert report['firstParagraphRetained'],report
 # Selected paragraphs receive the same requested indentation.
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({start:{id:'t0',offset:0},end:{id:'t2',offset:14}});true")
 key(Qt.Key_M,Qt.ControlModifier);ready()
 assert root().xpath('//w:body/w:p/w:pPr/w:ind/@w:left',namespaces={'w':W})==['720','720']
 # Custom tab stop, and its geometry at another zoom.
 from docx.shared import Inches
 custom=Document();p=custom.add_paragraph('Custom');p.paragraph_format.tab_stops.add_tab_stop(Inches(1))
 table=custom.add_table(rows=1,cols=2);table.cell(0,0).text='Cell one';table.cell(0,1).text='Cell two'
 customPath=Path('audit-output/tab-table-check.docx');custom.save(customPath)
 w.dirty=False;w._load(str(customPath.resolve()));w.publish();ready()
 js("restoreTyping({id:'t0',offset:0});true");key(Qt.Key_Tab);ready()
 assert abs(js("document.querySelector('#render .docx-tab-stop').getBoundingClientRect().width")-96)<.2
 js("document.getElementById('zoom').value='75%';zoom();true");key(Qt.Key_Tab);ready()
 widths=json.loads(js("JSON.stringify([...document.querySelectorAll('#render .docx-tab-stop')].map(t=>t.getBoundingClientRect().width))"))
 assert abs(widths[0]-72)<.5 and abs(widths[1]-36)<.5,widths
 js("let el=document.querySelector('#render td [data-joeos-id]');restoreTyping({id:el.dataset.joeosId,offset:0});true")
 key(Qt.Key_Tab);ready();assert js("getSelection().toString()")=='Cell two'
 key(Qt.Key_Tab,Qt.ShiftModifier);ready();assert js("getSelection().toString()")=='Cell one'
 key(Qt.Key_Left);key(Qt.Key_Tab,Qt.ControlModifier);ready()
 assert len(root().xpath('//w:tbl//w:tab',namespaces={'w':W}))==1
 # Two immediate presses must produce two tabs, not consume the first key.
 w.dirty=False;w._load(str(path));w.publish();ready()
 js("restoreTyping({id:'t2',offset:0});true")
 QTest.keyClick(w.view.focusProxy(),Qt.Key_Tab);QTest.keyClick(w.view.focusProxy(),Qt.Key_Tab);spin(1);ready()
 assert len(root().xpath('//w:tab',namespaces={'w':W}))==2
 assert '\u200b' not in ''.join(root().itertext())
 report['rapidDoubleTab']=True
 report['tableNavigationAndLiteralTab']=True;report['customTabAt75Percent']=widths
 report['tabCaret']=measurements
 Path('audit-output/spacing-shortcuts-test.json').write_text(json.dumps(report,indent=2))
 w.model.save('audit-output/spacing-shortcuts-saved.docx');Document('audit-output/spacing-shortcuts-saved.docx')
 assert path.read_bytes()==original
 print('PASS spacing keys, visible tab caret, repeated stops, deletion, indent/hanging/spacing/alignment, breaks, undo/redo, no blank/faded/replaced frames, DOCX reopen',flush=True)
finally:w.dirty=False;w.close();spin(.1)

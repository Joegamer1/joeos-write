"""JoeOS Write desktop shell and edit transactions for original DOCX packages."""
import base64,hashlib,json,os,sys,time
# Grayscale glyph edges stay consistent with PDF previews and avoid colored
# ClearType fringes when a page is scaled. This affects display, never OOXML.
_chromium_flags=os.environ.get('QTWEBENGINE_CHROMIUM_FLAGS','')
if '--disable-lcd-text' not in _chromium_flags.split():
    os.environ['QTWEBENGINE_CHROMIUM_FLAGS']=(_chromium_flags+' --disable-lcd-text').strip()
from pathlib import Path
from PySide6.QtCore import QObject,Signal,Slot,QUrl,QTimer,Qt
from PySide6.QtGui import QIcon,QFontDatabase
from PySide6.QtWidgets import QApplication,QMainWindow,QFileDialog,QMessageBox
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import QWebEnginePage,QWebEngineProfile,QWebEngineUrlRequestInterceptor,QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView
from preserved_docx import PreservedDocx
from layout_metrics import analyze_layout
from live_diagnostics import write_diagnostics
from local_fonts import LocalFonts
from text_editing import replace_range, format_range, empty_paragraphs, remap_format_selection, paragraph_format

ASSETS=Path(getattr(sys,'_MEIPASS',Path(__file__).parent))/'assets'
def digest(data):return hashlib.sha256(data).hexdigest()
def same_path(a,b):return os.path.normcase(str(Path(a).resolve()))==os.path.normcase(str(Path(b).resolve()))

class LocalOnly(QWebEngineUrlRequestInterceptor):
    def interceptRequest(self,info):
        url=info.requestUrl();scheme=url.scheme()
        if scheme in ('data','blob','qrc'):return
        if scheme=='file':
            try:
                Path(url.toLocalFile()).resolve().relative_to(ASSETS.resolve());return
            except ValueError:pass
        info.block(True)

class Page(QWebEnginePage):
    def acceptNavigationRequest(self,url,kind,main):
        return url.scheme()=='file' and Path(url.toLocalFile()).resolve()==(ASSETS/'editor/index.html').resolve()
    def createWindow(self,kind):return None
    def javaScriptConsoleMessage(self,level,message,line,source):
        if os.environ.get('JOEOS_DEBUG'):print('JS:',message,flush=True)

class Bridge(QObject):
    documentChanged=Signal(str);statusChanged=Signal(str);message=Signal(str)
    def __init__(self,window):super().__init__(window);self.w=window
    @Slot()
    def ready(self):self.w.publish()
    @Slot()
    def refreshPreview(self):self.w.publish()
    @Slot(int,int)
    def rendered(self,pages,texts):
        self.w.render_result={'pages':pages,'text_segments':texts};self.w.render_error=None
    @Slot(str)
    def renderFailed(self,error):self.w.render_error=error
    @Slot(str)
    def renderDiagnostics(self,value):
        try:write_diagnostics(self.w,json.loads(value))
        except Exception:pass
    @Slot(str,str,result=bool)
    def editText(self,identity,text):return self.w.change('text',identity,text)
    @Slot(str,str,str,str,result=bool)
    def typeText(self,identity,text,before,after):
        return self.w.change('text',identity,text,json.loads(before),json.loads(after))
    @Slot(str,result=bool)
    def editRange(self,value):
        return self.w.edit_range(json.loads(value))
    @Slot(str,str,str,result=bool)
    def formatText(self,identity,key,value):
        try:return self.w.change('format',identity,(key,json.loads(value)))
        except Exception as e:self.message.emit(str(e));return False
    @Slot(str,result=bool)
    def addParagraph(self,identity):return self.w.change('paragraph',identity,None)
    @Slot()
    def newDocument(self):self.w.new_document()
    @Slot()
    def openDocument(self):self.w.open_document()
    @Slot()
    def saveDocument(self):self.w.save_document()
    @Slot()
    def saveAs(self):self.w.save_document(True)
    @Slot()
    def undo(self):self.w.travel(False)
    @Slot()
    def redo(self):self.w.travel(True)
    @Slot()
    def help(self):
        QMessageBox.information(self.w,'Preserving your Word document',
        'An unchanged save is byte-for-byte identical to the opened DOCX. Typing and range edits patch touched text, runs or paragraphs; other package entries remain preserved.\n\n'
        'Click to place a cursor. Enter starts a paragraph, Shift+Enter inserts a line break, and Tab inserts a tab. Ctrl+B/I/U format selected characters or toggle subsequent typing. Ctrl+Z/Y undo/redo. Some protected structures and table operations remain restricted.\n\n'
        'Page display is approximate. Browser-based layout is not Microsoft Word pagination; missing fonts, floating objects and unsupported features may display differently. Those display limitations do not discard the original content on save. No external editor, account or document upload is used. Save regularly; crash recovery is not yet included.')

class Editor(QMainWindow):
    def __init__(self,path=None):
        super().__init__();self.resize(1280,900);self.setMinimumSize(900,600);self.setWindowIcon(QIcon(str(ASSETS/'write-icon.ico')))
        self.model=PreservedDocx.blank();self.path=None;self.disk_hash=None;self.imported=False;self.saved_digest=digest(self.model.bytes())
        self.undo_stack=[];self.redo_stack=[];self.undo_positions=[];self.redo_positions=[];self.last_edit=None;self.dirty=False;self.render_result=None;self.render_error=None
        self.local_fonts=LocalFonts();self.document_key=str(time.time_ns())
        self.pending_caret=None
        self.view=QWebEngineView(self);self.profile=QWebEngineProfile(self)
        self.interceptor=LocalOnly(self.profile);self.profile.setUrlRequestInterceptor(self.interceptor)
        self.page=Page(self.profile,self);self.view.setPage(self.page)
        settings=self.page.settings();settings.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls,False)
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls,True)
        settings.setAttribute(QWebEngineSettings.JavascriptCanOpenWindows,False)
        self.channel=QWebChannel(self.page);self.bridge=Bridge(self);self.channel.registerObject('bridge',self.bridge);self.page.setWebChannel(self.channel)
        self.setCentralWidget(self.view)
        if path:
            try:self._load(path)
            except Exception as e:QTimer.singleShot(0,lambda message=str(e):self.error(message))
        self.view.setUrl(QUrl.fromLocalFile(str(ASSETS/'editor/index.html')))

    def error(self,text):QMessageBox.critical(self,'JoeOS Write',text)
    def _load(self,path):
        self.document_key=str(time.time_ns())
        model=PreservedDocx.open(path)
        self.model=model;self.path=str(path);self.disk_hash=digest(model.original);self.imported=True;self.saved_digest=digest(model.original)
        self.undo_stack=[];self.redo_stack=[];self.dirty=False;self.last_edit=None

    def publish(self):
        installed=QFontDatabase.families();known={f.casefold() for f in installed}
        requested=set(self.model.fonts)|{v['font'] for v in self.model.format_edits.values() if 'font' in v}
        faces=self.local_fonts.faces(requested)
        resolved={f['family'].casefold() for f in faces}
        missing=sorted(f for f in requested if f.casefold() not in known|resolved)
        notice='Original Word structure retained. Click text to edit; unsupported objects stay in the file.'
        if missing:notice+=' Missing fonts: '+', '.join(missing[:8])+'. Display uses substitutes; saved font names are preserved.'
        fallbacks=[r['requested']+' -> '+r['sourceFamily'] for r in self.local_fonts.resolutions if r['compatibility']]
        if fallbacks:notice+=' Display font substitutions: '+', '.join(fallbacks)+'. Original font names remain unchanged on save.'
        if self.model.readonly:notice='Signed or protected document: read-only preservation mode.'
        self.render_result=None;self.render_error=None
        preview=self.model.bytes(preview=True)
        if self.model.preview_notices:notice+=' '+' '.join(self.model.preview_notices)
        payload={'data':base64.b64encode(preview).decode(),'texts':self.model.metadata(),'fonts':installed,'documentFonts':sorted(self.model.fonts),'layout':analyze_layout(self.model.bytes()),'notice':notice}
        payload['localFontFaces']=faces
        payload['fontResolutions']=self.local_fonts.resolutions
        payload['emptyParagraphs']=empty_paragraphs(self.model)
        payload['documentKey']=self.document_key
        payload['caret']=self.pending_caret
        self.pending_caret=None
        self.bridge.documentChanged.emit(json.dumps(payload,ensure_ascii=True));self.status()

    def edit_range(self,request):
        before=self.model.snapshot()
        try:
            if request.get('action')=='paragraph':
                self.model=paragraph_format(self.model,request['start'],request['end'],request['text'])
                self.pending_caret={'start':request['start'],'end':request['end']}
            elif request.get('action')=='format':
                previous=self.model
                self.model=format_range(previous,request['start'],request['end'],request['marks'])
                self.pending_caret=remap_format_selection(previous,self.model,request['start'],request['end'])
            else:
                self.model,self.pending_caret=replace_range(self.model,request['start'],request['end'],request.get('text',''),request.get('marks'))
            self.undo_stack.append(before);self.undo_stack=self.undo_stack[-50:]
            self.undo_positions.append(({'start':request['start'],'end':request['end']},self.pending_caret))
            self.undo_positions=self.undo_positions[-50:];self.redo_positions=[]
            self.redo_stack=[];self.last_edit=None
            self.dirty=digest(self.model.bytes())!=self.saved_digest
            self.publish();return True
        except Exception as e:
            self.model=PreservedDocx.from_snapshot(before)
            self.bridge.message.emit('Edit not applied: '+str(e));return False

    def status(self):
        name=Path(self.path).name if self.path else 'Untitled.docx'
        self.setWindowTitle(('* ' if self.dirty else '')+name+' — JoeOS Write')
        text_count=len(self.model.text_edits);format_count=len(self.model.format_edits)
        summary=f'{len(self.model.texts)} text segments • Original package retained'
        changes=f'{text_count} text segment(s) changed\n{format_count} run(s) formatted'
        if self.undo_stack and not text_count and not format_count:changes='Structural edits retained in memory'
        self.bridge.statusChanged.emit(json.dumps({'name':name,'dirty':self.dirty,'summary':summary,'changes':changes if self.dirty else 'No unsaved changes','undo':bool(self.undo_stack),'redo':bool(self.redo_stack)}))

    def change(self,kind,identity,value,before_cursor=None,after_cursor=None):
        before=self.model.snapshot()
        try:
            if kind=='text':self.model.set_text(identity,value)
            elif kind=='format':self.model.set_format(identity,*value)
            elif kind=='paragraph':self.model=self.model.append_paragraph(identity)
            else:raise ValueError('Unknown edit.')
            after=self.model.snapshot()
            if after==before:return True
            now=time.monotonic();group=(kind,identity)
            if not(kind=='text' and self.last_edit and self.last_edit[0]==group and now-self.last_edit[1]<1):
                self.undo_stack.append(before);self.undo_stack=self.undo_stack[-50:]
                self.undo_positions.append((before_cursor,after_cursor));self.undo_positions=self.undo_positions[-50:]
            elif self.undo_positions:
                self.undo_positions[-1]=(self.undo_positions[-1][0],after_cursor)
            self.redo_positions=[]
            self.last_edit=(group,now);self.redo_stack=[]
            self.dirty=digest(self.model.bytes())!=self.saved_digest
            if kind!='text':self.publish()
            else:self.status()
            return True
        except Exception as e:
            self.model=PreservedDocx.from_snapshot(before);self.bridge.message.emit('Edit not applied: '+str(e));return False

    def travel(self,redo):
        source=self.redo_stack if redo else self.undo_stack;target=self.undo_stack if redo else self.redo_stack
        if not source:return
        positions=self.redo_positions if redo else self.undo_positions
        opposite=self.undo_positions if redo else self.redo_positions
        position=positions.pop() if positions else (None,None);opposite.append(position)
        self.pending_caret=position[1] if redo else position[0]
        target.append(self.model.snapshot());self.model=PreservedDocx.from_snapshot(source.pop());self.last_edit=None
        self.dirty=digest(self.model.bytes())!=self.saved_digest;self.publish()

    def confirm_changes(self):
        if not self.dirty:return True
        answer=QMessageBox.question(self,'Unsaved changes','Save changes before continuing?',QMessageBox.Save|QMessageBox.Discard|QMessageBox.Cancel,QMessageBox.Save)
        if answer==QMessageBox.Cancel:return False
        return self.save_document() if answer==QMessageBox.Save else True

    def new_document(self):
        if not self.confirm_changes():return
        self.model=PreservedDocx.blank();self.path=None;self.disk_hash=None;self.imported=False;self.saved_digest=digest(self.model.bytes());self.dirty=False
        self.undo_stack=[];self.redo_stack=[];self.undo_positions=[];self.redo_positions=[];self.last_edit=None;self.publish()

    def open_document(self,path=None):
        if not self.confirm_changes():return
        if not path:path=QFileDialog.getOpenFileName(self,'Open Word document','','Word documents (*.docx)')[0]
        if not path:return
        try:self._load(path);self.publish()
        except Exception as e:self.error('Could not open the document.\n'+str(e))

    def save_document(self,save_as=False):
        target=self.path
        if save_as or not target or self.imported:
            suggested=str(Path(target).with_name(Path(target).stem+' - edited.docx')) if target and self.imported else target or 'Untitled.docx'
            target=QFileDialog.getSaveFileName(self,'Save Word document',suggested,'Word document (*.docx)')[0]
            if not target:return False
            if not target.lower().endswith('.docx'):
                target+='.docx'
                if Path(target).exists() and QMessageBox.question(self,'Replace existing file?','The final filename exists. Replace it?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return False
        try:
            if self.path and same_path(target,self.path):
                if not Path(target).exists() or digest(Path(target).read_bytes())!=self.disk_hash:
                    raise ValueError('The source file changed or was removed outside JoeOS Write. Use Save As to keep both versions.')
            self.model.save(target);self.disk_hash=digest(Path(target).read_bytes());self.saved_digest=self.disk_hash
            self.path=target;self.imported=False;self.dirty=False;self.last_edit=None;self.status()
            message='Saved locally. Unedited document parts retained.'
            self.bridge.message.emit(message);return True
        except Exception as e:self.error('Could not save. Your edits remain open.\n'+str(e));return False

    def closeEvent(self,event):
        if self.confirm_changes():event.accept()
        else:event.ignore()

def main():
    app=QApplication(sys.argv);app.setApplicationName('JoeOS Write');app.setStyle('Fusion')
    path=sys.argv[1] if len(sys.argv)>1 and not sys.argv[1].startswith('--') else None
    if len(sys.argv)>3 and sys.argv[1]=='--smoke-test':path=sys.argv[3]
    window=Editor(path);window.show()
    if len(sys.argv)>1 and sys.argv[1]=='--smoke-test':
        output=Path(sys.argv[2] if len(sys.argv)>2 else '.tmp');output.mkdir(parents=True,exist_ok=True)
        timer=QTimer(window);started=time.monotonic();settled=[None]
        def verify():
            if window.render_result:
                if settled[0] is None:settled[0]=time.monotonic();return
                if time.monotonic()-settled[0]<.6:return
                window.grab().save(str(output/'preservation-ui.png'))
                window.model.save(output/'preservation-smoke.docx')
                result={**window.render_result,'unchanged_save':(output/'preservation-smoke.docx').read_bytes()==window.model.original}
                (output/'preservation-smoke.json').write_text(json.dumps(result));app.exit(0)
            elif window.render_error or time.monotonic()-started>40:
                (output/'preservation-error.txt').write_text(window.render_error or 'Timed out');app.exit(1)
        timer.timeout.connect(verify);timer.start(250)
    sys.exit(app.exec())

if __name__=='__main__':main()

"""JoeOS entry point; document editing is delegated to local LibreOffice Writer."""
import os
import sys
import subprocess
import shutil
from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QFontDatabase
from PySide6.QtWidgets import QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QFileDialog,QMessageBox
from docx import Document
from docx.shared import Inches,Pt

ROOT=Path(sys.executable).parent if getattr(sys,'frozen',False) else Path(__file__).parent
ASSETS=Path(getattr(sys,'_MEIPASS',Path(__file__).parent))/'assets'

def find_writer():
    candidates=[Path(os.environ.get('PROGRAMFILES','C:/Program Files'))/'LibreOffice/program/soffice.exe',Path(os.environ.get('PROGRAMFILES(X86)','C:/Program Files (x86)'))/'LibreOffice/program/soffice.exe']
    return next((p for p in candidates if p.is_file()),None)

def launch_writer(path):
    engine=find_writer()
    if engine is None:raise RuntimeError('LibreOffice is not installed. Install it from libreoffice.org, then reopen JoeOS Write.')
    if not Path(path).is_file():raise FileNotFoundError('The selected document no longer exists.')
    # Direct argument list: Unicode, spaces, and shell characters remain literal.
    return subprocess.Popen([str(engine),'--writer','--nologo',str(Path(path).resolve())],close_fds=True)

class WriteHome(QMainWindow):
    def __init__(self):
        super().__init__();self.setWindowTitle('JoeOS Write');self.resize(790,570);self.setMinimumSize(650,480)
        self.setWindowIcon(QIcon(str(ASSETS/'write-icon.ico')))
        self.setStyleSheet('QWidget {background:#eadcc7;color:#292722;font-family:"Segoe UI";font-size:14px;} QPushButton {background:#faf7f0;border:1px solid #88745e;border-radius:8px;padding:15px;text-align:left;} QPushButton:hover {background:#f6d2aa;}')
        body=QWidget();box=QVBoxLayout(body);box.setContentsMargins(36,30,36,30);box.setSpacing(16)
        header=QHBoxLayout();icon=QLabel();icon.setPixmap(QIcon(str(ASSETS/'write-icon.ico')).pixmap(68,68));header.addWidget(icon)
        title=QLabel('JoeOS <span style="color:#ff6d03">/</span> Write');title.setStyleSheet('font-size:30px;font-weight:600;');header.addWidget(title);header.addStretch();box.addLayout(header)
        intro=QLabel('Your documents. Their original structure.\nPowered by local LibreOffice Writer.');intro.setWordWrap(True);box.addWidget(intro)
        for label,callback in [('New Word document',self.new_document),('Open a document',self.open_document),('Open an editable copy — keep the original',self.open_copy)]:
            button=QPushButton(label);button.clicked.connect(callback);box.addWidget(button)
        info=QLabel('Editing opens in Writer’s full desktop interface. Use Ctrl+S to save, or Ctrl+Shift+S to save a copy. Keep the Word (.docx) format when prompted.');info.setWordWrap(True);box.addWidget(info)
        caveat=QLabel('Word compatibility is improved, but identical pagination is not guaranteed. Missing fonts and complex Word features still need review.');caveat.setWordWrap(True);caveat.setStyleSheet('color:#5c5044;font-size:12px;');box.addWidget(caveat)
        box.addStretch();self.status=QLabel('Local only • No document upload • Working name');self.status.setWordWrap(True);self.status.setStyleSheet('color:#5c5044;font-size:12px;');box.addWidget(self.status);self.setCentralWidget(body)

    def error(self,e):QMessageBox.critical(self,'JoeOS Write',str(e))
    def open_path(self,path):
        try:
            launch_writer(path);self.status.setText('Opened in Writer: '+str(path))
        except Exception as e:self.error(e)

    def select_document(self):
        return QFileDialog.getOpenFileName(self,'Open a document','','Word documents (*.docx *.doc);;OpenDocument (*.odt);;All files (*)')[0]

    def open_document(self):
        path=self.select_document()
        if path:self.open_path(path)

    def open_copy(self):
        if find_writer() is None:return self.error('Install LibreOffice before creating an editable copy.')
        source=self.select_document()
        if not source:return
        p=Path(source);target,_=QFileDialog.getSaveFileName(self,'Save editable copy',str(p.with_name(p.stem+' - edited'+p.suffix)),'Document (*'+p.suffix+')')
        if not target:return
        if not Path(target).suffix:
            target+=p.suffix
            if Path(target).exists() and QMessageBox.question(self,'Replace file?','The final filename already exists. Replace it?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return
        try:
            if os.path.normcase(os.path.abspath(source))==os.path.normcase(os.path.abspath(target)):
                raise ValueError('Choose a different name to preserve the original.')
            import tempfile
            fd,tmp=tempfile.mkstemp(suffix=p.suffix,dir=Path(target).parent);os.close(fd)
            try:shutil.copy2(source,tmp);os.replace(tmp,target)
            finally:
                if os.path.exists(tmp):os.unlink(tmp)
            self.open_path(target)
        except Exception as e:self.error(e)

    def new_document(self):
        if find_writer() is None:return self.error('Install LibreOffice before creating a document.')
        target,_=QFileDialog.getSaveFileName(self,'Create a Word document','Untitled.docx','Word document (*.docx)')
        if not target:return
        if not target.lower().endswith('.docx'):
            target+='.docx'
            if Path(target).exists() and QMessageBox.question(self,'Replace file?','The final filename already exists. Replace it?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return
        try:
            d=Document();d.styles['Normal'].font.name='Calibri';d.styles['Normal'].font.size=Pt(11)
            s=d.sections[0];s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Inches(1)
            import tempfile
            fd,tmp=tempfile.mkstemp(suffix='.docx',dir=Path(target).parent);os.close(fd)
            try:d.save(tmp);os.replace(tmp,target)
            finally:
                if os.path.exists(tmp):os.unlink(tmp)
            self.open_path(target)
        except Exception as e:self.error(e)

def main():
    app=QApplication(sys.argv);app.setStyle('Fusion');app.setApplicationName('JoeOS Write');window=WriteHome();window.show()
    if len(sys.argv)>1 and sys.argv[1]=='--smoke-test':
        assert find_writer() is not None
        app.processEvents();window.grab().save(str(ROOT/'launcher-smoke.png'));return
    if len(sys.argv)>1:window.open_path(sys.argv[1])
    sys.exit(app.exec())

if __name__=='__main__':main()

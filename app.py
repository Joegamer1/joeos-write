import os
import sys
from pathlib import Path
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import (QAction, QIcon, QFont, QTextCharFormat, QTextCursor,
                          QTextListFormat, QTextTableFormat, QTextImageFormat, QImage, QTextDocument)
from PySide6.QtWidgets import (QApplication, QMainWindow, QTextEdit, QToolBar, QFileDialog,
    QMessageBox, QFontComboBox, QComboBox, QColorDialog, QInputDialog, QLabel, QWidget,
    QVBoxLayout, QHBoxLayout, QPushButton, QFrame)
from docx_io import load_docx, save_docx

APP_NAME = 'JoeOS Write'  # Working name; branding can be renamed here.
ASSETS = Path(getattr(sys, '_MEIPASS', Path(__file__).parent)) / 'assets'


class Editor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.path = None
        self.imported = False
        self.disk_stamp = None
        self.resize(1180, 850)
        self.setMinimumSize(760, 520)
        self.setWindowIcon(QIcon(str(ASSETS/'write-icon.ico')))
        self.editor = QTextEdit()
        self.editor.setAcceptRichText(True)
        self.editor.setFont(QFont('Calibri',11))
        self.editor.setStyleSheet('QTextEdit {background:#fff; color:#20242c; border:0; padding:38px; selection-background-color:#ffd0ae; selection-color:#171717;}')
        self.editor.document().setDocumentMargin(24)
        self.editor.setTabStopDistance(48)
        shell = QWidget(); layout = QVBoxLayout(shell); layout.setContentsMargins(26,16,26,16)
        header = QHBoxLayout()
        brand = QLabel(); brand.setPixmap(QIcon(str(ASSETS/'M2-master.svg')).pixmap(36,36))
        header.addWidget(brand)
        title = QLabel('JoeOS <span style="color:#ff6d03">/</span> Write')
        title.setStyleSheet('font-size:22px; font-weight:600;')
        header.addWidget(title); header.addStretch()
        label = QLabel('LOCAL WORKSPACE  •  WORKING NAME'); label.setStyleSheet('color:#5c5044; font-size:11px;')
        header.addWidget(label); layout.addLayout(header)
        self.notice = QLabel('Your words, on your computer. Documents stay local.'); self.notice.setWordWrap(True)
        self.notice.setStyleSheet('color:#5c5044; padding:4px 0 8px;')
        layout.addWidget(self.notice); layout.addWidget(self.editor)
        self.setCentralWidget(shell)
        self.setStyleSheet('''QMainWindow,QWidget {background:#eadcc7;color:#292722;font-family:"Segoe UI";}
            QMenuBar,QMenu,QToolBar {background:#eee5d7;color:#292722;}
            QMenu::item:selected,QToolButton:hover,QPushButton:hover {background:#edd5b2;}
            QToolBar {spacing:5px;padding:7px;border-bottom:1px solid #88745e;}
            QToolButton,QPushButton {padding:6px 9px;border-radius:4px;}
            QToolButton:checked {background:#ff6d03;color:#10151e;}
            QComboBox {background:#faf7f0;padding:5px;border:1px solid #88745e;border-radius:4px;}
            QStatusBar {background:#292722;color:#fff7eb;}
            QLineEdit {background:#fff;color:#222;padding:6px;}
        ''')
        filemenu = self.menuBar().addMenu('&File')
        editmenu = self.menuBar().addMenu('&Edit')
        insertmenu = self.menuBar().addMenu('&Insert')
        helpmenu = self.menuBar().addMenu('&Help')
        toolbar = QToolBar('Document'); toolbar.setMovable(False); self.addToolBar(toolbar)

        def action(menu, text, shortcut, fn, bar=None, check=False):
            a = QAction(text,self); a.setCheckable(check)
            if shortcut: a.setShortcut(shortcut)
            a.triggered.connect(fn); menu.addAction(a)
            if bar: bar.addAction(a)
            return a
        action(filemenu,'New','Ctrl+N',self.new,toolbar)
        action(filemenu,'Open…','Ctrl+O',self.open,toolbar)
        action(filemenu,'Save','Ctrl+S',self.save,toolbar)
        action(filemenu,'Save As…','Ctrl+Shift+S',lambda: self.save(True),toolbar)
        filemenu.addSeparator(); action(filemenu,'Exit','Alt+F4',self.close)
        action(editmenu,'Undo','Ctrl+Z',self.editor.undo,toolbar)
        action(editmenu,'Redo','Ctrl+Y',self.editor.redo,toolbar)
        action(editmenu,'Cut','Ctrl+X',self.editor.cut)
        action(editmenu,'Copy','Ctrl+C',self.editor.copy)
        action(editmenu,'Paste','Ctrl+V',self.editor.paste)
        action(editmenu,'Select All','Ctrl+A',self.editor.selectAll)
        action(editmenu,'Find…','Ctrl+F',self.find)
        action(insertmenu,'Image…',None,self.insert_image)
        action(insertmenu,'Table…',None,self.insert_table)
        action(helpmenu,'DOCX compatibility',None,self.compatibility)
        self.addToolBarBreak()
        fmtbar = QToolBar('Formatting'); fmtbar.setMovable(False); self.addToolBar(fmtbar)
        fm = self.menuBar().addMenu('F&ormat')
        self.fontbox = QFontComboBox(); self.fontbox.setCurrentFont(QFont('Calibri'))
        self.fontbox.currentFontChanged.connect(lambda f:self.apply_format('font',f))
        fmtbar.addWidget(self.fontbox)
        self.sizebox = QComboBox(); self.sizebox.addItems(['8','9','10','11','12','14','16','18','24','32','48','72'])
        self.sizebox.setCurrentText('11'); self.sizebox.setEditable(True); self.sizebox.setMaximumWidth(70)
        self.sizebox.activated.connect(lambda:self.set_size())
        self.sizebox.lineEdit().editingFinished.connect(self.set_size); fmtbar.addWidget(self.sizebox)
        self.bold = action(fm,'Bold','Ctrl+B',lambda v:self.apply_format('bold',v),fmtbar,True)
        self.italic = action(fm,'Italic','Ctrl+I',lambda v:self.apply_format('italic',v),fmtbar,True)
        self.underline = action(fm,'Underline','Ctrl+U',lambda v:self.apply_format('underline',v),fmtbar,True)
        action(fm,'Color…',None,self.color,fmtbar)
        fmtbar.addSeparator()
        for text,alignment in [('Left',Qt.AlignLeft),('Center',Qt.AlignHCenter),('Right',Qt.AlignRight),('Justify',Qt.AlignJustify)]:
            action(fm,text,None,lambda checked=False,a=alignment:self.editor.setAlignment(a),fmtbar)
        action(fm,'Bullets',None,lambda:self.make_list(False),fmtbar)
        action(fm,'Numbering',None,lambda:self.make_list(True),fmtbar)
        action(fm,'Clear list',None,self.clear_list)
        self.editor.currentCharFormatChanged.connect(self.sync_format)
        self.editor.textChanged.connect(self.update_status)
        self.editor.document().modificationChanged.connect(self.update_status)
        self.editor.document().setModified(False)
        self.update_status()

    def apply_format(self, kind, value):
        cf = QTextCharFormat()
        if kind=='font': cf.setFontFamilies([value.family()])
        elif kind=='bold': cf.setFontWeight(QFont.Bold if value else QFont.Normal)
        elif kind=='italic': cf.setFontItalic(value)
        elif kind=='underline': cf.setFontUnderline(value)
        elif kind=='size': cf.setFontPointSize(value)
        elif kind=='color': cf.setForeground(value)
        self.editor.mergeCurrentCharFormat(cf); self.editor.setFocus()

    def sync_format(self, cf):
        for control in [self.fontbox,self.sizebox,self.bold,self.italic,self.underline]: control.blockSignals(True)
        self.bold.setChecked(cf.fontWeight()>=QFont.Bold); self.italic.setChecked(cf.fontItalic()); self.underline.setChecked(cf.fontUnderline())
        self.fontbox.setCurrentFont(cf.font()); self.sizebox.setCurrentText(str(round(cf.fontPointSize() or 11)))
        for control in [self.fontbox,self.sizebox,self.bold,self.italic,self.underline]: control.blockSignals(False)

    def set_size(self):
        try:
            size = float(self.sizebox.currentText())
            if 6<=size<=144: self.apply_format('size',size)
        except ValueError: pass

    def color(self):
        c = QColorDialog.getColor(self.editor.textColor(),self)
        if c.isValid(): self.apply_format('color',c)

    def make_list(self, numbered):
        f = QTextListFormat(); f.setStyle(QTextListFormat.ListDecimal if numbered else QTextListFormat.ListDisc)
        self.editor.textCursor().createList(f)

    def clear_list(self):
        c = self.editor.textCursor(); f=c.blockFormat(); f.setObjectIndex(-1); c.setBlockFormat(f)

    def insert_table(self):
        rows,ok=QInputDialog.getInt(self,'Insert table','Rows',3,1,50)
        if not ok:return
        cols,ok=QInputDialog.getInt(self,'Insert table','Columns',3,1,12)
        if ok:
            f=QTextTableFormat(); f.setBorder(1); f.setCellPadding(6)
            self.editor.textCursor().insertTable(rows,cols,f)

    def insert_image(self):
        path,_=QFileDialog.getOpenFileName(self,'Insert image','','Images (*.png *.jpg *.jpeg *.bmp)')
        if not path:return
        image=QImage(path)
        if image.isNull():return self.error('Cannot read this image.')
        import uuid
        key='image-'+str(uuid.uuid4()); self.editor.document().addResource(QTextDocument.ImageResource,QUrl(key),image)
        f=QTextImageFormat(); f.setName(key); f.setWidth(min(image.width(),600)); f.setHeight(image.height()*f.width()/image.width())
        self.editor.textCursor().insertImage(f)

    def find(self):
        text,ok=QInputDialog.getText(self,'Find','Find text:')
        if ok and text and not self.editor.find(text):
            c=self.editor.textCursor(); c.movePosition(QTextCursor.Start); self.editor.setTextCursor(c)
            if not self.editor.find(text): QMessageBox.information(self,'Find','No matches found.')

    def update_status(self,*args):
        dirty=self.editor.document().isModified()
        self.setWindowTitle(('* ' if dirty else '')+(Path(self.path).name if self.path else 'Untitled')+' — '+APP_NAME)
        self.statusBar().showMessage(f'{len(self.editor.toPlainText().split()):,} words  •  '+('Unsaved changes' if dirty else 'Ready')+'  •  Local only')

    def confirm_changes(self):
        if not self.editor.document().isModified():return True
        choice=QMessageBox.question(self,'Unsaved changes','Save your changes before continuing?',QMessageBox.Save|QMessageBox.Discard|QMessageBox.Cancel,QMessageBox.Save)
        if choice==QMessageBox.Cancel:return False
        return self.save() if choice==QMessageBox.Save else True

    def new(self):
        if not self.confirm_changes():return
        self.editor.clear(); self.path=None; self.imported=False; self.disk_stamp=None
        self.editor.document().setModified(False)
        self.notice.setText('Your words, on your computer. Documents stay local.'); self.update_status()

    def open(self, checked=False, path=None):
        if not self.confirm_changes():return
        if not path:path,_=QFileDialog.getOpenFileName(self,'Open Word document','','Word documents (*.docx)')
        if not path:return
        try:
            document,warnings=load_docx(path)
            document.setParent(self.editor)
            self.editor.setDocument(document)
            document.modificationChanged.connect(self.update_status)
            self.path=str(path); self.imported=True; self.disk_stamp=self.stamp(path)
            self.notice.setText('Imported DOCX • Save creates an edited copy. Advanced Word layout may change.')
            self.update_status()
            if warnings:
                QMessageBox.information(self,'Word compatibility','Opened in the editable layout. Features that may change or be omitted:\n\n'+', '.join(warnings)+'.\n\nSave defaults to a separate copy. See Help → DOCX compatibility.')
        except Exception as e:self.error('Could not open this document.\n'+str(e))

    @staticmethod
    def stamp(path):
        import hashlib
        return hashlib.sha256(Path(path).read_bytes()).hexdigest()

    def save(self, save_as=False):
        target=self.path
        if save_as or not target or self.imported:
            suggested=str(Path(target).with_name(Path(target).stem+' - edited.docx')) if target and self.imported else target or 'Untitled.docx'
            target,_=QFileDialog.getSaveFileName(self,'Save Word document',suggested,'Word document (*.docx)')
            if not target:return False
            if not target.lower().endswith('.docx'):target+='.docx'
        if self.imported and self.path and os.path.normcase(os.path.abspath(target))==os.path.normcase(os.path.abspath(self.path)):
            if QMessageBox.warning(self,'Replace original?','Complex Word features may be lost. Replace the original document?',QMessageBox.Yes|QMessageBox.Cancel,QMessageBox.Cancel)!=QMessageBox.Yes:return False
        try:
            if self.path and os.path.normcase(os.path.abspath(target))==os.path.normcase(os.path.abspath(self.path)) and Path(target).exists() and self.stamp(target)!=self.disk_stamp:
                self.error('The file changed outside JoeOS Write. Use Save As to keep both versions.');return False
            save_docx(self.editor.document(),target)
            self.path=target;self.imported=False;self.disk_stamp=self.stamp(target)
            self.editor.document().setModified(False)
            self.notice.setText('Saved locally • '+target);self.update_status();return True
        except Exception as e:self.error('Could not save. Your edits are still open.\n'+str(e));return False

    def compatibility(self):
        QMessageBox.information(self,'DOCX compatibility','Supports text, fonts, sizes, bold, italic, underline, strikeout, text colors, alignment, basic lists, simple tables and inline PNG/JPEG images.\n\nThis is a personal editor, not a full Word layout engine. Page sizes/margins, pagination, sections, columns, headers/footers, footnotes, comments, tracked changes, fields, text boxes, shapes, charts, hyperlinks, advanced styles, merged/nested tables and precise list numbering are not preserved reliably. Some content inside unsupported objects may be omitted.\n\nReview important documents in Word. Imports default to a separate edited copy. Legacy .doc and macro-enabled .docm files are not supported. No cloud sync, telemetry or account is required. Save regularly; crash recovery is not included.')

    def error(self,text):QMessageBox.critical(self,'JoeOS Write',text)
    def closeEvent(self,event):
        if self.confirm_changes():event.accept()
        else:event.ignore()


def main():
    app=QApplication(sys.argv);app.setApplicationName(APP_NAME);app.setStyle('Fusion')
    window=Editor();window.show()
    if len(sys.argv)>2 and sys.argv[1]=='--smoke-test':
        output=Path(sys.argv[2]);output.mkdir(parents=True,exist_ok=True)
        window.editor.setHtml('<h1>JoeOS Write</h1><p>Packaged app <b>DOCX smoke test</b>.</p>')
        save_docx(window.editor.document(),output/'packaged.docx')
        doc,_=load_docx(output/'packaged.docx')
        assert 'DOCX smoke test' in doc.toPlainText()
        window.editor.setDocument(doc)
        app.processEvents();window.grab().save(str(output/'packaged.png'))
        (output/'packaged-test.txt').write_text('PASS: packaged executable started, saved DOCX, reopened content, and rendered its window.\n',encoding='utf-8')
        return
    if len(sys.argv)>1:window.open(path=sys.argv[1])
    sys.exit(app.exec())

if __name__=='__main__':main()

import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
from PySide6.QtGui import QImage, QColor, QTextCursor
from docx_io import load_docx, save_docx
from app import Editor

APP=QApplication.instance() or QApplication([])
QMessageBox.critical = lambda parent,title,text: (_ for _ in ()).throw(RuntimeError(text))
ROOT=Path(__file__).parent/'test-output'
ROOT.mkdir(exist_ok=True)


class RoundTrip(unittest.TestCase):
    def test_fractional_font_size(self):
        from docx.oxml.ns import qn
        source=Document(); r=source.add_paragraph().add_run('Fractional size')
        r.font.size=Pt(8)
        r._r.rPr.sz.set(qn('w:val'),'16.5')
        source.save(ROOT/'fractional.docx')
        qt,_=load_docx(ROOT/'fractional.docx')
        self.assertEqual(qt.begin().begin().fragment().charFormat().fontPointSize(),8.25)
        save_docx(qt,ROOT/'fractional-saved.docx')
        self.assertEqual(Document(ROOT/'fractional-saved.docx').paragraphs[0].text,'Fractional size')

    def test_rich_docx_roundtrip(self):
        source=Document()
        source.add_heading('JoeOS compatibility sample',1)
        p=source.add_paragraph()
        r=p.add_run('Bold orange');r.bold=True;r.font.size=Pt(18);r.font.color.rgb=RGBColor.from_string('FF6D03')
        r=p.add_run(' • italic café — 日本語');r.italic=True;r.underline=True
        p=source.add_paragraph('Centered');p.alignment=1
        source.add_paragraph('A bullet',style='List Bullet')
        source.add_paragraph('Number one',style='List Number')
        t=source.add_table(rows=2,cols=2)
        for i,row in enumerate(t.rows):
            for j,cell in enumerate(row.cells):cell.text=f'Cell {i},{j}'
        im=QImage(80,40,QImage.Format_RGB32);im.fill(QColor('#ff6d03'));im.save(str(ROOT/'sample.png'))
        source.add_picture(str(ROOT/'sample.png'),width=Inches(1))
        source.add_paragraph('End of document')
        source.save(ROOT/'input.docx')
        qt,warnings=load_docx(ROOT/'input.docx')
        cursor=QTextCursor(qt);cursor.movePosition(QTextCursor.End);cursor.insertText(' — edited locally')
        save_docx(qt,ROOT/'roundtrip.docx')
        check=Document(ROOT/'roundtrip.docx')
        text='\n'.join(p.text for p in check.paragraphs)
        for phrase in ['JoeOS compatibility sample','Bold orange','日本語','A bullet','Number one','edited locally']:
            self.assertIn(phrase,text)
        r=next(r for p in check.paragraphs for r in p.runs if r.text=='Bold orange')
        self.assertTrue(r.bold);self.assertEqual(r.font.size.pt,18);self.assertEqual(str(r.font.color.rgb),'FF6D03')
        r=next(r for p in check.paragraphs for r in p.runs if 'café' in r.text)
        self.assertTrue(r.italic);self.assertTrue(r.underline)
        self.assertEqual(len(check.tables),1);self.assertEqual(check.tables[0].cell(1,1).text,'Cell 1,1')
        self.assertEqual(len(check.inline_shapes),1)
        self.assertEqual(next(p for p in check.paragraphs if p.text=='Centered').alignment,1)
        self.assertEqual(next(p for p in check.paragraphs if p.text=='A bullet').style.name,'List Bullet')
        with zipfile.ZipFile(ROOT/'roundtrip.docx') as z:
            self.assertIsNone(z.testzip())
            self.assertIn('[Content_Types].xml',z.namelist())
        again,_=load_docx(ROOT/'roundtrip.docx')
        save_docx(again,ROOT/'second-roundtrip.docx')
        self.assertIn('edited locally',' '.join(p.text for p in Document(ROOT/'second-roundtrip.docx').paragraphs))

    def test_atomic_failure_keeps_original(self):
        path=ROOT/'protected.docx';source=Document();source.add_paragraph('keep me');source.save(path)
        before=path.read_bytes();qt,_=load_docx(path)
        with patch('docx_io.os.replace',side_effect=PermissionError('locked')):
            with self.assertRaises(PermissionError):save_docx(qt,path)
        self.assertEqual(path.read_bytes(),before)
        self.assertEqual(list(ROOT.glob('.joeos-*')),[])

    def test_editor_save_and_conflict(self):
        w=Editor();w.editor.setPlainText('Local draft')
        with patch.object(QFileDialog,'getSaveFileName',return_value=(str(ROOT/'ui.docx'),'')):
            self.assertTrue(w.save())
        self.assertFalse(w.editor.document().isModified())
        w.editor.append('More text')
        self.assertTrue(w.save())
        self.assertIn('More text',' '.join(p.text for p in Document(ROOT/'ui.docx').paragraphs))
        changed=Document();changed.add_paragraph('External edit');changed.save(ROOT/'ui.docx')
        with patch.object(w,'error') as error:
            self.assertFalse(w.save());error.assert_called_once()
        w.editor.document().setModified(False);w.close()

    def test_cancel_and_import_copy(self):
        source=Document();source.add_paragraph('Original');source.save(ROOT/'original.docx')
        before=(ROOT/'original.docx').read_bytes()
        w=Editor()
        with patch.object(QMessageBox,'information'):w.open(path=str(ROOT/'original.docx'))
        w.editor.append('Changed')
        with patch.object(QFileDialog,'getSaveFileName',return_value=('','')):self.assertFalse(w.save())
        self.assertTrue(w.editor.document().isModified())
        with patch.object(QFileDialog,'getSaveFileName',return_value=(str(ROOT/'copy.docx'),'')):
            self.assertTrue(w.save())
        self.assertEqual(before,(ROOT/'original.docx').read_bytes())
        w.close()

    def test_invalid_open_does_not_erase_edits(self):
        w=Editor();w.editor.setPlainText('existing')
        (ROOT/'invalid.docx').write_bytes(b'not a zip')
        with patch.object(w,'confirm_changes',return_value=True),patch.object(w,'error') as error:
            w.open(path=str(ROOT/'invalid.docx'));error.assert_called_once()
        self.assertEqual(w.editor.toPlainText(),'existing')
        w.editor.document().setModified(False);w.close()


if __name__=='__main__':unittest.main(verbosity=2)

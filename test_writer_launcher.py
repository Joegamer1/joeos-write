import os,unittest,tempfile
os.environ['QT_QPA_PLATFORM']='offscreen'
from pathlib import Path
from unittest.mock import patch
from PySide6.QtWidgets import QApplication,QFileDialog
from docx import Document
import writer_app as w
APP=QApplication.instance() or QApplication([])
class Launcher(unittest.TestCase):
    def test_missing_engine(self):
        with patch.object(w,'find_writer',return_value=None):
            with self.assertRaises(RuntimeError):w.launch_writer('anything.docx')
    def test_literal_arguments(self):
        with tempfile.TemporaryDirectory(dir='audit-output') as d:
            path=Path(d)/'Unicode café & words.docx';Document().save(path)
            with patch.object(w.subprocess,'Popen') as p:w.launch_writer(path)
            self.assertEqual(p.call_args.args[0][-1],str(path.resolve()))
            self.assertNotIn('shell',p.call_args.kwargs)
    def test_new_docx(self):
        with tempfile.TemporaryDirectory(dir='audit-output') as d:
            path=Path(d)/'New.docx';window=w.WriteHome()
            with patch.object(QFileDialog,'getSaveFileName',return_value=(str(path),'')),patch.object(window,'open_path'):
                window.new_document()
            doc=Document(path);self.assertEqual(doc.sections[0].left_margin.inches,1);window.close()
    def test_copy_preserves_source_bytes(self):
        with tempfile.TemporaryDirectory(dir='audit-output') as d:
            source=Path(d)/'source.docx';target=Path(d)/'copy.docx';doc=Document();doc.add_paragraph('Original');doc.save(source);before=source.read_bytes()
            window=w.WriteHome()
            with patch.object(window,'select_document',return_value=str(source)),patch.object(QFileDialog,'getSaveFileName',return_value=(str(target),'')),patch.object(window,'open_path'):
                window.open_copy()
            self.assertEqual(source.read_bytes(),before);self.assertEqual(target.read_bytes(),before);window.close()
if __name__=='__main__':unittest.main(verbosity=2)

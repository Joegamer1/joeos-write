import io,zipfile,tempfile,unittest,copy
from pathlib import Path
from unittest.mock import patch
from lxml import etree
from docx import Document
from docx.shared import Inches,Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from preserved_docx import PreservedDocx,W

def pack(doc):b=io.BytesIO();doc.save(b);return b.getvalue()
def parts(raw):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:return {i:z.read(i) for i in z.namelist()}
def ident(model,text):return next(k for k,(p,n) in model.texts.items() if n.text==text)
def source():
    d=Document();d.sections[0].left_margin=Inches(.45);d.sections[0].header.paragraphs[0].text='HEADER'
    p=d.add_paragraph();r=p.add_run('First text');r.bold=True;r.font.size=Pt(10.5);p.add_run(' — café 日本語')
    t=d.add_table(rows=2,cols=2);t.cell(0,0).merge(t.cell(0,1)).text='Merged';t.cell(1,1).add_table(rows=1,cols=1).cell(0,0).text='Nested'
    d.add_paragraph('Item one',style='List Number');d.add_paragraph('Item two',style='List Number')
    return d

class Preservation(unittest.TestCase):
    def test_no_edit_is_byte_identical(self):
        raw=pack(source());m=PreservedDocx(raw)
        self.assertEqual(m.bytes(),raw);m.bytes(preview=True);self.assertEqual(m.bytes(),raw)
    def test_only_text_element_changed(self):
        raw=pack(source());m=PreservedDocx(raw);m.set_text(ident(m,'First text'),'Changed & <safe> 😀')
        a,b=parts(raw),parts(m.bytes())
        self.assertEqual(a.keys(),b.keys())
        self.assertEqual([k for k in a if a[k]!=b[k]],['word/document.xml'])
        old=etree.fromstring(a['word/document.xml']);new=etree.fromstring(b['word/document.xml'])
        t=next(t for t in old.iter(qn('w:t')) if t.text=='First text');t.text='Changed & <safe> 😀';t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        self.assertEqual(etree.tostring(old),etree.tostring(new))
    def test_header_edit_leaves_body_unchanged(self):
        raw=pack(source());m=PreservedDocx(raw);m.set_text(ident(m,'HEADER'),'New header')
        a,b=parts(raw),parts(m.bytes());self.assertEqual(a['word/document.xml'],b['word/document.xml'])
        self.assertEqual(len([k for k in a if a[k]!=b[k]]),1)
    def test_nested_and_merged_cells_preserved(self):
        m=PreservedDocx(pack(source()));m.set_text(ident(m,'Nested'),'Nested edited');d=Document(io.BytesIO(m.bytes()))
        self.assertIs(d.tables[0].cell(0,0)._tc,d.tables[0].cell(0,1)._tc)
        self.assertEqual(d.tables[0].cell(1,1).tables[0].cell(0,0).text,'Nested edited')
    def test_format_changes_only_selected_run(self):
        raw=pack(source());m=PreservedDocx(raw);key=ident(m,'First text');m.set_format(key,'italic',True)
        a,b=parts(raw),parts(m.bytes());self.assertEqual([k for k in a if a[k]!=b[k]],['word/document.xml'])
        d=Document(io.BytesIO(m.bytes()));r=d.paragraphs[0].runs[0]
        self.assertTrue(r.bold);self.assertTrue(r.italic);self.assertEqual(r.font.size.pt,10.5)
        self.assertIsNone(d.paragraphs[0].runs[1].italic)
        self.assertEqual(d.sections[0].left_margin.inches,.45)
    def test_revert_text_restores_exact_original(self):
        raw=pack(source());m=PreservedDocx(raw);key=ident(m,'First text');m.set_text(key,'Edit');m.set_text(key,'First text');self.assertEqual(m.bytes(),raw)
    def test_snapshot_undo_restores_exact_package(self):
        raw=pack(source());m=PreservedDocx(raw);state=m.snapshot();m.set_text(ident(m,'First text'),'Edit');m=m.append_paragraph();self.assertEqual(PreservedDocx.from_snapshot(state).bytes(),raw)
    def test_protected_control_cannot_be_edited(self):
        d=source();p=d.paragraphs[0]._p;sdt=OxmlElement('w:sdt');content=OxmlElement('w:sdtContent');p.getparent().replace(p,sdt);sdt.append(content);content.append(p)
        raw=pack(d);m=PreservedDocx(raw)
        with self.assertRaises(ValueError):m.set_text(ident(m,'First text'),'bad')
        self.assertEqual(m.bytes(),raw)
    def test_field_result_cannot_be_edited(self):
        d=Document();p=d.add_paragraph();r=p.add_run();f=OxmlElement('w:fldChar');f.set(qn('w:fldCharType'),'begin');r._r.append(f)
        p.add_run('Field result');r=p.add_run();f=OxmlElement('w:fldChar');f.set(qn('w:fldCharType'),'end');r._r.append(f)
        m=PreservedDocx(pack(d))
        with self.assertRaises(ValueError):m.set_text(ident(m,'Field result'),'bad')
    def test_invalid_input_does_not_change_package(self):
        raw=pack(source());m=PreservedDocx(raw);key=ident(m,'First text')
        for text in ['bad\x00','new\nline','tab\ttext','\ud800']:
            with self.assertRaises(ValueError):m.set_text(key,text)
        self.assertEqual(m.bytes(),raw)
    def test_append_preserves_original_xml_and_parts(self):
        raw=pack(source());m=PreservedDocx(raw);new=m.append_paragraph(ident(m,'First text'));a,b=parts(raw),parts(new.bytes())
        self.assertEqual([k for k in a if a[k]!=b[k]],['word/document.xml'])
        old=etree.fromstring(a['word/document.xml']);now=etree.fromstring(b['word/document.xml']);body=now.find(qn('w:body'));body.remove(body[1]);self.assertEqual(etree.tostring(old),etree.tostring(now))
    def test_empty_new_document_is_editable(self):
        m=PreservedDocx.blank();key=next(iter(m.texts));m.set_text(key,'New Word document');self.assertIn('New Word document',Document(io.BytesIO(m.bytes())).paragraphs[0].text)
    def test_failed_replace_preserves_target(self):
        m=PreservedDocx(pack(source()))
        with tempfile.TemporaryDirectory(dir='.tmp') as d:
            path=Path(d)/'saved.docx';path.write_bytes(m.original)
            m.set_text(ident(m,'First text'),'Edit')
            with patch('preserved_docx.os.replace',side_effect=PermissionError('locked')):
                with self.assertRaises(PermissionError):m.save(path)
            self.assertEqual(path.read_bytes(),m.original);self.assertEqual(len(list(Path(d).iterdir())),1)
    def test_unknown_zip_part_retained(self):
        raw=pack(source());b=io.BytesIO(raw)
        with zipfile.ZipFile(b,'a') as z:z.writestr('customXml/unmodeled.bin',b'opaque\x00\xff')
        m=PreservedDocx(b.getvalue());m.set_text(ident(m,'First text'),'Edit');self.assertEqual(parts(m.bytes())['customXml/unmodeled.bin'],b'opaque\x00\xff')
    def test_signed_package_readonly(self):
        b=io.BytesIO(pack(source()))
        with zipfile.ZipFile(b,'a') as z:z.writestr('_xmlsignatures/sig1.xml','signature')
        m=PreservedDocx(b.getvalue());self.assertTrue(m.readonly)
        with self.assertRaises(ValueError):m.set_text(ident(m,'First text'),'Edit')
    def test_original_fractional_size_never_normalized(self):
        d=source();d.paragraphs[0].runs[0]._r.rPr.sz.set(qn('w:val'),'16.5');raw=pack(d);m=PreservedDocx(raw);m.set_text(ident(m,'First text'),'New');self.assertIn(b'w:val="16.5"',parts(m.bytes())['word/document.xml'])

if __name__=='__main__':unittest.main(verbosity=2)

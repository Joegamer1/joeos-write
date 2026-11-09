import io, unittest, zipfile
from docx import Document
from lxml import etree
from preserved_docx import PreservedDocx, W
from text_editing import replace_range, format_range

def make():
    d=Document();p=d.add_paragraph();p.add_run('Alpha ').bold=True;p.add_run('Beta').italic=True
    d.add_paragraph('Gamma');out=io.BytesIO();d.save(out);return PreservedDocx(out.getvalue())

def point(identity,offset):return {'id':identity,'offset':offset}
def xml(model):return etree.fromstring(model.parts['word/document.xml'])
NS={'w':W}

class RangeTests(unittest.TestCase):
    def test_paragraph_commands_preserve_content(self):
        from text_editing import paragraph_format
        original=make();a=point('t0',0);b=point('t2',0)
        result=paragraph_format(original,a,b,'indent')
        root=xml(result)
        self.assertEqual(root.xpath('//w:body/w:p[1]/w:pPr/w:ind/@w:left',namespaces=NS),['720'])
        self.assertEqual(root.xpath('//w:body/w:p[2]/w:pPr/w:ind',namespaces=NS),[])
        result=paragraph_format(result,a,a,'double')
        result=paragraph_format(result,a,a,'center')
        self.assertEqual(''.join(xml(result).itertext()),''.join(xml(original).itertext()))
        for part in original.parts:
            if part!='word/document.xml':self.assertEqual(original.parts[part],result.parts[part])
    def test_empty_paragraph_alignment(self):
        from text_editing import paragraph_format,empty_paragraphs
        d=Document();d.add_paragraph();out=io.BytesIO();d.save(out);original=PreservedDocx(out.getvalue());entry=empty_paragraphs(original)[0];a=point(entry['id'],0)
        result=paragraph_format(original,a,a,'right')
        self.assertEqual(xml(result).xpath('//w:pPr/w:jc/@w:val',namespaces=NS),['right'])
    def test_split_keeps_runs_and_other_parts(self):
        original=make();result,caret=replace_range(original,point('t0',3),point('t0',3),'\n')
        ps=xml(result).xpath('//w:body/w:p',namespaces=NS)
        self.assertEqual([''.join(p.itertext()) for p in ps],['Alp','ha Beta','Gamma'])
        self.assertEqual(len(ps[1].xpath('.//w:i',namespaces=NS)),1)
        self.assertEqual(caret['offset'],0)
        for part in original.parts:
            if part!='word/document.xml':self.assertEqual(original.parts[part],result.parts[part])
    def test_join_and_cross_run_selection(self):
        result,_=replace_range(make(),point('t0',2),point('t2',2),'X')
        self.assertEqual(''.join(xml(result).itertext()),'AlXmma')
    def test_controls_and_unicode(self):
        original=make();result,caret=replace_range(original,point('t0',0),point('t0',0),'A\tB\vC\fD\n\U0001f600')
        root=xml(result)
        self.assertEqual(len(root.xpath('//w:tab',namespaces=NS)),1)
        self.assertEqual(len(root.xpath('//w:br',namespaces=NS)),2)
        self.assertEqual(caret['offset'],2)
        self.assertTrue(PreservedDocx(result.bytes()))
    def test_partial_format(self):
        result=format_range(make(),point('t0',2),point('t1',2),{'underline':True})
        self.assertEqual(''.join(xml(result).itertext()),'Alpha BetaGamma')
        runs=xml(result).xpath('//w:r[w:rPr/w:u]',namespaces=NS)
        self.assertEqual(''.join(''.join(r.itertext()) for r in runs),'pha Be')
    def test_insert_style_does_not_change_neighbours(self):
        result,_=replace_range(make(),point('t0',2),point('t0',2),'x',{'bold':False})
        off=xml(result).xpath('//w:r[w:rPr/w:b[@w:val="0"]]',namespaces=NS)
        self.assertEqual(''.join(''.join(r.itertext()) for r in off),'x')
    def test_reject_invalid_offsets_without_mutation(self):
        original=make();raw=original.bytes()
        with self.assertRaises(ValueError):replace_range(original,point('t0',100),point('t1',0),'')
        self.assertEqual(original.bytes(),raw)

if __name__=='__main__':unittest.main()

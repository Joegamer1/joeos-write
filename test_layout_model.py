import io,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
from docx import Document
from docx.shared import Inches,Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from layout_metrics import analyze_layout
from preview_profiles import PreviewProfiles

def packed(document):
    b=io.BytesIO();document.save(b);return b.getvalue()

class LayoutModelTests(unittest.TestCase):
    def test_exact_geometry_and_style_inheritance(self):
        d=Document();s=d.sections[0];s.page_width=Inches(8.5);s.page_height=Inches(11);s.left_margin=Inches(.68);s.right_margin=Inches(.68)
        base=d.styles.add_style('ParentLayout',WD_STYLE_TYPE.PARAGRAPH);base.paragraph_format.space_before=Pt(6);base.paragraph_format.line_spacing=1.25
        child=d.styles.add_style('ChildLayout',WD_STYLE_TYPE.PARAGRAPH);child.base_style=base;child.paragraph_format.left_indent=Inches(.5)
        p=d.add_paragraph('Independent measurements',style=child);p.paragraph_format.space_after=Pt(3)
        r=analyze_layout(packed(d));section=r['sections']['s0'];metrics=r['paragraphs']['p0']
        self.assertEqual(section['widthPt'],612);self.assertEqual(section['heightPt'],792)
        self.assertEqual(section['marginsPt']['left'],48.95)
        self.assertEqual((metrics['beforePt'],metrics['afterPt'],metrics['lineRule'],metrics['line']),(6,3,'auto',1.25))
        self.assertEqual(metrics['indentsPt']['left'],36)
    def test_exact_line_rule_and_hanging_indent(self):
        d=Document();p=d.add_paragraph('Exact');p.paragraph_format.line_spacing=Pt(14);p.paragraph_format.first_line_indent=Inches(-.25)
        r=analyze_layout(packed(d))['paragraphs']['p0']
        self.assertEqual((r['lineRule'],r['line']),('exact',14));self.assertEqual(r['indentsPt']['hanging'],18)
    def test_style_cycle_is_bounded(self):
        d=Document();a=d.styles.add_style('CycleA',WD_STYLE_TYPE.PARAGRAPH);b=d.styles.add_style('CycleB',WD_STYLE_TYPE.PARAGRAPH);a.base_style=b;b.base_style=a
        d.add_paragraph('Cycle',style=a);self.assertIn('p0',analyze_layout(packed(d))['paragraphs'])
    def test_preview_preferences_are_independent_and_atomic(self):
        with tempfile.TemporaryDirectory(dir='.tmp') as folder:
            store=PreviewProfiles(Path(folder)/'preferences.json');profile={'fontSubstitutions':{'Aptos':'Noto Serif'},'pdfSpacing':True}
            store.put('original-hash',profile);self.assertEqual(store.get('original-hash'),profile);self.assertEqual(store.get('different-document'),{})
            before=store.path.read_bytes()
            with patch('preview_profiles.os.replace',side_effect=PermissionError):
                with self.assertRaises(PermissionError):store.put('new-hash',profile)
            self.assertEqual(store.path.read_bytes(),before)
            store.put('original-hash',{});self.assertEqual(store.get('original-hash'),{})

if __name__=='__main__':unittest.main()

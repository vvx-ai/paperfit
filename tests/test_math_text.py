import sys,tempfile,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from paperfit import *
from docx import Document
from docx.shared import Pt
class MathText(unittest.TestCase):
 def test_body_outside_math_only(self):
  with tempfile.TemporaryDirectory() as temp:
   t=Path(temp);a=t/'a.docx';b=t/'b.docx';c=t/'c.docx'
   d=Document();p=d.add_paragraph();run=p.add_run('公式前正文');run.font.name='仿宋';run.font.size=Pt(9)
   math=E.fromstring(f'<m:oMath xmlns:m="{M}" xmlns:w="{W}"><m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/><w:sz w:val="30"/></w:rPr><m:t>x+y</m:t></m:r></m:oMath>')
   p._p.append(math);p.add_run('公式后正文').font.size=Pt(18)
   # Standalone math must remain untouched.
   p=d.add_paragraph();p._p.append(copy.deepcopy(math));d.save(a)
   rules=parse('正文宋体小四单倍行距首行缩进2字符')
   report=format_docx(a,b,rules);self.assertTrue(report['math_xml_unchanged']);format_docx(b,c,rules)
   x=Document(b);p=x.paragraphs[0];self.assertEqual(p.style.name,'Body Text');self.assertEqual(p.style.font.size.pt,12);self.assertEqual(p.style.font.name,'宋体');self.assertEqual(p.style.paragraph_format.line_spacing,1)
   self.assertTrue(all(r.font.size is None and r.font.name is None for r in p.runs))
   with zipfile.ZipFile(a) as za,zipfile.ZipFile(b) as zb,zipfile.ZipFile(c) as zc:
    ra=E.fromstring(za.read('word/document.xml'));rb=E.fromstring(zb.read('word/document.xml'))
    self.assertEqual(math_snapshot(ra),math_snapshot(rb));self.assertEqual(zb.read('word/document.xml'),zc.read('word/document.xml'))
    self.assertEqual(fingerprint(ra),fingerprint(rb))

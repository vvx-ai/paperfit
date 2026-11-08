import unittest,tempfile,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from paperfit import *
from docx import Document
class Scope(unittest.TestCase):
 def test_preserved_inherited_style_rejected(self):
  with tempfile.TemporaryDirectory() as t:
   a=Path(t)/'a.docx';b=Path(t)/'b.docx';d=Document()
   d.add_paragraph('正文',style='Body Text')
   s=d.styles.add_style('Special',1);s.base_style=d.styles['Body Text'];d.add_paragraph('不能改变',style=s);d.save(a)
   r=parse('正文宋体小四');r['paragraph_roles']={'1':'preserve'}
   with self.assertRaisesRegex(ValueError,'共享样式'):format_docx(a,b,r)
   self.assertFalse(b.exists())
 def test_header_style_rejected(self):
  with tempfile.TemporaryDirectory() as t:
   a=Path(t)/'a.docx';b=Path(t)/'b.docx';d=Document();d.add_paragraph('正文');d.sections[0].header.paragraphs[0].style=d.styles['Body Text'];d.save(a)
   with self.assertRaisesRegex(ValueError,'共享样式'):format_docx(a,b,parse('正文宋体小四'))
   self.assertFalse(b.exists())

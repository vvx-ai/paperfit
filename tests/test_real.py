import unittest, tempfile, sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from paperfit import *
class Real(unittest.TestCase):
 def test_native_style(self):
  from docx import Document
  from docx.shared import Pt
  with tempfile.TemporaryDirectory() as t:
   a=Path(t)/'a.docx';b=Path(t)/'b.docx';c=Path(t)/'c.docx'
   d=Document();p=d.add_paragraph('测试标题',style='Heading 1');p.add_run('内容').font.size=Pt(30)
   d.add_paragraph('正文');d.add_table(rows=1,cols=1).cell(0,0).text='表格内容'
   d.styles['Heading 1'].element.set(q('styleId'),'1')
   p._p.pPr.pStyle.val='1'
   d.save(a);r=parse('一级标题：宋体小三加粗单倍行距；正文：宋体小四，单倍行距，首行缩进2字符')
   self.assertFalse(r['pending']);format_docx(a,b,r);format_docx(b,c,r)
   x=Document(b);self.assertEqual(x.paragraphs[0].style.style_id,'1');self.assertEqual(len([s for s in x.styles if s.name=='Heading 1']),1); self.assertEqual(x.paragraphs[0].style.name,'Heading 1')
   self.assertEqual(x.paragraphs[0].style.font.size.pt,15);self.assertTrue(x.paragraphs[0].style.font.bold)
   self.assertIsNone(x.paragraphs[0].runs[-1].font.size)
   with zipfile.ZipFile(b) as z,zipfile.ZipFile(c) as zz:self.assertEqual(z.read('word/document.xml'),zz.read('word/document.xml'))

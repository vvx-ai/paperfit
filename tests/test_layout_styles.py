import unittest,tempfile,sys,base64
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from isolate_layout_styles import isolate,q
from docx import Document
from docx.shared import Pt
class LayoutStyles(unittest.TestCase):
 def test_title_and_image_independence(self):
  with tempfile.TemporaryDirectory() as tmp:
   t=Path(tmp);pic=t/'a.png';pic.write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII='))
   d=Document();p=d.add_paragraph('总标题');p.runs[0].font.name='宋体';p.runs[0].font.size=Pt(16);d.add_picture(str(pic));d.add_paragraph('正文')
   a=t/'a.docx';b=t/'b.docx';c=t/'c.docx';d.save(a);isolate(a,b,0);isolate(b,c,0)
   x=Document(b);self.assertEqual(x.paragraphs[0].style.name,'论文总标题');self.assertEqual(x.paragraphs[1].style.name,'图片段落');self.assertIsNone(x.paragraphs[0].style.base_style);self.assertIsNone(x.paragraphs[1].style.base_style)
   self.assertIsNone(x.paragraphs[0].runs[0].font.size);self.assertEqual(x.paragraphs[0].style.font.size.pt,16)
   x.paragraphs[0].style.font.size=Pt(22);self.assertNotEqual(x.paragraphs[1].style.font.size,Pt(22))
   with ZipFile(a) as za,ZipFile(b) as zb,ZipFile(c) as zc:
    self.assertEqual(zb.read('word/document.xml'),zc.read('word/document.xml'))
    self.assertEqual(zb.read('word/styles.xml'),zc.read('word/styles.xml'))
    for n in za.namelist():
     if n not in ('word/document.xml','word/styles.xml'):self.assertEqual(za.read(n),zb.read(n))

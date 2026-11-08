import unittest,sys,tempfile,base64
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from paperfit import *
from docx import Document
from docx.shared import Inches
class Objects(unittest.TestCase):
 def test_objects_and_chinese_styles(self):
  with tempfile.TemporaryDirectory() as t:
   t=Path(t);pic=t/'pic.png'
   pic.write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII='))
   d=Document();s=d.styles.add_style('一级标题',1);p=d.add_paragraph('第一章 测试',s);d.add_paragraph('正文');d.add_picture(str(pic),width=Inches(1));d.add_table(rows=1,cols=1).cell(0,0).text='表格内容'
   a=t/'a.docx';b=t/'b.docx';c=t/'c.docx';d.save(a)
   r=parse('一级标题宋体小三加粗单倍行距；正文宋体小四单倍行距首行缩进2字符')
   format_docx(a,b,r);format_docx(b,c,r)
   x=Document(b);self.assertEqual(x.paragraphs[0].style.name,'一级标题');self.assertEqual(x.paragraphs[1].style.name,'Body Text');self.assertIsNone(x.paragraphs[1].runs[0].font.size)
   self.assertEqual(x.tables[0].cell(0,0).text,'表格内容');self.assertEqual(len(x.inline_shapes),1)
   with zipfile.ZipFile(a) as z,zipfile.ZipFile(b) as zz,zipfile.ZipFile(c) as zzz:
    for n in z.namelist():
     if n not in ('word/document.xml','word/styles.xml'):self.assertEqual(z.read(n),zz.read(n))
    for n in ('word/document.xml','word/styles.xml'):self.assertEqual(zz.read(n),zzz.read(n))

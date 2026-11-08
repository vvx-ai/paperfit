import sys, tempfile, unittest, zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from paperfit import *
class EngineTest(unittest.TestCase):
 def test_roundtrip(self):
  with tempfile.TemporaryDirectory() as d:
   src=Path(d)/'sample.docx'; out=Path(d)/'out.docx'; again=Path(d)/'again.docx'
   xml=f'<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章 绪论</w:t></w:r></w:p><w:p><w:r><w:t>正文内容</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格内容</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgMar w:top="1000"/></w:sectPr></w:body></w:document>'
   with zipfile.ZipFile(src,'w') as z:
    z.writestr('word/document.xml',xml); z.writestr('word/styles.xml',f'<w:styles xmlns:w="{W}"/>'); z.writestr('word/media/image.png',b'image-preservation-fixture')
   original=src.read_bytes(); rules=parse('正文宋体小四，1.5倍行距，首行缩进2字符；一级标题黑体三号；上页边距2.5厘米')
   self.assertFalse(rules['pending']); self.assertTrue(format_docx(src,out,rules)['integrity']); format_docx(out,again,rules)
   self.assertEqual(original,src.read_bytes())
   with zipfile.ZipFile(out) as z,zipfile.ZipFile(again) as zz:
    self.assertEqual(z.read('word/document.xml'),zz.read('word/document.xml')); self.assertEqual(z.read('word/media/image.png'),b'image-preservation-fixture')
    root=E.fromstring(z.read('word/document.xml')); self.assertEqual(root.find('.//'+q('pStyle')).get(q('val')),'Heading1')
    styles=E.fromstring(z.read('word/styles.xml')); self.assertEqual(styles.find('.//'+q('sz')).get(q('val')),'24')
   with self.assertRaises(ValueError): format_docx(src,src,rules)
 def test_ambiguity(self):
  self.assertTrue(parse('正文宋体黑体')['pending']); self.assertTrue(parse('按学校规定')['pending'])
if __name__=='__main__': unittest.main()

import unittest,tempfile,sys,zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from number_captions import number_captions,E,NS
from docx import Document
class NumberCaptions(unittest.TestCase):
 def test_swapped_numbers_fields_and_repeat(self):
  with tempfile.TemporaryDirectory() as tmp:
   t=Path(tmp);a=t/'a.docx';b=t/'b.docx';c=t/'c.docx';d=Document()
   d.add_table(rows=1,cols=1).cell(0,0).text='数据甲';cap=d.add_paragraph('表2 标题甲');cap.runs[0]._r.insert(0,E.Element('{'+NS['w']+'}lastRenderedPageBreak'))
   d.add_paragraph('正文分隔');d.add_table(rows=1,cols=1).cell(0,0).text='数据乙';d.add_paragraph('表1 标题乙');d.save(a)
   number_captions(a,b);number_captions(b,c)
   with zipfile.ZipFile(b) as x,zipfile.ZipFile(c) as y:
    self.assertEqual(x.read('word/document.xml'),y.read('word/document.xml'))
    root=E.fromstring(x.read('word/document.xml'))
    self.assertEqual(root.xpath('//w:fldSimple/@w:instr',namespaces=NS),[' SEQ 表 \\* ARABIC ']*2)
    settings=E.fromstring(x.read('word/settings.xml'));self.assertEqual(set(settings.xpath('w:captions/w:caption/@w:name',namespaces=NS)),{'图','表'})
    self.assertEqual(root.xpath('//w:fldSimple/w:r/w:t/text()',namespaces=NS),['1','2'])
   d.add_paragraph('如表2所示');d.save(a)
   with self.assertRaisesRegex(ValueError,'引用'):number_captions(a,t/'blocked.docx')
   self.assertFalse((t/'blocked.docx').exists())

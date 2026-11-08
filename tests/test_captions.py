import sys,tempfile,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from match_captions import match_captions
from docx import Document
class Captions(unittest.TestCase):
 def test_adjacent_and_ambiguous(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'a.docx';d=Document();d.add_table(rows=1,cols=1);d.add_paragraph('表2');d.add_paragraph('隔开的正文');d.add_paragraph('表3');d.save(p);raw=p.read_bytes();r=match_captions(p)
   self.assertEqual(r['matches'][0]['status'],'unique_adjacent_match');self.assertEqual(r['matches'][1]['status'],'needs_review');self.assertEqual(raw,p.read_bytes())
   d=Document();d.add_table(rows=1,cols=1);d.add_paragraph('表1');d.add_table(rows=1,cols=1);d.save(p)
   self.assertEqual(match_captions(p)['matches'][0]['status'],'needs_review')

import sys,tempfile,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from analyze_docx import analyze
from docx import Document
class Analysis(unittest.TestCase):
 def test_candidates_and_readonly(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'x.docx';d=Document()
   for text in ['1.第一节','1.第二节','图2-1 示例','正文引用[1]']:d.add_paragraph(text)
   d.save(p);raw=p.read_bytes();r=analyze(p)
   self.assertEqual(p.read_bytes(),raw);self.assertEqual(r['warnings'][0]['type'],'duplicate_heading_number');self.assertEqual(r['paragraphs'][2]['caption_candidate']['number'],'2-1');self.assertEqual(r['paragraphs'][3]['citation_candidates'],['[1]'])

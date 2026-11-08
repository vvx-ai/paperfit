import unittest,tempfile,sys
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from normalize_captions import normalize,q,NS
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
class Normalize(unittest.TestCase):
 def test_caption_overrides_and_fields(self):
  with tempfile.TemporaryDirectory() as tmp:
   t=Path(tmp);a=t/'a.docx';b=t/'b.docx';c=t/'c.docx';d=Document();d.add_paragraph('总标题')
   s=d.styles['Caption'];s.base_style=None
   for label,num in [('图',5),('表',3)]:
    p=d.add_paragraph(label+' ',style=s)
    if label=='图':p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    for kind,text in [('begin',None),(None,' SEQ '+label+' '+chr(92)+'* ARABIC '),('separate',None),(None,str(num)),('end',None)]:
     r=E.SubElement(p._p,q('r'))
     if kind:E.SubElement(r,q('fldChar'),{q('fldCharType'):kind})
     else:
      if 'SEQ' not in text:E.SubElement(E.SubElement(r,q('rPr')),q('noProof'))
      E.SubElement(r,q('instrText') if 'SEQ' in text else q('t')).text=text
   d.save(a);normalize(a,b);normalize(b,c)
   with ZipFile(a) as za,ZipFile(b) as zb,ZipFile(c) as zc:
    x=E.fromstring(za.read('word/document.xml'));y=E.fromstring(zb.read('word/document.xml'))
    self.assertEqual(x.xpath('//w:instrText/text()',namespaces=NS),y.xpath('//w:instrText/text()',namespaces=NS))
    self.assertFalse(y.xpath('//w:p[w:pPr/w:pStyle]/w:r/w:rPr/w:noProof',namespaces=NS))
    ps=y.findall('./'+q('body')+'/'+q('p'))
    self.assertEqual(E.tostring(ps[1].find(q('pPr'))),E.tostring(ps[2].find(q('pPr'))))
    for n in zb.namelist():self.assertEqual(zb.read(n),zc.read(n))

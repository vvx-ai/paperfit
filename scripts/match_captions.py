"""Read-only immediate-neighbor caption matching; never guesses across text."""
import argparse,json,re,zipfile,hashlib
from pathlib import Path
from lxml import etree as E
from word_fields import scan
from paperfit import q
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';NS={'w':W}
def match_captions(source):
 with zipfile.ZipFile(source) as z:root=E.fromstring(z.read('word/document.xml'))
 nodes=list(root.find('w:body',NS));objects={};captions=[]
 for i,e in enumerate(nodes):
  if e.tag=='{'+W+'}tbl':
   marker=e.find('w:tblPr/w:tblCaption',NS)
   if marker is None or marker.get('{'+W+'}val')!='PaperFit:equation:v1':objects[i]={'kind':'表','count':1}
  elif e.tag=='{'+W+'}p':
   count=len(e.findall('.//w:drawing',NS))
   if count:objects[i]={'kind':'图','count':count}
   text=''.join(e.xpath('.//w:t/text()',namespaces=NS))
   m=re.fullmatch(r'\s*([图表])\s*(\d+(?:[-－.]\d+)*)(.*)',text)
   if m:
    fields,_ = scan(e)
    ref_text = {n for f in fields if f.code.split() and f.code.split()[0].upper() in ('REF','PAGEREF','NOTEREF') for n in f.result}
    offset = 0
    for n in e.iter(q('t')):
     end = offset+len(n.text or '')
     if n in ref_text and offset < m.end(2) and end > m.start(2):
      m=None
      break
     offset=end
   if m:captions.append({'caption_body_index':i,'text':text,'kind':m[1],'number':m[2]})
 for cap in captions:
  i=cap['caption_body_index'];candidates=[j for j in (i-1,i+1) if j in objects and objects[j]['kind']==cap['kind']]
  cap['candidates']=candidates;cap['status']='needs_review'
  if len(candidates)==1 and objects[candidates[0]]['count']==1:
   cap['object_body_index']=candidates[0];cap['position']='after' if candidates[0]<i else 'before';cap['status']='unique_adjacent_match'
 for cap in captions:
  obj=cap.get('object_body_index')
  if obj is not None and sum(c.get('object_body_index')==obj for c in captions)>1:cap['status']='needs_review'
 matched={c['object_body_index'] for c in captions if c['status']=='unique_adjacent_match'}
 return {'source':str(Path(source).resolve()),'sha256':hashlib.sha256(Path(source).read_bytes()).hexdigest(),'matches':captions,'unmatched_objects':[i for i in objects if i not in matched],'scope':'正文直属对象的相邻关系候选，不代表浮动对象的视觉位置；不修改编号、题注或引用。'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');a=p.parse_args()
 with Path(a.output).open('x',encoding='utf-8') as f:json.dump(match_captions(a.source),f,ensure_ascii=False,indent=2)

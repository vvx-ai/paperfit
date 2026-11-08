"""Read-only OOXML inventory. Matches are candidates, never editing instructions."""
import argparse,json,re,zipfile,hashlib
from pathlib import Path
from lxml import etree as E
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}
def analyze(path):
 data=Path(path).read_bytes()
 with zipfile.ZipFile(path) as z:
  root=E.fromstring(z.read('word/document.xml'));styles=E.fromstring(z.read('word/styles.xml'))
  names={s.get('{'+NS['w']+'}styleId'):s.find('w:name',NS).get('{'+NS['w']+'}val') for s in styles if s.find('w:name',NS) is not None}
  media=[n for n in z.namelist() if n.startswith('word/media/')]
 result={'source':str(Path(path).resolve()),'sha256':hashlib.sha256(data).hexdigest(),'paragraphs':[],'tables':len(root.findall('.//w:tbl',NS)),'drawing_objects':len(root.findall('.//w:drawing',NS)),'media_parts':media,'equations':len(root.findall('.//m:oMath',NS)),'fields':root.xpath('//w:instrText/text()|//w:fldSimple/@w:instr',namespaces=NS),'warnings':[]}
 numbers={}
 for i,p in enumerate(root.findall('w:body/w:p',NS)):
  text=''.join(p.xpath('.//w:t/text()',namespaces=NS));style=p.find('w:pPr/w:pStyle',NS);sid=style.get('{'+NS['w']+'}val') if style is not None else None
  item={'index':i,'text':text,'style_id':sid,'style_name':names.get(sid),'heading_candidate':None,'caption_candidate':None,'citation_candidates':re.findall(r'\[\d+(?:\s*[-,，、]\s*\d+)*\]',text)}
  match=re.match(r'^\s*(\d+(?:\.\d+){0,2})[.．、\s]?([^\d].*)$',text)
  if match and len(text.strip())<70:
   num=match[1];item['heading_candidate']={'level':num.count('.')+1,'number':num,'confidence':'requires_review'}
   if num in numbers:result['warnings'].append({'type':'duplicate_heading_number','number':num,'paragraphs':[numbers[num],i]})
   numbers[num]=i
  cap=re.match(r'^\s*([图表])\s*(\d+(?:[-－.]\d+)?)\s*(.*)$',text)
  if cap:item['caption_candidate']={'kind':cap[1],'number':cap[2],'title':cap[3]}
  result['paragraphs'].append(item)
 result['warnings'].append({'type':'scope','message':'候选标题和引用需确认；正文直属段落以外的内容仅统计对象，未执行引用匹配或编号修改。'})
 return result
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');a=p.parse_args();out=Path(a.output)
 with out.open('x',encoding='utf-8') as f:json.dump(analyze(a.source),f,ensure_ascii=False,indent=2)

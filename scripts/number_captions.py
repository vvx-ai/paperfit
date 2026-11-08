"""Conservative global SEQ numbering for uniquely adjacent figure/table captions."""
import argparse,copy,json,re,zipfile
from pathlib import Path
from lxml import etree as E
from match_captions import match_captions
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';NS={'w':W}
def q(n):return '{'+W+'}'+n
IDS={'图':'图','表':'表'}
from caption_style import configure
def insert_seq(p,start,end,new,instruction):
 offset=0;target=None
 for t in p.iter(q('t')):
  length=len(t.text or '')
  if offset<=start and end<=offset+length:target=t;break
  offset+=length
 if target is None:raise ValueError('编号跨文本节点，需人工确认')
 run=target.getparent()
 if run.tag!=q('r') or run.getparent() is not p or any(x.tag not in (q('rPr'),q('t'),q('lastRenderedPageBreak')) for x in run) or len(run.findall(q('t')))!=1:raise ValueError('题注编号结构不支持')
 value=target.text;left=value[:start-offset];right=value[end-offset:];position=p.index(run);rp=run.find(q('rPr'))
 def make_run(value):
  r=E.Element(q('r'),dict(run.attrib))
  if rp is not None:r.append(copy.deepcopy(rp))
  t=E.SubElement(r,q('t'));t.text=value
  if value.startswith(' ') or value.endswith(' '):t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
  return r
 replacements=[]
 prefix=[copy.deepcopy(x) for x in list(run)[:run.index(target)] if x.tag==q('lastRenderedPageBreak')]
 suffix=[copy.deepcopy(x) for x in list(run)[run.index(target)+1:] if x.tag==q('lastRenderedPageBreak')]
 if left or prefix:
  lr=make_run(left)
  for marker in prefix:lr.insert(len(lr)-1,marker)
  replacements.append(lr)
 field=E.Element(q('fldSimple'),{q('instr'):instruction});field.append(make_run(new));replacements.append(field)
 if right or suffix:
  rr=make_run(right)
  for marker in suffix:rr.append(marker)
  replacements.append(rr)
 tail=run.tail;p.remove(run)
 for j,x in enumerate(replacements):p.insert(position+j,x)
 replacements[-1].tail=tail
 return field

def number_captions(source,output):
 source,output=Path(source),Path(output)
 if source.resolve()==output.resolve() or output.exists():raise ValueError('输出必须是新的文件')
 analysis=match_captions(source)
 if analysis['unmatched_objects'] or any(c['status']!='unique_adjacent_match' for c in analysis['matches']):raise ValueError('对象或题注匹配不唯一，停止编号')
 with zipfile.ZipFile(source) as z:parts=[(i,z.read(i.filename)) for i in z.infolist()]
 raw={i.filename:b for i,b in parts};root=E.fromstring(raw['word/document.xml']);nodes=list(root.find(q('body')))
 before=[E.tostring(x,method='c14n') for x in nodes];indices={c['caption_body_index'] for c in analysis['matches']}
 # No guessed reference rewriting. Reject textual references or fields outside captions.
 for name,data in raw.items():
  if not name.startswith('word/') or not name.endswith('.xml') or name=='word/styles.xml':continue
  tree=root if name=='word/document.xml' else E.fromstring(data)
  for p in tree.iter(q('p')):
   if name=='word/document.xml' and p in [nodes[i] for i in indices]:continue
   text=''.join(p.xpath('.//w:t/text()',namespaces=NS))
   if re.search(r'[图表]\s*\d',text) or p.find('.//'+q('instrText')) is not None or p.find('.//'+q('fldSimple')) is not None:
    raise ValueError('发现题注外引用或域，需先确认引用映射，未修改文件')
 counts={'图':0,'表':0};changes=[]
 for cap in analysis['matches']:
  index=cap['caption_body_index'];p=nodes[index];kind=cap['kind'];counts[kind]+=1;new=str(counts[kind])
  text=''.join(p.xpath('.//w:t/text()',namespaces=NS));m=re.match(r'^\s*[图表]\s*(\d+(?:[-－.]\d+)*)',text)
  if m is None:raise ValueError('无法定位题注编号')
  start,end=m.span(1);expected=text[:start]+new+text[end:];instruction=' SEQ '+IDS[kind]+' \\* ARABIC '
  fields=p.findall('.//'+q('fldSimple'))
  if fields:
   if len(fields)!=1 or fields[0].get(q('instr'),'').strip() not in (instruction.strip(),('SEQ '+{'图':'PaperFitFigure','表':'PaperFitTable'}[kind]+' \\* ARABIC')) or p.find('.//'+q('instrText')) is not None:raise ValueError('题注含不支持的已有域')
   ts=fields[0].findall('.//'+q('t'))
   if len(ts)!=1 or ts[0].text!=m[1]:raise ValueError('题注域内容异常')
   ts[0].text=new
   fields[0].set(q('instr'),instruction)
  else:
   if p.find('.//'+q('instrText')) is not None:raise ValueError('题注含复杂域')
   insert_seq(p,start,end,new,instruction)
  if ''.join(p.xpath('.//w:t/text()',namespaces=NS))!=expected:raise ValueError('题注内容校验失败')
  changes.append({'caption_index':index,'old':text,'new':expected,'field':instruction.strip()})
 for i,node in enumerate(nodes):
  if i not in indices and E.tostring(node,method='c14n')!=before[i]:raise ValueError('非题注内容被改变')
 styles=E.fromstring(raw['word/styles.xml']);settings=E.fromstring(raw['word/settings.xml'])
 caption_style_id=configure(root,styles,settings,sorted(indices))
 replacements={name:E.tostring(tree,encoding='utf-8',xml_declaration=True) for name,tree in [('word/document.xml',root),('word/styles.xml',styles),('word/settings.xml',settings)]}
 with output.open('xb') as f:
  with zipfile.ZipFile(f,'w') as z:
   for info,data in parts:z.writestr(info,replacements.get(info.filename,data))
 return {'caption_style_id':caption_style_id,'labels':['图','表'],'source_sha256':analysis['sha256'],'numbering':'全文连续，图表各自从1开始','changes':changes,'checks':['非题注节点XML完全不变','题注除编号外文字完全不变','除document/styles/settings外ZIP部件原样复制'],'limitations':['不自动更新已有正文引用；检测到引用或未知域时停止','SEQ域按文档顺序更新，不自动跟随对象移动题注','跨文本节点编号暂不处理']}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');a=p.parse_args();report=Path(a.output+'.report.json')
 if report.exists():raise ValueError('报告已存在')
 result=number_captions(a.source,a.output)
 with report.open('x',encoding='utf-8') as f:json.dump(result,f,ensure_ascii=False,indent=2)

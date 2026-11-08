"""Normalize existing Word captions without rebuilding fields or bookmarks."""
import argparse,copy,json,zipfile
from pathlib import Path
from lxml import etree as E
from paperfit import q,W,fingerprint
from word_fields import scan
NS={'w':W}
VISUAL={'rFonts','sz','szCs','b','bCs','i','iCs','color','spacing','kern','w','position','highlight','shd','u','strike','dstrike','caps','smallCaps','emboss','imprint','outline','shadow','noProof'}
def normalize(source,output):
 source,output=Path(source),Path(output)
 if output.exists() or source.resolve()==output.resolve():raise ValueError('输出必须是新文件')
 with zipfile.ZipFile(source) as z:parts=[(i,z.read(i.filename)) for i in z.infolist()]
 raw={i.filename:b for i,b in parts};root=E.fromstring(raw['word/document.xml']);styles=E.fromstring(raw['word/styles.xml'])
 matches=[s for s in styles if s.get(q('type'))=='paragraph' and s.find(q('name')) is not None and s.find(q('name')).get(q('val'),'').lower() in ('caption','题注')]
 if len(matches)!=1:raise ValueError('需唯一既有题注样式')
 style=matches[0];sid=style.get(q('styleId'));paras=root.find(q('body')).findall(q('p'));selected=[]
 for p in paras:
  st=p.find('./'+q('pPr')+'/'+q('pStyle'))
  if st is not None and st.get(q('val'))==sid:selected.append(p)
 if not selected:raise ValueError('没有已应用题注样式的段落')
 # Guard against collateral changes in other stories or inherited styles.
 for s in styles:
  b=s.find(q('basedOn'))
  if b is not None and b.get(q('val'))==sid:raise ValueError('题注样式有继承者，需确认范围')
 for name,data in raw.items():
  if not name.startswith('word/') or not name.endswith('.xml') or name=='word/styles.xml':continue
  tree=root if name=='word/document.xml' else E.fromstring(data)
  for p in tree.iter(q('p')):
   st=p.find('./'+q('pPr')+'/'+q('pStyle'))
   if st is not None and st.get(q('val'))==sid and p not in selected:raise ValueError('题注样式用于其他区域，需确认范围')
 if style.find(q('basedOn')) is not None:raise ValueError('题注样式仍有继承，需先展开有效格式')
 snapshot=fingerprint(root);untouched={p:E.tostring(p,method='c14n') for p in paras if p not in selected}
 original_styles={s:E.tostring(s,method='c14n') for s in styles if s is not style}
 pp=style.find(q('pPr'))
 if pp is None:
  pp=E.Element(q('pPr'));rp=style.find(q('rPr'));style.insert(style.index(rp) if rp is not None else len(style),pp)
 jc=pp.find(q('jc'))
 if jc is None:
  jc=E.SubElement(pp,q('jc'));jc.set(q('val'),'left') # Original Caption has no jc and no basedOn/default jc.
 for tag in ('autoRedefine','link','hidden','semiHidden'):
  for x in list(style):
   if x.tag==q(tag):style.remove(x)
 changes=[]
 for p in selected:
  text=''.join(p.xpath('.//w:t/text()',namespaces=NS));ppr=p.find(q('pPr'))
  if p.find('.//'+q('drawing')) is not None or ppr.find(q('sectPr')) is not None or p.find('.//'+q('ins')) is not None or p.find('.//'+q('del')) is not None:raise ValueError('题注结构不支持安全统一')
  fields,errors=scan(p)
  if errors or len(fields)!=1 or not fields[0].code.startswith(('SEQ 图 ','SEQ 表 ')):raise ValueError('题注编号域异常')
  removed=[]
  for x in list(ppr):
   if x.tag!=q('pStyle'):removed.append(E.QName(x).localname);ppr.remove(x)
  for rp in p.iter(q('rPr')):
   if rp.find(q('rStyle')) is not None:raise ValueError('题注含字符样式，需确认后再统一')
   for x in list(rp):
    if E.QName(x).localname in VISUAL:removed.append(E.QName(x).localname);rp.remove(x)
   if not len(rp):rp.getparent().remove(rp)
  changes.append({'text':text,'removed_overrides':removed})
 assert fingerprint(root)==snapshot
 assert all(E.tostring(p,method='c14n')==v for p,v in untouched.items())
 assert all(E.tostring(s,method='c14n')==original_styles[s] for s in styles if s is not style)
 replacements={'word/document.xml':E.tostring(root,encoding='utf-8',xml_declaration=True),'word/styles.xml':E.tostring(styles,encoding='utf-8',xml_declaration=True)}
 with output.open('xb') as f:
  with zipfile.ZipFile(f,'w') as z:
   for info,data in parts:z.writestr(info,replacements.get(info.filename,data))
 return {'caption_style_id':sid,'count':len(selected),'alignment':jc.get(q('val')),'changes':changes,'checks':['域代码与编号文字保持不变','非题注段落XML不变','其他样式定义不变','其余ZIP部件逐字节保留'],'note':'以后手动插入仍可能带入直接格式；需重置段落格式或重新执行本命令。不能通过DOCX禁止Word新增直接格式。'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');a=p.parse_args();report=Path(a.output+'.report.json')
 if report.exists():raise ValueError('报告已存在')
 r=normalize(a.source,a.output)
 with report.open('x',encoding='utf-8') as f:json.dump(r,f,ensure_ascii=False,indent=2)

"""Separate an explicitly selected title and image-only paragraphs from Normal."""
import argparse,copy,json,zipfile
from pathlib import Path
from lxml import etree as E
from paperfit import q,W,fingerprint

def isolate(source,output,title_index):
 source,output=Path(source),Path(output)
 if output.exists() or output.resolve()==source.resolve():raise ValueError('必须输出到新文件')
 with zipfile.ZipFile(source) as z:parts=[(i,z.read(i.filename)) for i in z.infolist()]
 data={i.filename:b for i,b in parts};root=E.fromstring(data['word/document.xml']);styles=E.fromstring(data['word/styles.xml']);paras=root.find(q('body')).findall(q('p'))
 if not 0<=title_index<len(paras):raise ValueError('总标题索引无效')
 title=paras[title_index]
 if title.find('.//'+q('drawing')) is not None or not ''.join(x.text or '' for x in title.iter(q('t'))).strip():raise ValueError('总标题必须是文字段落')
 before=fingerprint(root);snap=[E.tostring(p,method='c14n') for p in paras]
 selected=[(title_index,'PaperFitDocumentTitle','论文总标题')]+[(i,'PaperFitImageParagraph','图片段落') for i,p in enumerate(paras) if p.find('.//'+q('drawing')) is not None and not ''.join(x.text or '' for x in p.iter(q('t'))).strip()]
 lookup={s.get(q('styleId')):s for s in styles if s.tag==q('style')};default=next((s for s in lookup.values() if s.get(q('default'))=='1' and s.get(q('type'))=='paragraph'),None)
 def merge(dest,src):
  if src is None:return
  for x in src:
   if x.tag in {q('pStyle'),q('rStyle'),q('rPr'),q('sectPr'),q('numPr')}:continue
   old=dest.find(x.tag)
   if old is None:dest.append(copy.deepcopy(x))
   else:
    old.attrib.update(x.attrib)
    if len(x):
     for c in list(old):old.remove(c)
     for c in x:old.append(copy.deepcopy(c))
 created=[]
 for i,sid,name in selected:
  p=paras[i];pp=p.find(q('pPr'))
  if pp is None:pp=E.Element(q('pPr'));p.insert(0,pp)
  old=pp.find(q('pStyle'));oldid=old.get(q('val')) if old is not None else (default.get(q('styleId')) if default is not None else None)
  if sid not in lookup:
   style=E.SubElement(styles,q('style'),{q('type'):'paragraph',q('customStyle'):'1',q('styleId'):sid});E.SubElement(style,q('name'),{q('val'):name});E.SubElement(style,q('qFormat'));sp=E.SubElement(style,q('pPr'));sr=E.SubElement(style,q('rPr'))
   chain=[];seen=set();cur=oldid
   while cur in lookup:
    if cur in seen:raise ValueError('样式继承循环')
    seen.add(cur);s=lookup[cur];chain.append(s);b=s.find(q('basedOn'));cur=b.get(q('val')) if b is not None else None
   for s in reversed(chain):merge(sp,s.find(q('pPr')));merge(sr,s.find(q('rPr')))
   merge(sp,pp);merge(sr,pp.find(q('rPr')))
   if i==title_index:
    runs=[r for r in p.findall(q('r')) if r.find(q('t')) is not None]
    if runs:merge(sr,runs[0].find(q('rPr')))
   lookup[sid]=style;created.append(sid)
  elif oldid!=sid and sid not in created:raise ValueError('目标样式ID已存在但来源不同，需确认')
  if old is None:old=E.Element(q('pStyle'));pp.insert(0,old)
  old.set(q('val'),sid)
  if i==title_index:
   sr=lookup[sid].find(q('rPr'))
   for rp in p.iter(q('rPr')):
    for x in list(rp):
     if x.tag not in {q('rFonts'),q('sz'),q('szCs')}:continue
     model=sr.find(x.tag)
     if model is not None and {k:v for k,v in x.attrib.items() if k!=q('hint')}=={k:v for k,v in model.attrib.items() if k!=q('hint')}:rp.remove(x)
 assert fingerprint(root)==before
 changed={i for i,_,_ in selected}
 assert all(E.tostring(p,method='c14n')==snap[i] for i,p in enumerate(paras) if i not in changed)
 replacements={'word/document.xml':E.tostring(root,encoding='utf-8',xml_declaration=True),'word/styles.xml':E.tostring(styles,encoding='utf-8',xml_declaration=True)}
 with output.open('xb') as f:
  with zipfile.ZipFile(f,'w') as z:
   for info,b in parts:z.writestr(info,replacements.get(info.filename,b))
 return {'title_index':title_index,'styles':[(i,name) for i,_,name in selected],'created':created,'checks':['仅总标题和纯图片段落格式改变','文本与非格式结构不变','既有样式定义不变','新样式无basedOn/link/autoRedefine']}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');p.add_argument('--title-index',type=int,required=True);a=p.parse_args();print(json.dumps(isolate(a.source,a.output,a.title_index),ensure_ascii=False))

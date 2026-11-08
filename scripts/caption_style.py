"""Attach caption paragraphs to the document's Caption style and label sequences."""
import copy
from lxml import etree as E
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(n):return '{'+W+'}'+n
def configure(root,styles,settings,indices):
 nodes=list(root.find(q('body')))
 matches=[s for s in styles if s.get(q('type'))=='paragraph' and s.find(q('name')) is not None and s.find(q('name')).get(q('val'),'').lower() in ('caption','题注')]
 if len(matches)>1:raise ValueError('题注样式不唯一')
 s=matches[0] if matches else E.SubElement(styles,q('style'),{q('type'):'paragraph',q('styleId'):'Caption'})
 sid=s.get(q('styleId'));new=s.find(q('name')) is None
 if new:E.SubElement(s,q('name'),{q('val'):'caption'})
 selected=[nodes[i] for i in indices]
 # Do not redefine a style already used elsewhere, or inherited by other styles.
 for p in root.iter(q('p')):
  st=p.find('./'+q('pPr')+'/'+q('pStyle'))
  if p not in selected and st is not None and st.get(q('val'))==sid:raise ValueError('题注样式被非目标段落使用')
 for other in styles:
  b=other.find(q('basedOn'))
  if b is not None and b.get(q('val'))==sid:raise ValueError('题注样式存在继承者，需明确范围')
 # Preserve the original explicitly specified font, and make it editable via Caption.
 props=('rFonts','sz','szCs')
 first=selected[0] if selected else None
 already=first is not None and first.find('./'+q('pPr')+'/'+q('pStyle')) is not None and first.find('./'+q('pPr')+'/'+q('pStyle')).get(q('val'))==sid
 if first is not None and not already:
  sample=first.find('.//'+q('r')+'/'+q('rPr'))
  for tag in ('basedOn','link','autoRedefine','pPr','rPr'):
   for x in list(s):
    if x.tag==q(tag):s.remove(x)
  pp=first.find(q('pPr'));dest=E.SubElement(s,q('pPr'))
  if pp is not None:
   for x in pp:
    if x.tag not in (q('pStyle'),q('rPr'),q('sectPr'),q('numPr')):dest.append(copy.deepcopy(x))
  rp=E.SubElement(s,q('rPr'))
  if sample is not None:
   for x in sample:
    if x.tag!=q('rStyle'):rp.append(copy.deepcopy(x))
 for tag in ('hidden','semiHidden','autoRedefine','basedOn'):
  for x in list(s):
   if x.tag==q(tag):s.remove(x)
 if s.find(q('qFormat')) is None:
  marker=E.Element(q('qFormat'));pp=s.find(q('pPr'));s.insert(s.index(pp) if pp is not None else len(s),marker)
 for p in selected:
  pp=p.find(q('pPr'))
  if pp is None:pp=E.Element(q('pPr'));p.insert(0,pp)
  st=pp.find(q('pStyle'))
  if st is None:st=E.Element(q('pStyle'));pp.insert(0,st)
  st.set(q('val'),sid)
  for rp in p.iter(q('rPr')):
   for x in list(rp):
    # Preserve differing explicit typography instead of silently homogenizing it.
    if E.QName(x).localname in props:
     target=s.find('./'+q('rPr')+'/'+x.tag)
     if target is not None and {k:v for k,v in x.attrib.items() if k!=q('hint')}=={k:v for k,v in target.attrib.items() if k!=q('hint')}:rp.remove(x)
 captions=settings.find(q('captions'))
 if captions is None:
  captions=E.Element(q('captions'));after={'readModeInkLockDown','smartTagType','schemaLibrary','shapeDefaults','doNotEmbedSmartTags','decimalSymbol','listSeparator'}
  ix=next((i for i,x in enumerate(settings) if E.QName(x).localname in after or E.QName(x).namespace!=W),len(settings));settings.insert(ix,captions)
 for label in ('图','表'):
  c=next((x for x in captions if x.tag==q('caption') and x.get(q('name'))==label),None)
  if c is None:c=E.Element(q('caption'));captions.insert(0,c)
  for k,v in {'name':label,'pos':'below','numFmt':'decimal','chapNum':'0'}.items():c.set(q(k),v)
 return sid

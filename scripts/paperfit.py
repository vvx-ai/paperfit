"""Offline, style-based DOCX formatting engine."""
import argparse, copy, json, re, zipfile
from pathlib import Path
from lxml import etree as E
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
E.register_namespace('w', W)
M='http://schemas.openxmlformats.org/officeDocument/2006/math'
def in_math(element):
    return any(isinstance(a.tag,str) and a.tag.startswith('{'+M+'}') for a in (element, *element.iterancestors()))
def math_snapshot(root):
    return [E.tostring(x,method='c14n') for x in root.iter() if isinstance(x.tag,str) and x.tag in ('{'+M+'}oMath','{'+M+'}oMathPara') and not any(a.tag in ('{'+M+'}oMath','{'+M+'}oMathPara') for a in x.iterancestors())]
def q(n): return '{'+W+'}'+n
def child(e,n):
    x=e.find(q(n))
    if x is None: x=E.SubElement(e,q(n))
    return x
def setv(e,n,attrs):
    x=child(e,n)
    for k,v in attrs.items(): x.set(q(k),str(v))
def parse(text):
    rules={'styles':{},'page':{},'pending':[]}
    sizes={'小四':12,'四号':14,'三号':16,'小三':15,'五号':10.5,'小五':9,'二号':22}
    for clause in re.split('[；;\n。]',text):
        clause=clause.strip()
        if not clause: continue
        m=re.match(r'(正文|一级标题|二级标题|三级标题)\s*(.*)',clause)
        if not m:
            m2=re.fullmatch(r'(上|下|左|右)页边距\s*([\d.]+)\s*(?:厘米|cm)',clause)
            if m2: rules['page'][{'上':'top','下':'bottom','左':'left','右':'right'}[m2[1]]]=float(m2[2])
            else: rules['pending'].append(clause)
            continue
        role={'正文':'body','一级标题':'heading1','二级标题':'heading2','三级标题':'heading3'}[m[1]]
        r=rules['styles'].setdefault(role,{})
        rest=m[2].replace('使用','').replace('单倍行距','1倍行距').strip()
        patterns=[(r'加粗','bold',lambda v:True),(r'(宋体|黑体|楷体|仿宋|Times New Roman)','font',str),(r'(小四|四号|三号|小三|五号|小五|二号)','size',lambda v:sizes[v]),(r'([\d.]+)\s*倍行距','line',float),(r'首行缩进\s*([\d.]+)\s*字符','indent_chars',float),(r'段前\s*([\d.]+)\s*磅','before',float),(r'段后\s*([\d.]+)\s*磅','after',float),(r'(两端对齐|居中|左对齐|右对齐)','align',lambda v:{'两端对齐':'both','居中':'center','左对齐':'left','右对齐':'right'}[v])]
        for pat,key,conv in patterns:
            vals=re.findall(pat,rest)
            if vals:
                value=conv(vals[0])
                if len(set(vals))>1 or (key in r and r[key]!=value):
                    rules['pending'].append(clause+'：规则冲突'); continue
                r[key]=value; rest=re.sub(pat,'',rest)
        if rest.strip(' ，,、:：'): rules['pending'].append(clause+'：未解析 '+rest)
    return rules

def props(parent,r):
    pp=child(parent,'pPr'); rp=child(parent,'rPr')
    if 'font' in r: setv(rp,'rFonts',{k:r['font'] for k in ('ascii','hAnsi','eastAsia','cs')})
    if 'bold' in r:
        for n in ('b','bCs'): setv(rp,n,{'val':int(r['bold'])})
    if 'size' in r:
        for n in ('sz','szCs'): setv(rp,n,{'val':round(r['size']*2)})
    if 'align' in r: setv(pp,'jc',{'val':r['align']})
    if 'indent_chars' in r:
        ind=child(pp,'ind')
        for k in ('hanging','hangingChars'): ind.attrib.pop(q(k),None)
        setv(pp,'ind',{'firstLineChars':round(r['indent_chars']*100),'firstLine':round(r['indent_chars']*r.get('size',12)*20)})
    attrs={k:round(r[k]*20) for k in ('before','after') if k in r}
    if 'line' in r:
        attrs.update(line=round(r['line']*240),lineRule='auto')
        setv(pp,'snapToGrid',{'val':0})
    if attrs: setv(pp,'spacing',attrs)
    return pp,rp

def fingerprint(root):
    c=copy.deepcopy(root)
    for e in list(c.iter()):
        for x in list(e):
            if x.tag in {q('pPr'),q('rPr'),q('pgMar')}: e.remove(x)
    def signature(e): return (e.tag, tuple(sorted(e.attrib.items())), e.text, e.tail, tuple(signature(x) for x in e))
    return signature(c)

def format_docx(source,output,rules):
    source,output=Path(source),Path(output)
    if source.resolve()==output.resolve() or output.exists(): raise ValueError('输出必须为不存在的新文件')
    if rules.get('pending'): raise ValueError('请先解决 pending 规则')
    allowed={'font','size','line','indent_chars','before','after','align','bold'}
    for role,r in rules.get('styles',{}).items():
        if role not in ('body','heading1','heading2','heading3') or set(r)-allowed: raise ValueError('不支持的样式或属性')
        for k,v in r.items():
            if k not in ('font','align') and (not isinstance(v,(int,float)) or not 0<=v<=200): raise ValueError('无效数值')
    if set(rules.get('page',{}))-{'top','bottom','left','right'}: raise ValueError('不支持的页面属性')
    with zipfile.ZipFile(source) as z: parts={i.filename:(i,z.read(i)) for i in z.infolist()}
    root=E.fromstring(parts['word/document.xml'][1]); original=fingerprint(root); original_math=math_snapshot(root)
    styles=E.fromstring(parts['word/styles.xml'][1])
    report={'changes':[],'pending':[],'integrity':False}
    style_ids={}
    for role in rules.get('styles',{}):
        default=('Heading'+role[-1]) if role.startswith('heading') else 'BodyText'
        matches=[]
        if True:
            for candidate in styles:
                name=candidate.find(q('name'))
                if candidate.get(q('type'))=='paragraph' and name is not None and name.get(q('val'),'').lower().replace(' ','') in (('heading'+role[-1], '标题'+role[-1], {'heading1':'一级标题','heading2':'二级标题','heading3':'三级标题'}[role]) if role.startswith('heading') else ('bodytext','正文')): matches.append(candidate)
        if len(matches)>1:
            used={x.get(q('val')) for x in root.iter(q('pStyle'))}
            active=[x for x in matches if x.get(q('styleId')) in used]
            if len(active)==1: matches=active
            else: raise ValueError('存在多个候选样式，需要明确选择')
        style_ids[role]=matches[0].get(q('styleId')) if matches else default

    # Preflight shared styles before any mutation, including inherited consumers.
    paragraphs=root.findall('./'+q('body')+'/'+q('p'))
    def role_for(p,i):
        explicit=rules.get('paragraph_roles',{}).get(str(i))
        if explicit is not None: return explicit
        st=p.find('./'+q('pPr')+'/'+q('pStyle'))
        sid=st.get(q('val'),'') if st is not None else ''
        for role,value in style_ids.items():
            if sid==value:return role
        text=''.join(p.itertext()).strip()
        return 'heading1' if re.match(r'^第[一二三四五六七八九十百\d]+章',text) else 'body'
    def supported(p):
        return bool(''.join(x.text or '' for x in p.iter(q('t'))).strip()) and not any(p.find('.//'+q(n)) is not None for n in ('fldSimple','fldChar','instrText','drawing','object','ins','del','pict'))
    targets={p for i,p in enumerate(paragraphs) if supported(p) and role_for(p,i) in rules.get('styles',{})}
    parents={x.get(q('styleId')):x.find(q('basedOn')).get(q('val')) for x in styles if x.find(q('basedOn')) is not None}
    changed=set(style_ids.values())
    defaults=[x.get(q('styleId')) for x in styles if x.get(q('type'))=='paragraph' and x.get(q('default'))=='1']
    for name,(_,data) in parts.items():
        if not name.startswith('word/') or not name.endswith('.xml') or name=='word/styles.xml':continue
        tree=root if name=='word/document.xml' else E.fromstring(data)
        for p in tree.iter(q('p')):
            if name=='word/document.xml' and p in targets:continue
            st=p.find('./'+q('pPr')+'/'+q('pStyle'))
            sid=st.get(q('val')) if st is not None else (defaults[0] if defaults else None)
            seen=set()
            while sid and sid not in seen:
                if sid in changed:raise ValueError('共享样式影响非目标段落：'+name+' / '+sid+'；请扩大明确范围或调整样式配置')
                seen.add(sid);sid=parents.get(sid)

    for role,r in rules.get('styles',{}).items():
        sid=style_ids[role]
        s=next((x for x in styles if x.get(q('styleId'))==sid),None)
        if s is None:
            s=E.SubElement(styles,q('style'),{q('type'):'paragraph',q('styleId'):sid}); setv(s,'name',{'val':('heading '+role[-1]) if role.startswith('heading') else 'Body Text'})
        props(s,r)
        if True:
            child(s,'qFormat')
            for flag in ('semiHidden','hidden','unhideWhenUsed'):
                node=s.find(q(flag))
                if node is not None: s.remove(node)
            rf=s.find('./'+q('rPr')+'/'+q('rFonts'))
            if rf is not None and 'font' in r:
                for a in list(rf.attrib):
                    if 'Theme' in a: del rf.attrib[a]
        if role.startswith('heading'): setv(child(s,'pPr'),'outlineLvl',{'val':int(role[-1])-1})
    for i,p in enumerate(root.findall('./'+q('body')+'/'+q('p'))):
        text=''.join(p.itertext()) if False else ''.join(x.text or '' for x in p.iter(q('t')))
        if not text.strip(): continue
        if not supported(p):
            report['pending'].append({'paragraph':i,'reason':'复杂对象段落保留'}); continue
        ps=p.find('./'+q('pPr')+'/'+q('pStyle')); sid=ps.get(q('val'),'') if ps is not None else ''
        h=re.search(r'(?:Heading|heading|标题|PaperFit_heading)\s*([123])',sid)
        role=next((role for role,style_id in style_ids.items() if style_id==sid), 'heading'+h[1] if h else 'body')
        role=rules.get('paragraph_roles',{}).get(str(i),role)
        if role=='preserve': continue
        if str(i) not in rules.get('paragraph_roles',{}) and not h and re.match(r'^第[一二三四五六七八九十百\d]+章',text): role='heading1'
        if role not in rules.get('styles',{}): continue
        r=rules['styles'][role]; pp,rp=props(p,r)
        p.remove(rp)
        if list(p)[0] is not pp: p.remove(pp); p.insert(0,pp)
        target=style_ids[role]
        setv(pp,'pStyle',{'val':target})
        # Remove only properties controlled by the selected rule.
        for name,keys in {'spacing':('line','before','after'),'ind':('indent_chars',),'jc':('align',),'snapToGrid':('line',)}.items():
            x=pp.find(q(name))
            if x is not None and any(k in r for k in keys):
                if name=='spacing':
                    attrs=[]
                    if 'line' in r: attrs+=['line','lineRule']
                    if 'before' in r: attrs+=['before','beforeLines','beforeAutospacing']
                    if 'after' in r: attrs+=['after','afterLines','afterAutospacing']
                    for a in attrs: x.attrib.pop(q(a),None)
                    if not x.attrib: pp.remove(x)
                elif name=='ind':
                    for a in ('firstLine','firstLineChars','hanging','hangingChars'): x.attrib.pop(q(a),None)
                    if not x.attrib: pp.remove(x)
                else: pp.remove(x)
        for rr in p.iter(q('rPr')):
            if in_math(rr): continue
            for x in list(rr):
                if x.tag in {v.tag for v in rp}: rr.remove(x)
            cs=rr.find(q('rStyle'))
            if cs is not None:
                # Keep semantic character formatting; override only requested attributes.
                for x in rp: rr.append(copy.deepcopy(x))
                report['pending'].append({'paragraph':i,'reason':'保留字符样式，指定属性使用直接覆盖'})
        st=pp.find(q('pStyle')); pp.remove(st); pp.insert(0,st)
        report['changes'].append({'paragraph':i,'role':role,'before_style':sid,'applied':r})
    for sect in root.iter(q('sectPr')):
        for side,value in rules.get('page',{}).items():
            if not isinstance(value,(int,float)) or not 0<=value<=10: raise ValueError('页边距无效')
            mar=child(sect,'pgMar'); old=mar.get(q(side)); mar.set(q(side),str(round(value/2.54*1440)))
            report['changes'].append({'page_margin':side,'before':old,'after_cm':value})
    if fingerprint(root)!=original: raise ValueError('内容完整性校验失败')
    if math_snapshot(root)!=original_math: raise ValueError('公式结构完整性校验失败')
    report['integrity']=True
    report['math_xml_unchanged']=True
    replacements={'word/document.xml':E.tostring(root,encoding='utf-8',xml_declaration=True),'word/styles.xml':E.tostring(styles,encoding='utf-8',xml_declaration=True)}
    with output.open('xb') as f:
        with zipfile.ZipFile(f,'w') as z:
            for name,(info,data) in parts.items(): z.writestr(info,replacements.get(name,data))
    return report

def main(argv=None):
    p=argparse.ArgumentParser(epilog='Phase 2: extract SOURCE RULES.json | audit INPUT.docx AUDIT.json | run INPUT.docx RULES.json OUTPUT.docx'); sub=p.add_subparsers(dest='cmd',required=True)
    a=sub.add_parser('toc', help='生成一至三级 Word 原生目录'); a.add_argument('source'); a.add_argument('output')
    a=sub.add_parser('paginate', help='目录隐藏页码，正文页脚居中从 1 编号'); a.add_argument('source'); a.add_argument('output')
    a=sub.add_parser('flow', help='图表题注同页、一级标题另起一节、页眉页脚独立'); a.add_argument('source'); a.add_argument('output')
    a=sub.add_parser('parse'); a.add_argument('text'); a.add_argument('output')
    a=sub.add_parser('format'); a.add_argument('source'); a.add_argument('rules'); a.add_argument('output')
    a=p.parse_args(argv)
    if a.cmd in ('toc', 'paginate', 'flow'):
        from toc import add_toc
        from pagination import paginate
        from page_flow import enforce
        try:
            operation = {'toc':add_toc, 'paginate':paginate, 'flow':enforce}[a.cmd]
            operation(a.source,a.output)
        except (ValueError, KeyError, OSError) as error:
            p.exit(2, f'PaperFit: {error}\n')
    elif a.cmd=='parse': Path(a.output).write_text(json.dumps(parse(Path(a.text).read_text()),ensure_ascii=False,indent=2))
    else:
        reportpath=Path(a.output+'.report.json')
        if reportpath.exists(): raise ValueError('报告已存在')
        report=format_docx(a.source,a.output,json.loads(Path(a.rules).read_text()))
        reportpath.write_text(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':
    from cli import main as cli_main
    cli_main()

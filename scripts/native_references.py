"""Conservative native REF links and real Word bibliography list numbering.

Only unique caption labels and a reviewed-looking contiguous [1]...[n]
list below a unique bibliography heading are inferred. Ambiguity is reported.
"""
from collections import defaultdict
from copy import deepcopy
from hashlib import sha256
import json
import os
from pathlib import Path
import re
import shlex
from tempfile import TemporaryDirectory
from lxml import etree as E
from paperfit import q, math_snapshot
from structured import element, run, visible, CAP
from toc import val, heading_level
from word_fields import scan, bookmark_ranges

DOC='word/document.xml'
NUMBERING='word/numbering.xml'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
BIB_NAME='PaperFitBibliographyBrackets'
MENTION=re.compile(r'([图表])\s*(\d+(?:[-－.]\d+)*)(?!\d)|\[(\d+)\]')
LABEL=re.compile(r'^\[(\d+)\]')
COMBINED=re.compile(r'\[(\d+(?:\s*[-–,，、]\s*\d+)+)\]')
CITATION_STYLE='PaperFitCitation'


def superscript_citations(root, styles, bib_paras, changes, pending):
    """Format only REF number results pointing into the bibliography.

    A character style leaves body/caption/bibliography paragraph styles intact.
    MERGEFORMAT preserves the result formatting when Word recalculates a REF.
    """
    from phase2 import canonical
    ranges=bookmark_ranges(root)
    starts={b.get(q('name')):next(b.iterancestors(q('p')),None) for b in root.iter(q('bookmarkStart'))}
    for f in scan(root)[0]:
        try:tokens=[x.strip('"') for x in shlex.split(f.code,posix=False)]
        except ValueError:continue
        if len(tokens)<3 or tokens[0].upper()!='REF' or '\\n' not in tokens[2:] or starts.get(tokens[1]) not in bib_paras:continue
        if len(ranges.get(tokens[1],[]))!=1 or ranges[tokens[1]][0] is None:continue
        p=next(f.owner.iterancestors(q('p')),None)
        if p in bib_paras or any(a.tag==q('sdt') for a in f.owner.iterancestors()):continue
        if f.nested or not re.fullmatch(r'\[\d+\]',f.value):
            pending.append({'field':f.code,'reason':'文献引用域结构或显示结果复杂，上标需核对'});continue
        runs=list(dict.fromkeys(next(t.iterancestors(q('r'))) for t in f.result))
        if any(any(t not in f.result for t in r.iter(q('t'))) for r in runs):
            pending.append({'field':f.code,'reason':'引用与其他文字共享运行，上标需核对'});continue
        style=next((s for s in styles.findall(q('style')) if s.get(q('styleId'))==CITATION_STYLE),None)
        if style is None:
            named=[s for s in styles.findall(q('style')) if val(s,'name')=='文献引用']
            if len(named)>1:raise ValueError('文献引用字符样式重名，请核对')
            if named:style=named[0] # Word may rewrite style IDs on save.
        if style is None:
            style=element('style',type='character',styleId=CITATION_STYLE,customStyle='1')
            style.extend([element('name',val='文献引用'),element('qFormat')])
            rp=element('rPr');rp.append(element('vertAlign',val='superscript'));style.append(rp);styles.append(style)
        elif style.get(q('type'))!='character' or val(style,'name')!='文献引用' or val(style,'rPr/vertAlign')!='superscript':
            raise ValueError('文献引用字符样式冲突，请核对')
        style_id=style.get(q('styleId'))
        before=[canonical(r) for r in runs];old_code=f.code
        for r in runs:
            rp=r.find(q('rPr'))
            if rp is None:rp=element('rPr');r.insert(0,rp)
            st=rp.find(q('rStyle'))
            if st is None:st=element('rStyle',val=style_id);rp.insert(0,st)
            for n in list(rp):
                if n.tag in (q('vertAlign'),q('position')):rp.remove(n)
            # Preserve an existing character style; override only vertical alignment.
            if st.get(q('val'))!=style_id:
                successors={q(n) for n in ('rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange')}
                rp.insert(next((i for i,n in enumerate(rp) if n.tag in successors),len(rp)),element('vertAlign',val='superscript'))
        if not re.search(r'\\\*\s+MERGEFORMAT\b',f.code,re.I):f.set_code(f.code+' \\* MERGEFORMAT')
        if before!=[canonical(r) for r in runs] or old_code!=f.code:
            changes.append({'kind':'citation_superscript','display':f.value,'bookmark':tokens[1],'style':style_id,'field':f.code})


def splice(p, start, end, replacement=None):
    """Replace a text span across simple runs, preserving surrounding run props."""
    cursor=0;covered=[]
    for child in p:
        text=visible(child)
        a,b=cursor,cursor+len(text);cursor=b
        if a < b and a < end and b > start:covered.append((child,a,b,text))
    if not covered:raise ValueError('文字范围缺失')
    for child,_,_,_ in covered:
        if child.tag!=q('r') or len(child.findall(q('t')))!=1 or any(n.tag not in (q('rPr'),q('t')) for n in child):
            raise ValueError('标签跨越域、超链接或非简单文字节点')
    first,last=covered[0][0],covered[-1][0]
    between=list(p)[p.index(first):p.index(last)+1]
    text_runs=[x[0] for x in covered]
    if any(n not in text_runs and n.tag!=q('proofErr') for n in between):
        raise ValueError('标签跨越书签或其他范围')
    def fragment(old,text):
        new=deepcopy(old);new.find(q('t')).text=text
        new.find(q('t')).set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        return new
    left=covered[0][3][:start-covered[0][1]];right=covered[-1][3][end-covered[-1][1]:]
    replacements=[]
    if left:replacements.append(fragment(first,left))
    if replacement is not None:
        rp=first.find(q('rPr'))
        if rp is not None:
            target=replacement.find(q('r'))
            if target is not None:target.insert(0,deepcopy(rp))
        replacements.append(replacement)
    if right:replacements.append(fragment(last,right))
    pos=p.index(first)
    # Proofing markers are zero-width metadata. Retain them, including paired
    # markers whose other end lies outside the replaced label.
    for child in text_runs:p.remove(child)
    for i,child in enumerate(replacements):p.insert(pos+i,child)


def native_list(parts, style):
    from phase2 import xml
    numbering=xml(parts[NUMBERING]) if NUMBERING in parts else element('numbering')
    abstracts=numbering.findall(q('abstractNum'));nums=numbering.findall(q('num'))
    aid=max([int(n.get(q('abstractNumId'))) for n in abstracts],default=-1)+1
    nid=max([int(n.get(q('numId'))) for n in nums],default=0)+1
    abstract=element('abstractNum',abstractNumId=aid)
    abstract.extend([element('multiLevelType',val='singleLevel'),element('name',val=BIB_NAME)])
    lvl=element('lvl',ilvl=0)
    lvl.extend([element('start',val=1),element('numFmt',val='decimal'),element('suff',val='nothing'),element('lvlText',val='[%1]'),element('lvlJc',val='left')])
    # The number adopts the bibliography font/size while the paragraph retains its style.
    if style is not None and style.find(q('rPr')) is not None:lvl.append(deepcopy(style.find(q('rPr'))))
    abstract.append(lvl)
    numbering.insert(next((i for i,n in enumerate(numbering) if n.tag==q('num')),len(numbering)),abstract)
    num=element('num',numId=nid);num.append(element('abstractNumId',val=aid));numbering.append(num)
    parts[NUMBERING]=E.tostring(numbering,xml_declaration=True,encoding='UTF-8')
    rels=xml(parts['word/_rels/document.xml.rels']);types=xml(parts['[Content_Types].xml'])
    if not any(n.get('Type')==REL+'/numbering' for n in rels):
        existing={n.get('Id') for n in rels};rid='rIdPaperFitNumbering'
        while rid in existing:rid+='x'
        E.SubElement(rels,'{'+PKG+'}Relationship',Id=rid,Type=REL+'/numbering',Target='numbering.xml')
    if not any(n.get('PartName')=='/word/numbering.xml' for n in types):
        E.SubElement(types,'{'+CT+'}Override',PartName='/word/numbering.xml',ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
    parts['word/_rels/document.xml.rels']=E.tostring(rels,xml_declaration=True,encoding='UTF-8')
    parts['[Content_Types].xml']=E.tostring(types,xml_declaration=True,encoding='UTF-8')
    return str(nid)


def owned_numbering(parts):
    from phase2 import xml
    if NUMBERING not in parts:return set()
    root=xml(parts[NUMBERING]);abstracts={n.get(q('abstractNumId')) for n in root.findall(q('abstractNum')) if val(n,'name')==BIB_NAME and val(n,'lvl/lvlText')=='[%1]' and val(n,'lvl/start')=='1' and val(n,'lvl/numFmt')=='decimal' and len(n.findall(q('lvl')))==1}
    # A user can restart numbering in Word. Do not infer its displayed numbers
    # from paragraph order when an override changes the original list definition.
    return {n.get(q('numId')) for n in root.findall(q('num')) if val(n,'abstractNumId') in abstracts and n.find(q('lvlOverride')) is None}


def convert_parts(parts):
    """Return a new package and report. Never mutate the caller's package."""
    from phase2 import xml,canonical
    parts=dict(parts);original=dict(parts);root=xml(parts[DOC]);body=root.find(q('body'))
    styles=xml(parts['word/styles.xml']);defs={s.get(q('styleId')):s for s in styles.findall(q('style'))}
    if any(list(root.iter(q(t))) for t in ('ins','del','moveFrom','moveTo')) or any(k.startswith('_xmlsignatures/') for k in parts):
        raise ValueError('修订或签名文档需先核对')
    settings=xml(parts['word/settings.xml'])
    if any(settings.find(q(t)) is not None for t in ('documentProtection','trackRevisions')):raise ValueError('保护或修订已开启')
    fields,errors=scan(root)
    if errors:raise ValueError('存在未闭合域')
    expected={p:visible(p) for p in root.iter(q('p'))}
    before_math=math_snapshot(root)
    before_objects={tag:[canonical(n) for n in root.iter(q(tag))] for tag in ('drawing','object','pict')}
    # Run children are changed; all table cell text is still checked by the text ledger.
    changes=[];pending=[];targets=defaultdict(list);caption_paras=set();bib_paras=[]
    names={b.get(q('name')) for b in root.iter(q('bookmarkStart'))}
    nextid=max([int(b.get(q('id'))) for b in root.iter(q('bookmarkStart'))],default=-1)+1
    def bookmark(p,prefix,end=None):
        nonlocal nextid
        for start in p.findall(q('bookmarkStart')):
            name=start.get(q('name'),'')
            if name.startswith(prefix) and len(bookmark_ranges(root).get(name,[]))==1:
                ranges=bookmark_ranges(root)[name]
                if ranges[0] is not None:return name
        while f'{prefix}{nextid}' in names:nextid+=1
        name=f'{prefix}{nextid}';key=nextid;nextid+=1;names.add(name)
        if end is None:
            pos=len(p)
        else:
            # Include only the label and SEQ, excluding caption title text.
            count=0;pos=None
            for i,child in enumerate(list(p)):
                value=visible(child)
                if count+len(value)>end:
                    if child.tag!=q('r') or len(child.findall(q('t')))!=1 or any(n.tag not in (q('rPr'),q('t')) for n in child):raise ValueError('题注编号边界结构不支持')
                    cut=end-count;left=deepcopy(child);right=deepcopy(child)
                    left.find(q('t')).text=value[:cut];right.find(q('t')).text=value[cut:]
                    for x in (left,right):x.find(q('t')).set('{http://www.w3.org/XML/1998/namespace}space','preserve')
                    p.remove(child);p.insert(i,left);p.insert(i+1,right);pos=i+1;break
                count+=len(value)
                if count==end:
                    pos=i+1
                    # Word-save complex SEQ fields close in a zero-text run.
                    while pos<len(p) and not visible(p[pos]) and p[pos].tag==q('r'):
                        pos+=1
                    break
            if pos is None:raise ValueError('题注编号范围缺失')
        p.insert(pos,element('bookmarkEnd',id=key))
        p.insert(1 if p.find(q('pPr')) is not None else 0,element('bookmarkStart',id=key,name=name))
        return name
    # Caption candidates require the native caption style AND matching SEQ label.
    for p in body.findall(q('p')):
        fs,_=scan(p);m=CAP.match(visible(p));s=defs.get(val(p,'pPr/pStyle'))
        if m and s is not None and val(s,'name','').lower() in ('caption','题注'):
            seqs=[f for f in fs if f.code.split()[:2]==['SEQ',m[1]]]
            if len(seqs)==1:
                caption_paras.add(p);targets[(m[1],m[2])].append((p,m.end(2),m[0].strip(),'_PaperFitCap'))
    headings=[p for p in body.findall(q('p')) if re.sub(r'\s+','',visible(p)).lower() in ('参考文献','references')]
    if len(headings)==1:
        for n in list(body)[body.index(headings[0])+1:]:
            if n.tag==q('sectPr'):break
            if n.tag!=q('p'):
                if n.tag!=q('sdt'):pending.append({'reason':'参考文献区域含非段落对象'})
                break
            if heading_level(n,defs):break
            if visible(n).strip():bib_paras.append(n)
    elif len(headings)>1:pending.append({'reason':'参考文献标题不唯一，未推断条目'})
    else:
        bib_paras=[p for p in body.findall(q('p')) if val(defs.get(val(p,'pPr/pStyle'),element('style')),'name','') in ('参考文献条目','参考文献正文')]
    own=owned_numbering(parts);manual=[LABEL.match(visible(p)) for p in bib_paras]
    all_manual=bool(bib_paras) and all(manual) and [int(m[1]) for m in manual]==list(range(1,len(bib_paras)+1))
    all_native=bool(bib_paras) and all(val(p,'pPr/numPr/numId') in own and val(p,'pPr/numPr/ilvl','0')=='0' for p in bib_paras)
    if all_native:
        ids={val(p,'pPr/numPr/numId') for p in bib_paras}
        all_native=len(ids)==1 and [p for p in body.iter(q('p')) if val(p,'pPr/numPr/numId') in ids]==bib_paras
    if all_manual:
        try:
            for p,m in zip(bib_paras,manual):
                if p.find('./'+q('pPr')+'/'+q('numPr')) is not None or scan(p)[0]:raise ValueError('文献含现有列表或域，需要核对')
                # Validate every span before changing any entry.
                splice(deepcopy(p),0,m.end())
            nid=native_list(parts,defs.get(val(bib_paras[0],'pPr/pStyle')))
            for p,m in zip(bib_paras,manual):
                old=visible(p);splice(p,0,m.end());expected[p]=old[m.end():]
                pp=p.find(q('pPr'))
                if pp is None:pp=element('pPr');p.insert(0,pp)
                num=element('numPr');num.extend([element('ilvl',val=0),element('numId',val=nid)])
                pp.insert(1 if pp.find(q('pStyle')) is not None else 0,num)
                from page_flow import P_ORDER
                pp[:]=sorted(pp,key=lambda n:P_ORDER.index(E.QName(n).localname) if E.QName(n).localname in P_ORDER else 999)
                changes.append({'kind':'bibliography_number','before':m[0],'after':'native_list','numId':nid})
            all_native=True
        except ValueError as error:pending.append({'reason':str(error)})
    elif bib_paras and not all_native:
        pending.append({'reason':'参考文献编号不是连续唯一的[1]至[n]，或含未知自动列表；保留原状'})
    if all_native:
        for i,p in enumerate(bib_paras,1):targets[('bib',str(i))].append((p,None,f'[{i}]','_PaperFitBib'))
    # Existing native fields are left intact, including their bookmarks and caches.
    cached={n for f in scan(root)[0] for n in f.result}
    all_paras=list(body.iter(q('p')))
    for idx,p in enumerate(all_paras):
        if p in caption_paras or p in bib_paras or any(a.tag==q('sdt') for a in p.iterancestors()):continue
        text=visible(p);offsets=[];offset=0
        for t in p.iter(q('t')):
            offsets.append((offset,offset+len(t.text or ''),t));offset+=len(t.text or '')
        for match in COMBINED.finditer(text):
            if re.fullmatch(r'\d{4}-\d{1,2}-\d{1,2}',match[1]):continue
            if not any(t in cached for a,b,t in offsets if a<match.end() and b>match.start()):
                pending.append({'paragraph':idx,'label':match[0],'reason':'组合文献引用需逐项确认，保留原状'})
        for match in reversed(list(MENTION.finditer(text))):
            covered=[t for a,b,t in offsets if a<match.end() and b>match.start()]
            if any(t in cached for t in covered):continue
            key=(match[1],match[2]) if match[1] else ('bib',match[3])
            candidates=targets.get(key,[])
            location={'paragraph':idx,'label':match[0]}
            if len(candidates)!=1:
                pending.append(dict(location,reason='引用没有唯一目标'));continue
            if match[1] and re.match(r'\s*[、,，至到~～–—]\s*\d',text[match.end():]):
                pending.append(dict(location,reason='组合引用需逐项确认'));continue
            target,end,label,prefix=candidates[0]
            try:
                # Do not create a target bookmark if the source span is unsupported.
                splice(deepcopy(p),match.start(),match.end(),element('fldSimple'))
                name=bookmark(target,prefix,end)
                code=f' REF {name}'+(' \\n' if key[0]=='bib' else '')+' \\h '
                field=element('fldSimple',instr=code);field.append(run(label))
                splice(p,match.start(),match.end(),field)
                expected[p]=expected[p][:match.start()]+label+expected[p][match.end():]
                changes.append(dict(location,kind='cross_reference',bookmark=name,field=code.strip(),display=label))
            except ValueError as error:pending.append(dict(location,reason=str(error)))
    superscript_citations(root,styles,bib_paras,changes,pending)
    if canonical(styles)!=canonical(xml(original['word/styles.xml'])):
        parts['word/styles.xml']=E.tostring(styles,xml_declaration=True,encoding='utf-8')
    if expected!={p:visible(p) for p in root.iter(q('p'))}:raise ValueError('交叉引用转换改变了授权范围以外的文字')
    if math_snapshot(root)!=before_math:raise ValueError('公式内容改变')
    if any(before_objects[t]!=[canonical(n) for n in root.iter(q(t))] for t in before_objects):raise ValueError('图片或对象改变')
    if scan(root)[1]:raise ValueError('新域结构未闭合')
    from word_fields import audit
    field_audit=audit({DOC:root})
    for ref in field_audit['references']:
        if ref['status']=='目标缺失、重复或未闭合':
            pending.append({'field':ref['field'],'reason':ref['status']})
    if canonical(root)!=canonical(xml(original[DOC])):parts[DOC]=E.tostring(root,xml_declaration=True,encoding='utf-8')
    report={'changes':changes,'pending':pending,'native_bibliography_entries':len(bib_paras) if all_native else 0,
            'cross_references_created':sum(c['kind']=='cross_reference' for c in changes),
            'citations_superscripted':sum(c['kind']=='citation_superscript' for c in changes),'text_integrity':True,'field_audit':field_audit,
            'status':'needs_review' if pending else 'requires_field_update',
            'note':'文献序号为原生列表，引用为REF域；增删条目后更新域并核对正文引用。'}
    return parts,report


def convert(source,output):
    from phase2 import read_package,save_package
    source,output=Path(source),Path(output);report_path=Path(str(output)+'.report.json')
    if output.exists() or report_path.exists() or source.resolve()==output.resolve():raise ValueError('请使用新的输出文件')
    parts,report=convert_parts(read_package(source))
    with TemporaryDirectory(dir=output.parent) as tmp:
        staged=Path(tmp)/'result.docx';save_package(parts,{},staged)
        report.update(source_sha256=sha256(source.read_bytes()).hexdigest(),output_sha256=sha256(staged.read_bytes()).hexdigest())
        record=Path(tmp)/'report.json';record.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        os.link(staged,output)
        try:os.link(record,report_path)
        except BaseException:output.unlink();raise
    return report

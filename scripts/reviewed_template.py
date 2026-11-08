"""Apply an explicitly reviewed template manifest without reclassifying content.

The manifest is bound to the source SHA-256. This is deliberately not an
unsupervised template recognizer. All block indexes refer to the original body.
"""
import argparse
from copy import deepcopy
from hashlib import sha256
import json
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from lxml import etree as E

from paperfit import q, props, in_math, math_snapshot
from phase2 import DOC, STYLES, read_package, save_package, xml, canonical
from phase3 import REVISIONS
from structured import element, run, seq_field, visible
from toc import val, heading_level, entry_styles, toc_block, marker, instruction, CODE
from pagination import add_footer, RELS
from page_flow import (properties, bind_captions, split_chapters, keep_subheadings,
                       independent_stories, disable_picture_compression, P_ORDER, WP, A, PIC)
from match_captions import match_captions
from caption_style import configure
from word_fields import scan


def style_definition(styles, sid, name, rule, outline=9):
    existing = [s for s in styles.findall(q('style')) if s.get(q('styleId')) == sid]
    s = existing[0] if existing else element('style', type='paragraph', styleId=sid)
    if not existing:
        styles.append(s)
    # Each controlled style is independent; cover style inheritance is untouched.
    for n in list(s):
        s.remove(n)
    s.extend([element('name', val=name), element('next', val='PaperFitBody'), element('qFormat')])
    pp, rp = props(s, rule)
    fonts = rp.find(q('rFonts'))
    if rule.get('latin_font'):
        for key in ('ascii', 'hAnsi'):
            fonts.set(q(key), rule['latin_font'])
    pp.append(element('outlineLvl', val=outline))
    if outline < 3:
        pp.extend([element('keepNext'), element('keepLines')])
    return s


def apply_style(p, sid):
    pp = properties(p)
    # Controlled paragraphs get layout from the paragraph style. Preserve native
    # sections, numbering, bookmarks, inline objects, semantic run properties.
    for n in list(pp):
        if n.tag not in {q('sectPr'), q('numPr')}:
            pp.remove(n)
    pp.insert(0, element('pStyle', val=sid))
    for rp in p.iter(q('rPr')):
        if in_math(rp):
            continue
        for n in list(rp):
            if n.tag in {q(t) for t in ('rStyle','rFonts','sz','szCs','b','bCs','color','spacing','kern','position','snapToGrid')}:
                rp.remove(n)


def remove_suffix(p, suffix):
    text = visible(p)
    if not suffix or not text.endswith(suffix):
        raise ValueError('说明后缀与审核清单不一致')
    remaining = len(text) - len(suffix)
    for t in p.iter(q('t')):
        content = t.text or ''
        t.text = content[:max(0, remaining)]
        remaining -= len(content)


def three_line_table(table, styles, header_rows=1):
    rows = table.findall(q('tr'))
    if not 0 < header_rows < len(rows) or len(list(table.iter(q('tbl')))) != 1:
        raise ValueError('三线表需要明确表头行，且暂不支持嵌套表格')
    sid = 'PaperFitThreeLine'
    s = next((n for n in styles.findall(q('style')) if n.get(q('styleId')) == sid), None)
    if s is None:
        s = element('style', type='table', styleId=sid)
        s.append(element('name', val='三线表'))
        styles.append(s)
    tp = table.find(q('tblPr'))
    if tp is None:
        tp = element('tblPr'); table.insert(0, tp)
    for n in list(tp):
        if n.tag in {q('tblBorders'), q('tblStyle'), q('tblLook')}:
            tp.remove(n)
    tp.insert(0, element('tblStyle', val=sid))
    borders = element('tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        borders.append(element(edge, val='single' if edge in ('top','bottom') else 'nil',
                               sz=12 if edge in ('top','bottom') else 0, color='000000'))
    tp.append(borders)
    # Explicit cell borders defeat imported style/row overrides in Word and WPS.
    for row_index, row in enumerate(rows):
        for ex in row.findall(q('tblPrEx')):
            for b in ex.findall(q('tblBorders')):
                ex.remove(b)
        for cell in row.findall(q('tc')):
            cp = cell.find(q('tcPr'))
            if cp is None:
                cp = element('tcPr'); cell.insert(0, cp)
            for b in cp.findall(q('tcBorders')):
                cp.remove(b)
            cb = element('tcBorders')
            for edge in ('top','left','bottom','right','insideH','insideV','tl2br','tr2bl'):
                size = 12 if (edge == 'top' and row_index == 0) or (edge == 'bottom' and row_index == len(rows)-1) else 6 if edge == 'bottom' and row_index == header_rows-1 else 0
                cb.append(element(edge, val='single' if size else 'nil', sz=size, color='000000'))
            cp.append(cb)


def insert_toc(root, styles, anchor, width):
    body = root.find(q('body'))
    if any(f.code.strip().upper().startswith('TOC ') for f in scan(root)[0]):
        raise ValueError('已有目录，请使用目录更新入口')
    defs = {s.get(q('styleId')):s for s in styles.findall(q('style'))}
    headings = [(p, heading_level(p,defs)) for p in body.findall(q('p')) if heading_level(p,defs)]
    ids = entry_styles(styles, {level for _,level in headings}, width)
    for level,sid in ids.items():
        s = next(s for s in styles.findall(q('style')) if s.get(q('styleId')) == sid)
        pp,rp = props(s, {'font':'宋体','size':12,'bold':False,'line':1,'before':0,'after':0,'indent_chars':0})
        rf = rp.find(q('rFonts'))
        for k in ('ascii','hAnsi'):rf.set(q(k),'Times New Roman')
        for n in list(s):
            if n.tag in {q('basedOn'),q('autoRedefine')}:s.remove(n)
        for tag in ('tabs','ind','outlineLvl'):
            for n in pp.findall(q(tag)):pp.remove(n)
        tabs=element('tabs');tabs.append(element('tab',val='right',leader='dot',pos=width))
        pp.extend([tabs,element('ind',left=(level-1)*280,firstLine=0),element('outlineLvl',val=9)])
    block,content=toc_block()
    apply_style(content[0], 'PaperFitFrontTitle')
    body.insert(body.index(anchor),block)
    next_id=max([int(x.get(q('id'))) for x in root.iter(q('bookmarkStart'))],default=-1)+1
    names={x.get(q('name')) for x in root.iter(q('bookmarkStart'))}
    for i,(heading,level) in enumerate(headings):
        while f'_PaperFitToc{next_id}' in names:next_id+=1
        name=f'_PaperFitToc{next_id}'
        heading.insert(1,element('bookmarkStart',id=next_id,name=name));heading.append(element('bookmarkEnd',id=next_id));next_id+=1
        p=element('p');p.append(element('pPr'));p[0].append(element('pStyle',val=ids[level]))
        if i==0:
            marker(p,'begin');instruction(p,CODE);marker(p,'separate')
        link=element('hyperlink',anchor=name);link.append(run(visible(heading)))
        r=element('r');r.append(element('tab'));link.append(r)
        marker(link,'begin');instruction(link,f' PAGEREF {name} \\h ');marker(link,'separate');link.append(run('…'));marker(link,'end')
        p.append(link)
        if i==len(headings)-1:marker(p,'end')
        content.append(p)
    return block


def format_template(source, manifest, output):
    source,output=Path(source),Path(output)
    report_path=Path(str(output)+'.report.json')
    if output.exists() or report_path.exists() or output.resolve()==source.resolve():
        raise ValueError('必须使用新的输出文件')
    if sha256(source.read_bytes()).hexdigest()!=manifest['source_sha256']:
        raise ValueError('原文件已经变化，请重新审核清单')
    if 'native_references' in manifest and type(manifest['native_references']) is not bool:
        raise ValueError('native_references 必须为布尔值')
    parts=read_package(source);original=dict(parts)
    root,styles,settings=xml(parts[DOC]),xml(parts[STYLES]),xml(parts['word/settings.xml'])
    if any(list(root.iter(q(tag))) for tag in REVISIONS) or any(k.startswith('_xmlsignatures/') for k in parts):
        raise ValueError('修订或数字签名需要单独处理')
    if any(settings.find(q(tag)) is not None for tag in ('documentProtection','trackRevisions')):
        raise ValueError('文档保护或修订已开启')
    body=root.find(q('body'));nodes=list(body)
    source_sections=list(body.iter(q('sectPr')))
    if len(source_sections)!=4:
        raise ValueError('此审核模板入口要求封面、摘要、正文首节、正文后续节共四节，请先调整适配器')
    if nodes[manifest['cover_end']].find('./'+q('pPr')+'/'+q('sectPr')) is not source_sections[0]:
        raise ValueError('封面边界与首节不一致')
    before_text=[visible(n) for n in nodes]
    before_cover=[canonical(n) for n in nodes[:manifest['cover_end']+1]]
    for index in manifest.get('exclude_from_outline', []):
        if index > manifest['cover_end']:
            raise ValueError('仅允许明确排除封面大纲')
        pp=properties(nodes[index])
        for n in pp.findall(q('outlineLvl')):pp.remove(n)
        pp.append(element('outlineLvl',val=9))
    before_math=math_snapshot(root)
    before_cells=[[visible(c) for c in t.iter(q('tc'))] for t in root.iter(q('tbl'))]
    for item in manifest['paragraphs']:
        if visible(nodes[item['index']])!=item['text']:
            raise ValueError('段落位置或文本与审核清单不一致')
    selected={nodes[i['index']] for i in manifest['paragraphs']}
    controlled=set(manifest['styles'])
    parents={s.get(q('styleId')):val(s,'basedOn') for s in styles.findall(q('style'))}
    for part,data in parts.items():
        if not part.startswith('word/') or not part.endswith('.xml') or part==STYLES:continue
        tree=root if part==DOC else xml(data)
        for p in tree.iter(q('p')):
            if part==DOC and p in selected:continue
            sid=val(p,'pPr/pStyle');seen=set()
            while sid and sid not in seen:
                if sid in controlled:raise ValueError('共享样式影响非目标段落 '+str(sid))
                seen.add(sid);sid=parents.get(sid)
    for sid,spec in manifest['styles'].items():
        style_definition(styles,sid,spec['name'],spec['rule'],spec.get('outline',9))
    changes=[]
    for item in manifest['paragraphs']:
        p=nodes[item['index']]
        if item.get('remove_suffix'):remove_suffix(p,item['remove_suffix'])
        apply_style(p,item['style'])
        changes.append({'index':item['index'],'before_style':val(xml(original[DOC]).find(q('body'))[item['index']],'pPr/pStyle'), 'after_style':item['style'],'removed_instruction':item.get('remove_suffix')})
    captions=[]
    for item in manifest['captions']:
        obj=nodes[item['index']]
        if 'caption_index' in item:
            p=nodes[item['caption_index']]
            import re
            from number_captions import insert_seq
            match=re.match(r'^'+item['label']+r'\s*(\d+)',visible(p))
            if match is None or int(match[1])!=item['number'] or scan(p)[0]:
                raise ValueError('原有手打题注与审核编号不一致')
            # A caption can carry the end of its original section. Moving that
            # paragraph above a table must not move the section boundary too.
            section=p.find('./'+q('pPr')+'/'+q('sectPr'))
            if section is not None:
                boundary=element('p');pp=element('pPr')
                pp.append(element('spacing',before=0,after=0,line=1,lineRule='exact'))
                pp.append(section);boundary.append(pp);p.addnext(boundary)
                changes.append({'index':item['caption_index'],'change':'题注移动，原分节位置保留'})
            insert_seq(p,*match.span(1),match[1],f' SEQ {item["label"]} \\* ARABIC ')
        else:
            p=element('p');p.append(run(item['label']));p.append(seq_field(f'SEQ {item["label"]} \\* ARABIC',item['number']))
        (obj.addnext if item['label']=='图' else obj.addprevious)(p)
        captions.append(p)
    sid=configure(root,styles,settings,[body.index(p) for p in captions])
    style_definition(styles,sid,'caption',manifest['caption_rule'])
    for p in captions:apply_style(p,sid)
    for c in settings.find(q('captions')):
        if c.get(q('name'))=='表':c.set(q('pos'),'above')
    for item in manifest['tables']:
        t=nodes[item['index']];three_line_table(t,styles,item['header_rows'])
        for p in t.iter(q('p')):apply_style(p,'PaperFitTableText')
    for item in manifest.get('image_scales', []):
        scale=item['scale']
        if not 0 < scale <= 1:raise ValueError('图片只支持等比例缩小')
        drawing=nodes[item['index']].find('.//'+q('drawing'))
        ext=drawing.find('.//{'+WP+'}extent')
        local=drawing.find('.//{'+PIC+'}spPr/{'+A+'}xfrm/{'+A+'}ext')
        if ext is None or local is None:raise ValueError('图片尺寸结构不支持')
        for n in (ext,local):
            for key in ('cx','cy'):n.set(key,str(round(int(n.get(key))*scale)))
    # Existing sections are explicitly reviewed in the manifest. Cover is retained.
    secs=list(body.iter(q('sectPr')))
    for s in secs[1:]:
        margin=s.find(q('pgMar'));margin.set(q('bottom'),'1440');margin.set(q('footer'),'720')
        for tag in ('titlePg','footerReference','pgNumType'):
            for n in s.findall(q(tag)):s.remove(n)
        for k,v in add_footer(parts,s).items():parts[k]=E.tostring(v,xml_declaration=True,encoding='UTF-8')
        num=element('pgNumType',fmt='upperRoman' if s is secs[1] else 'decimal')
        if s in secs[1:3]:num.set(q('start'),'1')
        s.insert(next((i for i,n in enumerate(s) if n.tag in {q('cols'),q('docGrid')}),len(s)),num)
    first=nodes[manifest['body_start']]
    section=secs[2];margin=section.find(q('pgMar'));width=int(section.find(q('pgSz')).get(q('w')))-sum(int(margin.get(q(k))) for k in ('left','right'))
    toc=insert_toc(root,styles,first,width)
    front=deepcopy(section)
    for tag in ('footerReference','pgNumType'):
        for n in front.findall(q(tag)):front.remove(n)
    # Explicit blank footer prevents inheritance from the abstract.
    added=add_footer(parts,front)
    for k,v in added.items():parts[k]=E.tostring(v,xml_declaration=True,encoding='UTF-8')
    rid=front.find(q('footerReference')).get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    target=next(n.get('Target') for n in xml(parts[RELS]) if n.get('Id')==rid)
    empty=element('ftr');empty.append(element('p'));parts['word/'+target]=E.tostring(empty,xml_declaration=True,encoding='UTF-8')
    boundary=element('p');pp=element('pPr');pp.extend([element('spacing',before=0,after=0,line=1,lineRule='exact'),front]);boundary.append(pp);toc.addnext(boundary)
    disable_picture_compression(settings)
    with TemporaryDirectory(dir=output.parent) as tmp:
        staged=Path(tmp)/'staged.docx'
        save_package(parts,{DOC:root,STYLES:styles,'word/settings.xml':settings},staged)
        matches=match_captions(staged)
        ignored=[i for i in matches['unmatched_objects'] if body.index(first)>i]
        matches['unmatched_objects']=[i for i in matches['unmatched_objects'] if i not in ignored]
        caption_changes,pending=bind_captions(root,styles,matches)
        _,heading_changes=split_chapters(root,styles);keep_subheadings(root,styles)
        trees,story_changes=independent_stories(parts,root)
        # Normalize property ordering after composition from multiple helpers.
        for pp in root.iter(q('pPr')):
            pp[:]=sorted(pp,key=lambda n:P_ORDER.index(E.QName(n).localname) if E.QName(n).localname in P_ORDER else 999)
        for s in styles.findall(q('style')):
            pp=s.find(q('pPr'))
            if pp is not None:pp[:]=sorted(pp,key=lambda n:P_ORDER.index(E.QName(n).localname) if E.QName(n).localname in P_ORDER else 999)
        expected=before_text[:]
        for item in manifest['paragraphs']:
            if item.get('remove_suffix'):expected[item['index']]=expected[item['index']][:-len(item['remove_suffix'])]
        if [visible(n) for n in nodes]!=expected:raise ValueError('原始文本完整性校验失败')
        if before_cover!=[canonical(n) for n in nodes[:manifest['cover_end']+1]]:
            # Header/footer references may be duplicated for independent stories.
            a=[deepcopy(n) for n in xml(original[DOC]).find(q('body'))[:manifest['cover_end']+1]]
            b=[deepcopy(n) for n in nodes[:manifest['cover_end']+1]]
            for group in (a,b):
                for index,n in enumerate(group):
                    if index in manifest.get('exclude_from_outline',[]):
                        for x in list(n.iter(q('outlineLvl'))):x.getparent().remove(x)
                    for tag in ('headerReference','footerReference'):
                        for x in list(n.iter(q(tag))):x.getparent().remove(x)
            if [canonical(n) for n in a]!=[canonical(n) for n in b]:raise ValueError('封面发生未授权变化')
        if before_math!=math_snapshot(root):raise ValueError('公式结构发生变化')
        if before_cells!=[[visible(c) for c in t.iter(q('tc'))] for t in root.iter(q('tbl'))]:raise ValueError('表格内容变化')
        if scan(root)[1]:raise ValueError('域结构错误')
        trees.update({DOC:root,STYLES:styles,'word/settings.xml':settings})
        final=Path(tmp)/'final.docx';save_package(parts,trees,final)
        native_report=None
        if manifest.get('native_references',True):
            from native_references import convert_parts
            native_parts,native_report=convert_parts(read_package(final))
            final=Path(tmp)/'native.docx';save_package(native_parts,{},final)
            pending.extend(native_report['pending'])
        finalparts=read_package(final)
        media=[k for k in original if k.startswith('word/media/')]
        if any(original[k]!=finalparts[k] for k in media):raise ValueError('图片数据变化')
        report={'source_sha256':manifest['source_sha256'],'output_sha256':sha256(final.read_bytes()).hexdigest(),
                'paragraph_changes':changes,'captions':caption_changes,'headings':heading_changes,'sections':len(list(root.iter(q('sectPr')))),
                'integrity':{'original_text_except_reviewed_instructions':True,'table_cells':True,'math':True,'media_bytes':True,'media_count':len(media),'cover_preserved_except_independent_story_references':True},
                'excluded_cover_objects':ignored,'pending':pending,'assumptions':manifest.get('assumptions',[]),
                'layout_check':'必须在 Word/WPS 更新目录并检查全页'}
        report['native_references']=native_report
        record=Path(tmp)/'report.json';record.write_text(json.dumps(report,ensure_ascii=False,indent=2))
        os.link(final,output)
        try:os.link(record,report_path)
        except BaseException:output.unlink();raise
    return report


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source');parser.add_argument('manifest');parser.add_argument('output')
    args=parser.parse_args()
    format_template(args.source,json.loads(Path(args.manifest).read_text()),args.output)

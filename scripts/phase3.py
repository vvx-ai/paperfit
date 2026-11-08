"""Local chapter edits and explicitly mapped sections of Word templates."""
import argparse
from hashlib import sha256
import json
import math
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import BadZipFile

from lxml import etree as ET
from paperfit import q, W
from phase2 import DOC, STYLES, canonical, read_package, save_package, xml
from structured import element, visible
from toc import heading_level, val
from word_fields import scan
from toc_repair import managed_links, repair_links

FACES = ('ascii', 'hAnsi', 'eastAsia', 'cs')
LAYOUT = {'pgSz': ('w', 'h', 'orient'), 'pgMar': ('left', 'right', 'top', 'bottom')}
PARAGRAPH = ('keepNext','keepLines','pageBreakBefore','widowControl','snapToGrid',
             'spacing','ind','jc')
REVISIONS = ('ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange', 'sectPrChange')


def definitions(styles):
    result = {}
    for style in styles.findall(q('style')):
        sid = style.get(q('styleId'))
        if not sid or sid in result:
            raise ValueError('样式 ID 缺失或重复，需人工核对')
        result[sid] = style
    return result


def style_chain(sid, styles):
    seen = set()
    while sid:
        if sid in seen or sid not in styles:
            raise ValueError('样式继承循环或基样式缺失')
        seen.add(sid)
        style = styles[sid]
        yield style
        sid = val(style, 'basedOn')


def typography(layers):
    """Report literal or theme fonts and sizes, least specific layer first."""
    fonts, sizes = {}, {}
    for props in layers:
        if props is None:
            continue
        names = props.find(q('rFonts'))
        if names is not None:
            for face in FACES:
                theme = names.get(q('cstheme' if face == 'cs' else face+'Theme'))
                literal = names.get(q(face))
                if theme or literal:
                    fonts[face] = {'theme': theme} if theme else literal
        for key in ('sz', 'szCs'):
            size = val(props, key)
            if size is not None:
                sizes[key] = int(size)
    return {'fonts': fonts, 'half_points': sizes}


def layers_for(style_root, styles, sid):
    defaults = style_root.find('./'+q('docDefaults')+'/'+q('rPrDefault')+'/'+q('rPr'))
    return [defaults] + [s.find(q('rPr')) for s in reversed(list(style_chain(sid, styles)))]


def paragraph_values(style_root, styles, sid):
    default = style_root.find('./'+q('docDefaults')+'/'+q('pPrDefault')+'/'+q('pPr'))
    layers = [default] + [s.find(q('pPr')) for s in reversed(list(style_chain(sid,styles)))]
    result = {}
    for props in layers:
        if props is None:
            continue
        for tag in PARAGRAPH:
            node = props.find(q(tag))
            if node is not None:
                attrs = {ET.QName(k).localname:v for k,v in node.attrib.items()}
                if not attrs and tag not in ('spacing','ind','jc'):
                    attrs = {'val':'1'}
                result.setdefault(tag,{}).update(attrs)
    return result


def template_type(parts, name):
    root = xml(parts[STYLES])
    styles = definitions(root)
    matches = [sid for sid,s in styles.items()
               if s.get(q('type')) == 'paragraph' and val(s,'name') == name]
    if len(matches) != 1:
        raise ValueError('模板段落样式名称缺失或不唯一')
    values = typography(layers_for(root, styles, matches[0]))
    fonts, sizes = values['fonts'], values['half_points']
    if (any(face not in fonts or not isinstance(fonts[face], str) for face in ('ascii','hAnsi','eastAsia'))
            or any(not isinstance(v, str) for v in fonts.values()) or not sizes.get('sz')):
        raise ValueError('模板没有给出明确的中西文字体和字号，主题字体需先确认')
    if any(n <= 0 or n > 3276 for n in sizes.values()):
        raise ValueError('模板字号超出支持范围')
    values['paragraph'] = paragraph_values(root,styles,matches[0])
    return values


def section_values(section):
    result = {}
    for tag, attrs in LAYOUT.items():
        node = section.find(q(tag))
        if node is None:
            raise ValueError('节缺少纸张尺寸或页边距')
        result[tag] = {a:node.get(q(a)) for a in attrs}
    return result


def sections(body):
    found = list(body.iter(q('sectPr')))
    if not found or found[-1] is not body[-1]:
        raise ValueError('缺少文档结束分节设置')
    for section in found[:-1]:
        parent = section.getparent()
        if parent.tag != q('pPr') or parent.getparent().getparent() is not body:
            raise ValueError('分节位于嵌套结构或修订中，请先复核')
    return found


def copy_layout(root, template, target_no, source_no):
    body = root.find(q('body'))
    targets = sections(body)
    sources = sections(xml(template[DOC]).find(q('body')))
    if not 1 <= target_no <= len(targets) or not 1 <= source_no <= len(sources):
        raise ValueError('请指定文档中存在的节序号，从 1 开始')
    target = targets[target_no-1]
    before, after = section_values(target), section_values(sources[source_no-1])
    for tag, attrs in after.items():
        node = target.find(q(tag))
        for key, value in attrs.items():
            if key == 'orient':
                if value not in (None,'portrait','landscape'):
                    raise ValueError('模板纸张方向无效')
            elif value is None or int(value) < 0 or (tag == 'pgSz' and int(value) == 0):
                raise ValueError('模板纸张或边距值需复核')
            if value is None:
                node.attrib.pop(q(key), None)
            else:
                node.set(q(key), value)
    if int(after['pgSz']['w']) <= sum(int(after['pgMar'][k]) for k in ('left','right')) or int(after['pgSz']['h']) <= sum(int(after['pgMar'][k]) for k in ('top','bottom')):
        raise ValueError('模板页边距超出纸张')
    return {'section':target_no,'template_section':source_no,'before':before,'after':after}


def chapter_targets(body, styles, chapter, level):
    outline = []
    for i,node in enumerate(body):
        if node.tag == q('p'):
            tier = heading_level(node, styles)
            if tier:
                outline.append((i,node,tier))
        elif val(node,'sdtPr/tag') != 'PaperFitTOC' and any(heading_level(p, styles) for p in node.iter(q('p'))):
            raise ValueError('存在嵌套标题，请先核对章节边界')
    starts = [i for i,p,tier in outline if tier == 1 and visible(p).strip() == chapter]
    if len(starts) != 1:
        raise ValueError('一级标题文字缺失或重复，请用原文唯一定位章节')
    start = starts[0]
    end = next((i for i,p,tier in outline if i > start and tier == 1),len(body))
    result = [(i,p) for i,p,tier in outline if start <= i < end and tier == level]
    if not result:
        raise ValueError('指定章节内没有该级标题')
    return result


def variant(base, chapter, level, values):
    key = json.dumps([base.get(q('styleId')),chapter,level,values], ensure_ascii=False,sort_keys=True)
    sid = 'PaperFitLocal'+sha256(key.encode()).hexdigest()[:20]
    style = element('style', type='paragraph', customStyle=1, styleId=sid)
    label = ' · '.join([chapter,val(base,'name',base.get(q('styleId'))),
                       str(values['fonts'].get('eastAsia','')),f"{values['half_points']['sz']/2:g}磅"])
    style.extend([element('name',val=label),element('basedOn',val=base.get(q('styleId')))])
    next_style = val(base,'next')
    if next_style:
        style.append(element('next',val=next_style))
    style.append(element('qFormat'))
    pp, rp = element('pPr'), element('rPr')
    for tag in PARAGRAPH:
        if tag in values.get('paragraph',{}):
            pp.append(element(tag,**values['paragraph'][tag]))
    pp.append(element('outlineLvl',val=level-1))
    rp.append(element('rFonts',**values['fonts']))
    for k,n in values['half_points'].items():
        rp.append(element(k,val=n))
    style.extend([pp,rp])
    return style


def apply_chapter(root, style_root, chapter, level, values, parts):
    styles = definitions(style_root)
    body = root.find(q('body'))
    targets = chapter_targets(body,styles,chapter,level)
    fields, errors = scan(body)
    if errors or any(f.code.upper().split()[:1] == ['TOC'] and '\\u' not in f.code.lower().split() for f in fields):
        raise ValueError('目录按特定样式收录或域未闭合，请人工核对')
    linked = set()
    if 'word/numbering.xml' in parts:
        linked = {p.get(q('val')) for p in xml(parts['word/numbering.xml']).iter(q('pStyle'))}
    changes = []
    for index,p in targets:
        sid = val(p,'pPr/pStyle')
        if sid not in styles or styles[sid].get(q('type')) != 'paragraph':
            raise ValueError('目标标题缺少明确的段落样式')
        if any(s.get(q('styleId')) in linked for s in style_chain(sid,styles)):
            raise ValueError('目标基样式与多级列表直接绑定，需要核对编号继承')
        if any(n.tag in {q(k) for k in ('fldSimple','fldChar','drawing','object','pict','rStyle','hyperlink')} or ET.QName(n).namespace != W for n in p.iter() if isinstance(n.tag,str)):
            raise ValueError('目标标题含复杂对象、域或字符样式，请先复核')
        layers = layers_for(style_root,styles,sid)
        effective = typography(layers)
        effective['paragraph'] = paragraph_values(style_root,styles,sid)
        # Word may rename style IDs and omit properties identical to the base on save.
        already_styled = all(effective[group].get(key) == value
                             for group in ('fonts','half_points') for key,value in values[group].items())
        already_styled = already_styled and all(
            effective['paragraph'].get(tag,{}).get(key) == value
            for tag,attrs in values.get('paragraph',{}).items() for key,value in attrs.items())
        requested_indent = values.get('paragraph',{}).get('ind',{})
        inherited_indent = effective['paragraph'].get('ind',{})
        direct_indent = p.find('./'+q('pPr')+'/'+q('ind'))
        if direct_indent is not None:
            inherited_indent = dict(inherited_indent, **{ET.QName(k).localname:v for k,v in direct_indent.attrib.items()})
        for requested, conflicting in ((('firstLine','firstLineChars'),('hanging','hangingChars')),
                                       (('hanging','hangingChars'),('firstLine','firstLineChars'))):
            if any(int(requested_indent.get(k,'0')) for k in requested) and any(int(inherited_indent.get(k,'0')) for k in conflicting):
                raise ValueError('模板缩进与既有首行或悬挂缩进冲突，需要核对')
        if already_styled:
            new_id = sid
            candidate = styles[sid]
        else:
            candidate = variant(styles[sid],chapter,level,values)
            new_id = candidate.get(q('styleId'))
            if new_id in styles:
                if canonical(styles[new_id]) != canonical(candidate):
                    raise ValueError('已有局部样式经过编辑，请先复核')
            else:
                if any(val(s,'name') == val(candidate,'name') for s in styles.values()):
                    raise ValueError('新样式名称与已有样式冲突')
                style_root.append(candidate)
                styles[new_id] = candidate
        before = [typography(layers+[r.find(q('rPr'))]) for r in p.iter(q('r')) if r.find(q('t')) is not None]
        before = [item for j,item in enumerate(before) if item not in before[:j]]
        ref = p.find('./'+q('pPr')+'/'+q('pStyle'))
        ref.set(q('val'),new_id)
        removed = []
        pp = p.find(q('pPr'))
        for tag,attrs in values.get('paragraph',{}).items():
            node = pp.find(q(tag))
            if node is not None:
                for key in attrs:
                    node.attrib.pop(q(key),None)
                if not node.attrib:
                    pp.remove(node)
                removed.append(tag)
        for rp in p.iter(q('rPr')):
            fonts = rp.find(q('rFonts'))
            if fonts is not None:
                for face in values['fonts']:
                    for key in (face, 'cstheme' if face=='cs' else face+'Theme'):
                        if q(key) in fonts.attrib:
                            removed.append(key)
                            del fonts.attrib[q(key)]
                if not fonts.attrib:
                    rp.remove(fonts)
            for tag in values['half_points']:
                old = rp.find(q(tag))
                if old is not None:
                    removed.append(tag)
                    rp.remove(old)
        changes.append({'body_index':index,'heading':visible(p),'before_style':sid,
                        'after_style':new_id,'style_name':val(candidate,'name'),
                        'before_format':before,'requested_format':values,
                        'cleared_direct_properties':sorted(set(removed))})
    return changes


def publish(source, output, operation):
    source, output = Path(source), Path(output)
    report_path = Path(str(output)+'.report.json')
    if output.suffix.lower() != '.docx' or output.exists() or report_path.exists():
        raise ValueError('请选择新的 DOCX 和报告路径')
    parts = read_package(source)
    root = xml(parts[DOC])
    if any(True for n in REVISIONS for _ in root.iter(q(n))) or any(k.startswith('_xmlsignatures/') for k in parts):
        raise ValueError('检测到修订或数字签名，请核对后处理')
    if 'word/settings.xml' in parts:
        settings = xml(parts['word/settings.xml'])
        if any(settings.find(q(tag)) is not None for tag in ('documentProtection','trackRevisions')):
            raise ValueError('文档启用保护或修订，请先复核')
    styles = xml(parts[STYLES])
    report = operation(root,styles,parts)
    with TemporaryDirectory(dir=output.parent) as tmp:
        doc = Path(tmp)/'result.docx'
        save_package(parts,{DOC:root, STYLES:styles},doc)
        result = read_package(doc)
        changed = [k for k in parts if parts[k] != result[k]]
        before = xml(parts[DOC])
        # Compare source and output after removing just the explicitly editable properties.
        after = xml(result[DOC])
        if 'changes' in report:
            for item in report['changes']:
                idx = item['body_index']
                for tree in (before,after):
                    p = tree.find(q('body'))[idx]
                    p.find('./'+q('pPr')+'/'+q('pStyle')).attrib.pop(q('val'))
                    for tag,attrs in item['requested_format'].get('paragraph',{}).items():
                        node = p.find('./'+q('pPr')+'/'+q(tag))
                        if node is not None:
                            for key in attrs:
                                node.attrib.pop(q(key),None)
                            if not node.attrib:
                                node.getparent().remove(node)
                    for rp in p.iter(q('rPr')):
                        for k in item['requested_format']['half_points']:
                            node = rp.find(q(k))
                            if node is not None:
                                rp.remove(node)
                        fonts = rp.find(q('rFonts'))
                        if fonts is not None:
                            for face in item['requested_format']['fonts']:
                                fonts.attrib.pop(q(face),None)
                                fonts.attrib.pop(q('cstheme' if face=='cs' else face+'Theme'),None)
                            if not fonts.attrib:
                                rp.remove(fonts)
        elif 'repairs' in report:
            for tree in (before,after):
                links = managed_links(tree)
                for repair in report['repairs']:
                    links[repair['link']].attrib.pop(q('anchor'),None)
        else:
            for tree in (before,after):
                sec = sections(tree.find(q('body')))[report['layout']['section']-1]
                for tag,attrs in LAYOUT.items():
                    for key in attrs:
                        sec.find(q(tag)).attrib.pop(q(key),None)
        if canonical(before) != canonical(after):
            raise ValueError('范围外结构发生变化，取消输出')
        old_styles = xml(parts[STYLES])
        new_styles = xml(result[STYLES])
        for sid in {c['after_style'] for c in report.get('changes',[])} - set(definitions(old_styles)):
            new_styles.remove(definitions(new_styles)[sid])
        if canonical(old_styles) != canonical(new_styles) or set(changed)-{DOC, STYLES}:
            raise ValueError('已有样式或其他部件发生变化')
        report.update(source_sha256=sha256(source.read_bytes()).hexdigest(),
                      output_sha256=sha256(doc.read_bytes()).hexdigest(),
                      changed_parts=changed,scope_integrity=True)
        record = Path(tmp)/'report.json'
        record.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        os.link(record,report_path)
        try:
            os.link(doc,output)
        except BaseException:
            report_path.unlink()
            raise
    return report


def template_inventory(parts):
    styles = xml(parts[STYLES])
    defs = definitions(styles)
    records = []
    for sid, style in defs.items():
        if style.get(q('type')) == 'paragraph':
            records.append({'id':sid,'name':val(style,'name'),'base':val(style,'basedOn'),
                            'typography':typography(layers_for(styles,defs,sid))})
    body = xml(parts[DOC]).find(q('body'))
    layout = [{'number':i,'layout':section_values(sec),'page_start':sec.find(q('pgNumType')).get(q('start')) if sec.find(q('pgNumType')) is not None else None,
               'header_footer_references':sum(child.tag in (q('headerReference'),q('footerReference')) for child in sec)}
              for i,sec in enumerate(sections(body),1)]
    return {'styles':records,'sections':layout}


def main(argv):
    parser = argparse.ArgumentParser(description='PaperFit 局部样式与分节模板')
    subs = parser.add_subparsers(dest='cmd',required=True)
    repair = subs.add_parser('repair-toc')
    repair.add_argument('source',type=Path)
    repair.add_argument('output',type=Path)
    audit = subs.add_parser('template-info')
    audit.add_argument('source',type=Path)
    audit.add_argument('output',type=Path)
    diff = subs.add_parser('compare')
    diff.add_argument('before',type=Path)
    diff.add_argument('after',type=Path)
    diff.add_argument('output',type=Path)
    sub = subs.add_parser('chapter')
    sub.add_argument('--chapter',required=True,help='一级标题原文')
    sub.add_argument('--level',choices=(1,2,3),type=int,required=True)
    sub.add_argument('--font')
    sub.add_argument('--size',type=float)
    sub.add_argument('--template',type=Path)
    sub.add_argument('--style',help='模板段落样式的名称')
    layout = subs.add_parser('layout')
    layout.add_argument('--template',required=True,type=Path)
    layout.add_argument('--template-section',required=True,type=int)
    layout.add_argument('--section',required=True,type=int)
    for command in (sub, layout):
        command.add_argument('source',type=Path)
        command.add_argument('output',type=Path)
    args = parser.parse_args(argv)
    try:
        if args.cmd == 'repair-toc':
            publish(args.source,args.output,lambda root,styles,parts:repair_links(root,definitions(styles)))
            return
        if args.cmd in ('template-info','compare'):
            if args.cmd == 'compare':
                from docx_diff import compare
                result = compare(args.before,args.after)
            else:
                result = template_inventory(read_package(args.source))
            with args.output.open('x',encoding='utf-8') as handle:
                json.dump(result,handle,ensure_ascii=False,indent=2)
            return
        template = read_package(args.template) if args.template else None
        if args.template and args.output.resolve()==args.template.resolve():
            raise ValueError('输出不能覆盖模板')
        provenance = {'template_sha256':sha256(args.template.read_bytes()).hexdigest()} if template else {}
        if args.cmd == 'chapter':
            if template is not None:
                if not args.style or args.font is not None or args.size is not None:
                    raise ValueError('模板模式仅接受明确的 --style')
                values = template_type(template,args.style)
                provenance['template_style'] = args.style
            else:
                if args.style or not args.font or not args.size or not math.isfinite(args.size) or not 1 <= args.size <= 1638 or not (args.size*2).is_integer():
                    raise ValueError('请指定字体与以半磅为单位的有效字号')
                values = {'fonts':dict.fromkeys(FACES,args.font),'half_points':dict.fromkeys(('sz','szCs'),int(args.size*2))}
            def operation(root,styles,parts):
                return dict(provenance, chapter=args.chapter,level=args.level,
                            changes=apply_chapter(root,styles,args.chapter,args.level,values,parts))
        else:
            def operation(root,styles,parts):
                return dict(provenance,layout=copy_layout(root,template,args.section,args.template_section))
        publish(args.source,args.output,operation)
    except (ValueError,OSError,KeyError,BadZipFile,ET.XMLSyntaxError) as error:
        parser.exit(2,f'PaperFit: {error}\n')

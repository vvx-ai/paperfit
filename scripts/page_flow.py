"""Keep captions with their objects and start level-one headings in new sections."""
from copy import deepcopy
from hashlib import sha256
import json
import math
import os
from pathlib import Path
import posixpath
from tempfile import TemporaryDirectory

from lxml import etree as ET
from paperfit import q, math_snapshot
from phase2 import DOC, STYLES, canonical, read_package, save_package, xml
from phase3 import definitions, paragraph_values, layers_for, typography, sections, REVISIONS
from pagination import REL, RELS, TYPES, section_property
from match_captions import match_captions
from structured import element, visible
from toc import heading_level, val

WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
FLAGS = ('keepNext', 'keepLines', 'pageBreakBefore')
P_ORDER = ('pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl',
           'numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens',
           'kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN',
           'bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing',
           'mirrorIndents','suppressOverlap','jc','textDirection','textAlignment',
           'textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange')


def properties(paragraph):
    pp = paragraph.find(q('pPr'))
    if pp is None:
        pp = element('pPr')
        paragraph.insert(0, pp)
    return pp


def flag(pp, tag, enabled=True):
    node = pp.find(q(tag))
    if node is None:
        node = element(tag)
        rank = P_ORDER.index(tag)
        index = next((i for i,n in enumerate(pp)
                      if ET.QName(n).localname in P_ORDER and P_ORDER.index(ET.QName(n).localname) > rank), len(pp))
        pp.insert(index, node)
    node.set(q('val'), '1' if enabled else '0')


def section_for(body, paragraph):
    index = body.index(paragraph)
    for node in list(body)[index:]:
        found = node if node.tag == q('sectPr') else node.find('./'+q('pPr')+'/'+q('sectPr'))
        if found is not None:
            return found
    raise ValueError('缺少分节设置')


def page_area(section):
    size, margin = section.find(q('pgSz')), section.find(q('pgMar'))
    if size is None or margin is None:
        raise ValueError('缺少纸张与页边距设置')
    width = int(size.get(q('w'))) - sum(int(margin.get(q(k),'0')) for k in ('left','right','gutter'))
    height = int(size.get(q('h'))) - sum(int(margin.get(q(k),'0')) for k in ('top','bottom'))
    if width <= 0 or height <= 0:
        raise ValueError('页面可用区域无效')
    return width, height


def reserve_height(p, style_root, styles, width):
    """Conservative caption allowance in twips; final pagination needs Word/WPS."""
    sid = val(p,'pPr/pStyle')
    sizes = typography(layers_for(style_root, styles, sid))['half_points']
    font = sizes.get('sz',24) * 10
    for rp in p.iter(q('rPr')):
        font = max(font, int(val(rp,'sz',str(font//10)))*10)
    attrs = paragraph_values(style_root,styles,sid)
    pp = p.find(q('pPr'))
    if pp is not None:
        for node in pp:
            name = ET.QName(node).localname
            attrs.setdefault(name,{}).update({ET.QName(k).localname:v for k,v in node.attrib.items()})
    spacing = attrs.get('spacing',{})
    lines = max(1, math.ceil(len(visible(p))*font/max(width,1))) + len(list(p.iter(q('br'))))
    line = max(font*1.5, int(spacing.get('line','240'))*font/240 if spacing.get('lineRule','auto')=='auto' else int(spacing.get('line','0')))
    return lines*line + sum(int(spacing.get(k,'0')) for k in ('before','after')) + 480


def bind_captions(root, style_root, matches):
    body, styles = root.find(q('body')), definitions(style_root)
    changes, pending = [], []
    for item in matches['matches']:
        if item['status'] != 'unique_adjacent_match':
            pending.append({'caption':item['text'],'reason':'题注与对象未能唯一匹配'})
            continue
        caption, obj = body[item['caption_body_index']], body[item['object_body_index']]
        section=section_for(body,obj)
        columns=section.find(q('cols'))
        if columns is not None and (columns.get(q('num'),'1')!='1' or len(columns.findall(q('col')))>1):
            pending.append({'caption':item['text'],'reason':'多栏版面需要核对对象所在栏宽'})
            continue
        if (list(caption.iter(q('br'))) or list(caption.iter(q('sectPr')))
                or section_for(body,caption) is not section_for(body,obj)):
            pending.append({'caption':item['text'],'reason':'题注含显式换行或与对象跨节，需复核'})
            continue
        pp = properties(caption)
        dimensions = None
        if item['kind'] == '图':
            drawings = obj.findall('.//'+q('drawing'))
            inline = drawings[0].find('{'+WP+'}inline') if len(drawings)==1 else None
            if inline is None or visible(obj).strip() or list(obj.iter(q('br'))):
                pending.append({'caption':item['text'],'reason':'浮动或混合内容图片需要单独处理'})
                continue
            extent = inline.find('{'+WP+'}extent')
            transform = inline.find('.//{'+PIC+'}spPr/{'+A+'}xfrm')
            rotation = int(transform.get('rot','0')) % 21600000 if transform is not None else 0
            if extent is None or transform is None or rotation not in (0,5400000,10800000,16200000):
                pending.append({'caption':item['text'],'reason':'图片尺寸或旋转结构不支持安全缩放'})
                continue
            local_extent = transform.find('{'+A+'}ext')
            if local_extent is None:
                pending.append({'caption':item['text'],'reason':'图片缺少形状尺寸'})
                continue
            width, height = page_area(section_for(body,obj))
            max_height = height - reserve_height(caption,style_root,styles,width) - 480
            old = {k:int(extent.get(k)) for k in ('cx','cy')}
            if min(old.values())<=0 or max_height<=0:
                raise ValueError('图片或题注可用尺寸无效')
            rotated = rotation in (5400000,16200000)
            display_width,display_height = (old['cy'],old['cx']) if rotated else (old['cx'],old['cy'])
            scale = min(1, width*635/display_width, max_height*635/display_height)
            if scale < 1:
                for target in (extent,local_extent):
                    for axis in ('cx','cy'):
                        target.set(axis,str(max(1,int(int(target.get(axis))*scale))))
                dimensions = {'before_emu':old,'after_emu':{k:int(extent.get(k)) for k in old}}
            image_pp = properties(obj)
            flag(image_pp,'keepLines')
            flag(image_pp,'keepNext',item['position']=='after')
            if item['position']=='before':
                flag(image_pp,'pageBreakBefore',False)
        else:
            if obj.find('./'+q('tblPr')+'/'+q('tblpPr')) is not None or len(list(obj.iter(q('tbl'))))!=1:
                pending.append({'caption':item['text'],'reason':'浮动或嵌套表格需要单独处理'})
                continue
            if list(obj.iter(q('br'))) or list(obj.iter(q('sectPr'))):
                pending.append({'caption':item['text'],'reason':'表格含显式换行或分节，需复核'})
                continue
            rows = obj.findall(q('tr'))
            minimum_height=0
            for row in rows:
                rp = row.find(q('trPr'))
                if rp is None:
                    rp = element('trPr'); row.insert(0,rp)
                height=rp.find(q('trHeight'))
                minimum_height+=max(360,int(height.get(q('val'),'0')) if height is not None else 0)
                split = rp.find(q('cantSplit'))
                if split is None:
                    split = element('cantSplit')
                    successors={q(n) for n in ('trHeight','tblHeader','tblCellSpacing','jc','hidden','ins','del','trPrChange')}
                    index=next((i for i,n in enumerate(rp) if n.tag in successors),len(rp))
                    rp.insert(index,split)
                split.set(q('val'),'1')
                for cell in row.findall(q('tc')):
                    paras = cell.findall(q('p'))
                    for p in paras:
                        props = properties(p)
                        flag(props,'keepLines')
                        flag(props,'keepNext',item['position']=='after' or row is not rows[-1] or p is not paras[-1])
                        flag(props,'pageBreakBefore',False)
            width,height=page_area(section)
            if minimum_height+reserve_height(caption,style_root,styles,width)>height:
                pending.append({'caption':item['text'],'reason':'表格高度估算超过单页，需调整布局或采用跨页续表'})
            # Word can split an oversized table despite keepNext. Require rendered verification.
        flag(pp,'keepLines')
        flag(pp,'keepNext',item['position']=='before')
        if item['position']=='after':
            flag(pp,'pageBreakBefore',False)
        changes.append({'caption':item['text'],'kind':item['kind'],'position':item['position'],
                        'image_resize':dimensions,'verification':'需要排版应用核对同页；超长表格需另行处理'})
    for index in matches['unmatched_objects']:
        pending.append({'body_index':index,'reason':'对象缺少唯一相邻题注'})
    return changes,pending


def split_chapters(root, style_root):
    body, styles = root.find(q('body')), definitions(style_root)
    sections(body)
    headings = []
    for node in body:
        if node.tag == q('p') and heading_level(node,styles)==1:
            headings.append(node)
        elif node.tag != q('p') and val(node,'sdtPr/tag')!='PaperFitTOC' and any(heading_level(p,styles)==1 for p in node.iter(q('p'))):
            raise ValueError('嵌套一级标题需要核对分节边界')
    inserted, changes = [], []
    # Process backwards so each split retains the original section's restart at its first segment.
    for heading in reversed(headings):
        index = body.index(heading)
        current = section_for(body,heading)
        previous = heading.getprevious()
        existing = previous is not None and previous.find('./'+q('pPr')+'/'+q('sectPr')) is not None
        if index and not existing:
            previous_section = deepcopy(current)
            boundary = element('p')
            pp = element('pPr')
            pp.extend([element('spacing',before=0,after=0,line=1,lineRule='exact'),
                       element('outlineLvl',val=9),previous_section])
            boundary.append(pp)
            heading.addprevious(boundary)
            inserted.append(boundary)
            numbering = current.find(q('pgNumType'))
            if numbering is not None:
                numbering.attrib.pop(q('start'),None)
        if index:
            section_property(current,'type',('pgSz','pgMar','paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg','textDirection','bidi','rtlGutter','docGrid')).set(q('val'),'nextPage')
        pp = properties(heading)
        sid = val(heading,'pPr/pStyle')
        if sid not in styles:
            raise ValueError('一级标题需要明确的段落样式')
        style = styles[sid]
        probe=element('p')
        probe_props=element('pPr')
        probe_props.append(element('pStyle',val=sid))
        probe.append(probe_props)
        if heading_level(probe,styles)!=1:
            raise ValueError('一级标题需使用具有一级大纲的段落样式')
        # A parent style may also be used for lower headings. Override those descendants below.
        spp = style.find(q('pPr'))
        if spp is None:
            spp=element('pPr')
            rp=style.find(q('rPr'))
            style.insert(style.index(rp) if rp is not None else len(style),spp)
        for tag in FLAGS:
            flag(spp,tag)
            direct = pp.find(q(tag))
            if direct is not None:
                pp.remove(direct)
        changes.append({'heading':visible(heading),'section_break':'nextPage','style_id':sid})
    return inserted,list(reversed(changes))


def keep_subheadings(root, style_root):
    styles=definitions(style_root)
    for p in root.find(q('body')).findall(q('p')):
        if heading_level(p,styles) not in (2,3):
            continue
        style=styles.get(val(p,'pPr/pStyle'))
        if style is None:
            continue
        pp=style.find(q('pPr'))
        if pp is None:
            pp=element('pPr')
            rp=style.find(q('rPr'))
            style.insert(style.index(rp) if rp is not None else len(style),pp)
        for tag in ('keepNext','keepLines'):
            flag(pp,tag)
            direct=p.find('./'+q('pPr')+'/'+q(tag))
            if direct is not None:
                direct.getparent().remove(direct)


def independent_stories(parts, root):
    """Give every section its own header/footer parts, including currently inactive variants."""
    rels, types = xml(parts[RELS]), xml(parts[TYPES])
    relationships = {r.get('Id'):r for r in rels}
    section_list = sections(root.find(q('body')))
    inherited, resolved = {}, []
    for section in section_list:
        for kind in ('header','footer'):
            refs = section.findall(q(kind+'Reference'))
            for ref in refs:
                relationship = relationships.get(ref.get('{'+REL+'}id'))
                if relationship is None or relationship.get('TargetMode')=='External':
                    raise ValueError('页眉页脚关系无效')
                target = posixpath.normpath(posixpath.join('word',relationship.get('Target')))
                if target not in parts or not target.startswith('word/'):
                    raise ValueError('页眉页脚部件缺失或路径异常')
                inherited[kind,ref.get(q('type'))] = target
        resolved.append(dict(inherited))
    used, records = set(), []
    for number,(section, effective) in enumerate(zip(section_list,resolved),1):
        for kind in ('header','footer'):
            for variant in ('default','first','even'):
                target = effective.get((kind,variant))
                ref = next((r for r in section.findall(q(kind+'Reference')) if r.get(q('type'))==variant),None)
                if target is not None and target not in used and ref is not None:
                    used.add(target)
                    continue
                index = 1
                while f'word/{kind}PaperFit{index}.xml' in parts:
                    index += 1
                name = f'{kind}PaperFit{index}.xml'
                path = 'word/'+name
                if target:
                    parts[path] = parts[target]
                    source_rels = posixpath.dirname(target)+'/_rels/'+posixpath.basename(target)+'.rels'
                    if source_rels in parts:
                        parts['word/_rels/'+name+'.rels'] = parts[source_rels]
                else:
                    tree = element('hdr' if kind=='header' else 'ftr')
                    tree.append(element('p'))
                    parts[path] = ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
                rid = 'rIdPaperFitFlow'+str(len(relationships)+1)
                while rid in relationships:
                    rid += '_'
                relation = ET.SubElement(rels,ET.QName(ET.QName(rels).namespace,'Relationship'),
                                         Id=rid,Type=REL+'/'+kind,Target=name)
                relationships[rid]=relation
                ET.SubElement(types,ET.QName(ET.QName(types).namespace,'Override'),PartName='/'+path,
                              ContentType=f'application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml')
                if ref is None:
                    ref=element(kind+'Reference',type=variant)
                    tags = (q('headerReference'),) if kind=='header' else (q('headerReference'),q('footerReference'))
                    section.insert(sum(n.tag in tags for n in section),ref)
                ref.set('{'+REL+'}id',rid)
                used.add(path)
                records.append({'section':number,'kind':kind,'variant':variant,'part':path,'copied_from':target})
    return {RELS:rels,TYPES:types},records


def content_signature(root):
    """All text/fields/bookmarks/math plus drawings stripped only of display dimensions."""
    result = {}
    for name in ('t','instrText','fldChar','bookmarkStart','bookmarkEnd','object','pict'):
        result[name] = [canonical(n) for n in root.iter(q(name))]
    drawings=[]
    for node in root.iter(q('drawing')):
        copy=deepcopy(node)
        for tag in ('{'+WP+'}extent','{'+A+'}ext'):
            for extent in copy.iter(tag):
                extent.attrib.pop('cx',None); extent.attrib.pop('cy',None)
        drawings.append(canonical(copy))
    result['drawings']=drawings
    result['math']=math_snapshot(root)
    result['table_cells']=[len(list(t.iter(q('tc')))) for t in root.iter(q('tbl'))]
    return result


def layout_projection(root, inserted_indices=()):
    tree=deepcopy(root)
    body=tree.find(q('body'))
    for index in sorted(inserted_indices,reverse=True):
        body.remove(body[index])
    for tag in (*FLAGS,'cantSplit','sectPr'):
        for node in list(tree.iter(q(tag))):
            node.getparent().remove(node)
    for tag in ('pPr','trPr'):
        for node in list(tree.iter(q(tag))):
            if not len(node) and not node.attrib:
                node.getparent().remove(node)
    for drawing in tree.iter(q('drawing')):
        for tag in ('{'+WP+'}extent','{'+A+'}ext'):
            for node in drawing.iter(tag):
                node.attrib.pop('cx',None)
                node.attrib.pop('cy',None)
    return canonical(tree)


def disable_picture_compression(settings):
    node=settings.find(q('doNotAutoCompressPictures'))
    if node is None:
        node=element('doNotAutoCompressPictures')
        following={q(n) for n in ('forceUpgrade','captions','readModeInkLockDown','smartTagType',
                                 'shapeDefaults','doNotEmbedSmartTags','decimalSymbol','listSeparator')}
        index=next((i for i,n in enumerate(settings) if n.tag in following or ET.QName(n).namespace not in (
            'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'http://schemas.openxmlformats.org/officeDocument/2006/math')),len(settings))
        settings.insert(index,node)
    node.set(q('val'),'1')


def enforce(source, output):
    source,output=Path(source),Path(output)
    record_path=Path(str(output)+'.report.json')
    if output.suffix.lower()!='.docx' or output.exists() or record_path.exists():
        raise ValueError('请指定新的 DOCX 和报告路径')
    original=read_package(source)
    parts=dict(original)
    root,styles=xml(parts[DOC]),xml(parts[STYLES])
    if any(list(root.iter(q(tag))) for tag in REVISIONS) or any(k.startswith('_xmlsignatures/') for k in parts):
        raise ValueError('修订或签名文档需要先复核')
    settings=xml(parts['word/settings.xml'])
    if any(settings.find(q(tag)) is not None for tag in ('documentProtection','trackRevisions')):
        raise ValueError('文档保护或修订已开启')
    original_levels={p:heading_level(p,definitions(styles)) for p in root.iter(q('p'))}
    projected=layout_projection(root)
    before=content_signature(root)
    matches=match_captions(source)
    caption_changes,pending=bind_captions(root,styles,matches)
    inserted,heading_changes=split_chapters(root,styles)
    keep_subheadings(root,styles)
    # Stop page-break inheritance at lower-heading styles and paragraph overrides as needed.
    defs=definitions(styles)
    old_styles=xml(original[STYLES])
    old_defs=definitions(old_styles)
    for p,level in original_levels.items():
        if level!=1 and heading_level(p,defs)!=1:
            sid=val(p,'pPr/pStyle')
            if sid and paragraph_values(styles,defs,sid).get('keepNext') is not None:
                # Preserve pre-existing effective flags outside level-one headings.
                old_values=paragraph_values(old_styles,old_defs,sid)
                new_values=paragraph_values(styles,defs,sid)
                for tag in FLAGS:
                    if level in (2,3) and tag in ('keepNext','keepLines'):
                        continue
                    if old_values.get(tag)!=new_values.get(tag) and p.find('./'+q('pPr')+'/'+q(tag)) is None:
                        flag(properties(p),tag,old_values.get(tag,{}).get('val','0') not in ('0','false','off'))
    trees,story_changes=independent_stories(parts,root)
    disable_picture_compression(settings)
    trees.update({DOC:root,STYLES:styles,'word/settings.xml':settings})
    if before!=content_signature(root):
        raise ValueError('文字、域、书签或对象内容校验失败')
    if projected!=layout_projection(root,[root.find(q('body')).index(n) for n in inserted]):
        raise ValueError('分页属性以外的文档结构发生变化')
    if any(original[k]!=parts[k] for k in original):
        raise ValueError('原包部件被意外覆盖')
    report={'captions':caption_changes,'headings':heading_changes,'sections':len(sections(root.find(q('body')))),
            'header_footer_changes':story_changes,'pending':pending,'content_integrity':True,
            'automatic_picture_compression':False,
            'status':'needs_review' if pending else 'requires_layout_check',
            'layout_check':'需在 Word/WPS 核对题注同页，更新整个目录。超过一页的表格须另行处理。',
            'source_sha256':sha256(source.read_bytes()).hexdigest()}
    with TemporaryDirectory(dir=output.parent) as tmp:
        doc=Path(tmp)/'result.docx'
        save_package(parts,trees,doc)
        result=read_package(doc)
        report['changed_parts']=[k for k in result if original.get(k)!=result[k]]
        report['output_sha256']=sha256(doc.read_bytes()).hexdigest()
        record=Path(tmp)/'report.json'
        record.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        os.link(record,record_path)
        try:
            os.link(doc,output)
        except BaseException:
            record_path.unlink()
            raise
    return report

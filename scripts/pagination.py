"""Split a PaperFit TOC from the body and restart the body's native page field."""
from copy import deepcopy
from hashlib import sha256
import json
import os
from pathlib import Path
from tempfile import TemporaryDirectory

from lxml import etree as ET
from paperfit import q
from phase2 import DOC, STYLES, canonical, read_package, save_package, xml
from structured import element, run
from toc import heading_level, instruction, marker, val
from word_fields import scan

REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
RELS = 'word/_rels/document.xml.rels'
TYPES = '[Content_Types].xml'


def section_property(section, tag, following):
    node = section.find(q(tag))
    if node is None:
        node = element(tag)
        successors = {q(t) for t in following}
        index = next((i for i, child in enumerate(section) if child.tag in successors), len(section))
        section.insert(index, node)
    return node


def add_footer(parts, section):
    rels, types = xml(parts[RELS]), xml(parts[TYPES])
    number = 1
    used = {node.get('PartName') for node in types}
    while f'word/footer{number}.xml' in parts or f'/word/footer{number}.xml' in used:
        number += 1
    target = f'footer{number}.xml'
    rid = 'rIdPaperFitPage'
    existing = {node.get('Id') for node in rels}
    while rid in existing:
        rid += '_'
    ET.SubElement(rels, ET.QName(ET.QName(rels).namespace, 'Relationship'),
                  Id=rid, Type=REL+'/footer', Target=target)
    ET.SubElement(types, ET.QName(ET.QName(types).namespace, 'Override'), PartName='/word/'+target,
                  ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
    ref = element('footerReference', type='default')
    ref.set('{'+REL+'}id', rid)
    section.insert(sum(n.tag == q('headerReference') for n in section), ref)
    footer, p, props = element('ftr'), element('p'), element('pPr')
    props.extend([element('ind', left=0, right=0, firstLine=0, firstLineChars=0),
                  element('jc', val='center')])
    p.append(props)
    marker(p, 'begin')
    instruction(p, ' PAGE ')
    marker(p, 'separate')
    p.append(run('1'))
    marker(p, 'end')
    footer.append(p)
    parts['word/'+target] = ET.tostring(footer, encoding='UTF-8', xml_declaration=True)
    return {RELS: rels, TYPES: types}


def paginate(source, output):
    source, output = Path(source), Path(output)
    report_path = Path(str(output) + '.report.json')
    if output.suffix.lower() != '.docx' or output.exists() or report_path.exists():
        raise ValueError('请使用新的 DOCX 和报告路径')
    parts = read_package(source)
    root = xml(parts[DOC])
    body = root.find(q('body'))
    sections = list(body.iter(q('sectPr')))
    if len(sections) != 1 or body[-1] is not sections[0]:
        raise ValueError('此入口仅从单节稿开始，已有分节请核对后处理')
    section = sections[0]
    settings = xml(parts['word/settings.xml'])
    if (section.find(q('footerReference')) is not None or section.find(q('titlePg')) is not None
            or settings.find(q('evenAndOddHeaders')) is not None):
        raise ValueError('已有页脚、首页或奇偶页设置，需要核对后处理')
    for name, data in parts.items():
        if name.startswith('word/header') and name.endswith('.xml'):
            fields, errors = scan(xml(data))
            if errors or any(f.code.upper().split()[:1] == ['PAGE'] for f in fields):
                raise ValueError('页眉已有页码或复杂域，请先核对')
    blocks = [s for s in body.findall(q('sdt')) if val(s, 'sdtPr/tag') == 'PaperFitTOC']
    if len(blocks) != 1:
        raise ValueError('需要一份带唯一 PaperFit 原生目录的文档')
    block = blocks[0]
    heading = block.getnext()
    styles = {s.get(q('styleId')): s for s in xml(parts[STYLES]).findall(q('style'))}
    if heading is None or heading.tag != q('p') or not heading_level(heading, styles):
        raise ValueError('目录后未紧接分级标题，请核对正文起点')
    content = block.find(q('sdtContent'))
    if content is None or not len(content):
        raise ValueError('目录缺少内容')
    last = content[-1]
    br = last.find('./'+q('r')+'/'+q('br'))
    if (last.tag != q('p') or br is None or br.get(q('type')) != 'page'
            or list(last.iter()) != [last, br.getparent(), br]):
        raise ValueError('目录末尾的独立分页符已被编辑，请核对边界')
    fields, errors = scan(block)
    if errors or sum(f.code.upper().split()[:1] == ['TOC'] for f in fields) != 1:
        raise ValueError('目录域结构异常')
    # Outside the content control, the section boundary survives Word rebuilding the TOC.
    front = deepcopy(section)
    content.remove(last)
    boundary, pp = element('p'), element('pPr')
    pp.append(front)
    boundary.append(pp)
    block.addnext(boundary)
    tail = ('cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection',
            'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange')
    section_property(section, 'type', ('pgSz','pgMar','paperSrc','pgBorders','lnNumType','pgNumType') + tail).set(q('val'), 'nextPage')
    number = section_property(section, 'pgNumType', tail)
    number.attrib.clear()
    number.set(q('fmt'), 'decimal')
    number.set(q('start'), '1')
    trees = add_footer(parts, section)
    trees[DOC] = root
    changes = []
    for field in fields:
        if field.code.upper().split()[:1] == ['PAGEREF']:
            if len(field.result) != 1 or len(field.code.split()) < 2:
                raise ValueError('目录页码域结构异常')
            changes.append({'bookmark': field.code.split()[1], 'before': field.value})
            field.result[0].text = '…'
        if field.code.upper().split()[:1] in (['PAGEREF'], ['TOC']):
            field.owner.attrib.pop(q('dirty'), None)
    if not changes:
        raise ValueError('目录缺少可更新的页码域')
    # Compare the complete original text, objects, styles and bookmarks outside owned regions.
    original = xml(parts[DOC]).find(q('body'))
    untouched = [canonical(n) for n in body if n not in (block, boundary, section)]
    pos = body.index(block)
    if untouched != [canonical(n) for i, n in enumerate(original) if i not in (pos, len(original)-1)]:
        raise ValueError('正文完整性检查失败')
    with TemporaryDirectory(dir=output.parent) as tmp:
        doc = Path(tmp) / 'result.docx'
        save_package(parts, trees, doc)
        report = {'sections': 2, 'front_footer': None, 'body_page_start': 1, 'page_field': 'PAGE',
                  'alignment': 'center', 'toc_entries': changes, 'integrity': True,
                  'page_cache': '请用 Word 更新整个目录以计算新页码',
                  'source_sha256': sha256(source.read_bytes()).hexdigest(),
                  'output_sha256': sha256(doc.read_bytes()).hexdigest()}
        record = Path(tmp) / 'report.json'
        record.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        os.link(record, report_path)
        try:
            os.link(doc, output)
        except BaseException:
            report_path.unlink()
            raise
    return report

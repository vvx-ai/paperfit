"""Add a native Word TOC at the first body heading in a new DOCX."""
from hashlib import sha256
import json
import os
from pathlib import Path
import re
from tempfile import TemporaryDirectory

from paperfit import q
from phase2 import DOC, STYLES, canonical, read_package, save_package, xml
from structured import element, run, visible
from word_fields import scan

CODE = r' TOC \o "1-3" \h \u '
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def val(parent, path, default=None):
    node = parent.find('/'.join(q(name) for name in path.split('/')))
    return default if node is None else node.get(q('val'), default)


def heading_level(p, styles):
    raw = val(p, 'pPr/outlineLvl')
    sid = val(p, 'pPr/pStyle')
    seen = set()
    while raw is None and sid in styles:
        if sid in seen:
            raise ValueError('段落样式继承循环')
        seen.add(sid)
        style = styles[sid]
        raw = val(style, 'pPr/outlineLvl')
        native = re.fullmatch(r'(?:heading|标题)\s*([1-9])', val(style, 'name', ''), re.I)
        if raw is None and native:
            raw = str(int(native[1]) - 1)
        sid = val(style, 'basedOn')
    return int(raw) + 1 if raw in ('0', '1', '2') else None


def marker(parent, kind, **attrs):
    r = element('r')
    r.append(element('fldChar', fldCharType=kind, **attrs))
    parent.append(r)


def instruction(parent, code):
    r = element('r')
    text = element('instrText')
    text.set(XML_SPACE, 'preserve')
    text.text = code
    r.append(text)
    parent.append(r)


def toc_block():
    block = element('sdt')
    props = element('sdtPr')
    props.append(element('tag', val='PaperFitTOC'))
    gallery = element('docPartObj')
    gallery.append(element('docPartGallery', val='Table of Contents'))
    gallery.append(element('docPartUnique'))
    props.append(gallery)
    block.append(props)
    content = element('sdtContent')
    block.append(content)
    title = element('p')
    pp = element('pPr')
    pp.extend([element('keepNext'), element('jc', val='center'), element('outlineLvl', val=9)])
    title.append(pp)
    title.append(run('目录'))
    content.append(title)
    return block, content


def entry_styles(styles, levels, width):
    """Reuse native TOC styles, or define the missing levels for the cached entries."""
    result = {}
    existing = styles.findall(q('style'))
    ids = {s.get(q('styleId')) for s in existing}
    for level in sorted(levels):
        matches = [s for s in existing if s.get(q('type')) == 'paragraph'
                   and val(s, 'name', '').lower().replace(' ', '') in (f'toc{level}', f'目录{level}')]
        if len(matches) > 1:
            raise ValueError('存在多个同级目录样式，请核对')
        if matches:
            result[level] = matches[0].get(q('styleId'))
            continue
        sid = f'TOC{level}'
        if sid in ids:
            raise ValueError('目录样式编号已被其他样式使用')
        style = element('style', type='paragraph', styleId=sid)
        style.append(element('name', val=f'toc {level}'))
        pp = element('pPr')
        tabs = element('tabs')
        tabs.append(element('tab', val='right', leader='dot', pos=width))
        pp.extend([tabs, element('ind', left=(level - 1) * 280, firstLine=0), element('outlineLvl', val=9)])
        style.append(pp)
        styles.append(style)
        result[level] = sid
    return result


def add_toc(source, output):
    source, output = Path(source), Path(output)
    report_path = Path(str(output) + '.report.json')
    if output.suffix.lower() != '.docx' or output.exists() or report_path.exists():
        raise ValueError('请指定尚不存在的 DOCX 和报告路径')
    parts = read_package(source)
    root, styles = xml(parts[DOC]), xml(parts[STYLES])
    body = root.find(q('body'))
    fields, errors = scan(body)
    if errors:
        raise ValueError('原稿存在未闭合的域，请先复核')
    if any(re.match(r'^TOC\b', f.code, re.I) for f in fields):
        raise ValueError('文档已有目录域，请在 Word 中选择「更新整个目录」')
    if any(re.sub(r'\s+', '', visible(p)).lower() in ('目录', 'contents', 'tableofcontents')
           for p in body.iter(q('p'))):
        raise ValueError('检测到已有目录标题，请先核对位置')
    sections = list(body.iter(q('sectPr')))
    if len(sections) != 1 or sections[0].getparent() is not body:
        raise ValueError('当前目录入口支持单节正文，请先核对分节')
    section = sections[0]
    columns = section.find(q('cols'))
    if columns is not None and columns.get(q('num'), '1') != '1':
        raise ValueError('多栏目录需要人工核对')
    size, margins = section.find(q('pgSz')), section.find(q('pgMar'))
    if size is None or margins is None:
        raise ValueError('缺少纸宽或页边距，无法定位目录页码')
    width = int(size.attrib[q('w')]) - sum(int(margins.attrib[q(k)]) for k in ('left', 'right')) - int(margins.get(q('gutter'), '0'))
    by_id = {s.get(q('styleId')): s for s in styles.findall(q('style')) if s.get(q('type')) == 'paragraph'}
    headings = []
    for p in body.iter(q('p')):
        level = heading_level(p, by_id)
        if level and visible(p).strip():
            if p.getparent() is not body:
                raise ValueError('嵌套标题需要人工确认目录范围')
            headings.append((p, level))
    if not headings:
        raise ValueError('未识别到一至三级大纲标题')
    style_ids = entry_styles(styles, {level for _, level in headings}, width)
    block, content = toc_block()
    body.insert(body.index(headings[0][0]), block)
    names = {b.get(q('name'), '').casefold() for b in root.iter(q('bookmarkStart'))}
    ids = [int(b.get(q('id'))) for b in root.iter() if b.tag in (q('bookmarkStart'), q('bookmarkEnd'))]
    next_id = max(ids, default=-1) + 1
    entries, added_marks = [], []
    for index, (heading, level) in enumerate(headings):
        while f'_PaperFitToc{next_id}'.casefold() in names:
            next_id += 1
        anchor = f'_PaperFitToc{next_id}'
        start = element('bookmarkStart', id=next_id, name=anchor)
        end = element('bookmarkEnd', id=next_id)
        next_id += 1
        heading.insert(int(heading.find(q('pPr')) is not None), start)
        heading.append(end)
        added_marks.extend([start, end])
        p, pp = element('p'), element('pPr')
        pp.append(element('pStyle', val=style_ids[level]))
        p.append(pp)
        if index == 0:
            marker(p, 'begin', dirty='true')
            instruction(p, CODE)
            marker(p, 'separate')
        link = element('hyperlink', anchor=anchor)
        text = visible(heading).strip()
        link.append(run(text))
        tab = element('r')
        tab.append(element('tab'))
        link.append(tab)
        marker(link, 'begin', dirty='true')
        instruction(link, f' PAGEREF {anchor} \\h ')
        marker(link, 'separate')
        link.append(run('…'))
        marker(link, 'end')
        p.append(link)
        if index == len(headings) - 1:
            marker(p, 'end')
        content.append(p)
        entries.append({'heading': text, 'level': level, 'bookmark': anchor})
    page = element('p')
    r = element('r')
    r.append(element('br', type='page'))
    page.append(r)
    content.append(page)
    with TemporaryDirectory(dir=output.parent) as scratch:
        doc = Path(scratch) / 'toc.docx'
        save_package(parts, {DOC: root, STYLES: styles}, doc)
        body.remove(block)
        for mark in added_marks:
            mark.getparent().remove(mark)
        if canonical(root) != canonical(xml(parts[DOC])):
            raise ValueError('目录以外的原稿结构发生变化')
        result = {'entries': entries, 'native_field': CODE.strip(), 'integrity': True,
                  'source_sha256': sha256(source.read_bytes()).hexdigest(),
                  'output_sha256': sha256(doc.read_bytes()).hexdigest(),
                  'page_cache': '请在 Word 中更新整个目录以计算页码'}
        record = Path(scratch) / 'report.json'
        record.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        os.link(record, report_path)
        try:
            os.link(doc, output)
        except BaseException:
            report_path.unlink()
            raise
    return result

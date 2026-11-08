"""Local extraction of reviewable requirements, with a location for every clause."""
import re
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET
from paperfit import parse, q


def source_blocks(path):
    suffix = path.suffix.lower()
    if suffix == '.pdf':
        from pypdf import PdfReader
        for i, page in enumerate(PdfReader(path).pages, 1):
            yield f'第{i}页', page.extract_text() or ''
            resources = page.get('/Resources', {})
            objects = resources.get_object().get('/XObject', {}) if resources else {}
            if objects and any(ref.get_object().get('/Subtype') == '/Image' for ref in objects.get_object().values()):
                yield f'第{i}页图像区域待核对', ''
    elif suffix == '.docx':
        with ZipFile(path) as archive:
            root = ET.fromstring(archive.read('word/document.xml'))
            for name in archive.namelist():
                if re.fullmatch(r'word/(?:header\d*|footer\d*|footnotes|endnotes)\.xml',name):
                    extra = ET.fromstring(archive.read(name))
                    if any((n.text or '').strip() for n in extra.iter(q('t'))):
                        yield f'{name} 附加区域待核对', ''
        for i, item in enumerate(root.find(q('body')), 1):
            if item.tag == q('p'):
                value = ''.join(item.itertext(q('t')))
                if value.strip():
                    yield f'正文块{i}', value
                if item.find('.//' + q('drawing')) is not None:
                    yield f'正文块{i}图像区域待核对', ''
            elif item.tag == q('tbl'):
                for j, row in enumerate(item.findall(q('tr')), 1):
                    cells = [' '.join(''.join(p.itertext(q('t'))) for p in cell.findall(q('p'))) for cell in row.findall(q('tc'))]
                    yield f'表格块{i}第{j}行', ' '.join(cells)
            elif item.tag != q('sectPr'):
                yield f'正文块{i}特殊结构待核对', ''
    elif suffix == '.txt':
        for i, value in enumerate(path.read_text(encoding='utf-8-sig').splitlines(), 1):
            if value.strip():
                yield f'第{i}行', value
    else:
        raise ValueError('规范文件支持 TXT、DOCX、PDF')


def clause_rules(text):
    match = re.fullmatch(r'(图表|图片和表格|公式)(?:题注)?[：:\s]*(?:按)?(全文连续|章节)(?:自动)?编号', text)
    if match:
        key = 'equations' if match[1] == '公式' else 'captions'
        return [(('structure', key), 'continuous' if match[2] == '全文连续' else 'chapter')], []
    match = re.fullmatch(r'参考文献[：:\s]*(?:按)?(?:顺序|连续)编号(?:[，,、\s]*悬挂缩进\s*(\d+(?:\.\d+)?)\s*字符)?', text)
    if match:
        props = [(('structure', 'references', 'number'), True)]
        if match[1]:
            props.append((('structure', 'references', 'hanging_chars'), float(match[1])))
        return props, []
    parsed = parse(text)
    changes = [(('styles', role, k), v) for role, attrs in parsed['styles'].items() for k, v in attrs.items()]
    changes += [(('page', k), v) for k, v in parsed['page'].items()]
    return changes, parsed['pending']


def extract(path):
    result = {'styles': {}, 'page': {}, 'structure': {}, 'pending': [], 'sources': []}
    origins = {}
    for location, block in source_blocks(Path(path)):
        if not block.strip():
            result['pending'].append(f'{location} 缺少可提取文字，请人工核对')
        block = re.sub(r'[，,](?=\s*(?:一[级]|二级|三级|正文|公式|参考文献|图表))', '；', block)
        for clause in filter(None, (s.strip() for s in re.split(r'[；;。\n]', block))):
            changes, pending = clause_rules(clause)
            result['sources'].append({'location': location, 'text': clause})
            result['pending'].extend(f'{location} {issue}' for issue in pending)
            if not changes and not pending:
                result['pending'].append(f'{location} 缺少可执行要求：{clause}')
            for keys, value in changes:
                target = result
                for key in keys[:-1]:
                    target = target.setdefault(key, {})
                if keys in origins and target[keys[-1]] != value:
                    result['pending'].append(f'{origins[keys]} 与 {location} 冲突：{".".join(keys)}')
                else:
                    origins[keys] = location
                    target[keys[-1]] = value
    if not result['sources'] and not result['pending']:
        result['pending'].append('规范没有可提取的规则，请核对原文件')
    return result

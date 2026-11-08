"""Phase 2 CLI and reusable offline processing pipeline."""
import argparse
from collections import Counter
from hashlib import sha256
import json
import math
import os
from pathlib import Path
import re
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from lxml import etree as ET

from paperfit import format_docx, math_snapshot, q, W
from analyze_docx import analyze
from bibliography import references
from isolate_layout_styles import isolate
from match_captions import match_captions
from normalize_captions import normalize
from rule_sources import extract
from structured import captions, chapters, equations, visible, MATH, CAP, MENTION
from word_fields import audit
from style_audit import audit_styles
from text_integrity import TextIntegrity

DOC = 'word/document.xml'
STYLES = 'word/styles.xml'
SETTINGS = 'word/settings.xml'
STORY = re.compile(r'word/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$')


def read_package(path):
    with ZipFile(path) as z:
        if len(z.namelist()) != len(set(z.namelist())):
            raise ValueError('文件包含重复的包部件')
        return {n:z.read(n) for n in z.namelist()}


def xml(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def stories(parts):
    return {k:xml(v) for k,v in parts.items() if STORY.fullmatch(k)}


def save_package(parts, trees, path):
    with ZipFile(path,'x',compression=8) as z:
        for name, data in parts.items():
            if name in trees and canonical(trees[name]) != canonical(xml(data)):
                data = ET.tostring(trees[name], encoding='UTF-8', xml_declaration=True)
            z.writestr(name,data)


def canonical(node):
    return ET.tostring(node,method='c14n',exclusive=True)


def managed(table):
    marker = table.find('./'+q('tblPr')+'/'+q('tblCaption'))
    return marker is not None and marker.get(q('val')) == 'PaperFit:equation:v1'


def logical_paragraphs(root):
    result = []
    for node in root.find(q('body')):
        if node.tag == q('p'):
            result.append(node)
        elif node.tag == q('tbl') and managed(node):
            cells = node.findall('./'+q('tr')+'/'+q('tc'))
            if len(cells) != 3 or len(cells[1].findall(q('p'))) != 1 or len(list(cells[1].iter(MATH))) != 1:
                raise ValueError('已有公式布局已编辑，需要复核')
            result.append(cells[1].find(q('p')))
    return result


def inventory(path, plan=None):
    parts = read_package(path)
    docs = stories(parts)
    info = analyze(path)
    matches = match_captions(path)
    refs = audit(docs)
    known = {c['caption_body_index'] for c in matches['matches']}
    body = docs[DOC].find(q('body'))
    mentions = []
    numbers = Counter((c['kind'], c['number']) for c in matches['matches'])
    for idx,node in enumerate(body):
        if idx not in known and node.tag == q('p'):
            for m in MENTION.finditer(visible(node)):
                n = numbers[m[1],m[2]]
                mentions.append({'body_index':idx,'text':m[0],'targets':n})
                if n != 1:
                    refs['issues'].append({'body_index':idx,'issue':'图表文字引用缺少唯一目标','mention':m[0]})
    logical = logical_paragraphs(docs[DOC])
    result = {'document':info,'caption_matches':matches,'field_audit':refs,'textual_mentions':mentions,
              'paragraph_map':[{'index':i,'text':visible(p),'math':len(list(p.iter(MATH)))} for i,p in enumerate(logical)]}
    if plan is not None:
        result['style_audit'] = audit_styles(logical,xml(parts[STYLES]),plan.get('paragraph_roles',{}),plan.get('styles',{}))
    return result


def check_keys(obj, allowed, name):
    if not isinstance(obj,dict) or set(obj)-set(allowed):
        raise ValueError(name+' 含未知配置项或类型有误')


def bounded(value, low, high):
    return type(value) in (float,int) and math.isfinite(value) and low <= value <= high


def check_plan(plan, count):
    check_keys(plan, {'styles','page','pending','sources','paragraph_roles','structure','source_sha256','native_references'},'规则')
    if 'native_references' in plan and type(plan['native_references']) is not bool:
        raise ValueError('native_references 必须为布尔值')
    if plan.get('pending'):
        raise ValueError('存在待确认规则，请先核对规范来源')
    opts = plan.get('structure',{})
    check_keys(opts, {'captions','chapters','equations','references','position','sync_refs','title_index'},'structure')
    for k in ('captions','equations'):
        if k in opts and opts[k] not in ('continuous','chapter'):
            raise ValueError(k+' 仅支持 continuous 或 chapter')
    if 'sync_refs' in opts and type(opts['sync_refs']) is not bool:
        raise ValueError('sync_refs 必须为布尔值')
    pos = opts.get('position',{})
    check_keys(pos, {'图','表'},'position')
    if any(v not in ('before','after') for v in pos.values()) or pos and 'captions' not in opts:
        raise ValueError('题注位置需要编号模式及 before/after')
    def index(i):
        if type(i) is not int or not 0 <= i < count:
            raise ValueError('段落索引超出原稿范围')
    for name in ('chapters','paragraph_roles'):
        items = opts.get(name,{}) if name=='chapters' else plan.get(name,{})
        if not isinstance(items,dict):
            raise ValueError(name+' 需要段落映射')
        for i,v in items.items():
            if not re.fullmatch(r'0|[1-9]\d*',i):
                raise ValueError('段落索引必须为非负整数')
            index(int(i))
            if name == 'chapters' and (type(v) is not int or v < 1):
                raise ValueError('章节号必须为正整数')
            if name=='paragraph_roles' and v not in ('preserve','body','heading1','heading2','heading3'):
                raise ValueError('段落角色不受支持')
    chapter_ids = list(opts.get('chapters',{}).values())
    if len(set(chapter_ids)) != len(chapter_ids):
        raise ValueError('章节号重复，请核对')
    if 'title_index' in opts:
        index(opts['title_index'])
    if 'references' in opts:
        ref = opts['references']
        check_keys(ref, {'entries','number','hanging_chars','style'},'references')
        ids = ref.get('entries',[])
        if not isinstance(ids,list) or not ids:
            raise ValueError('文献处理需要核对后的 entries 段落索引')
        for i in ids:
            index(i)
        if ids != sorted(set(ids)):
            raise ValueError('文献索引应按顺序排列且唯一')
        if 'number' in ref and type(ref['number']) is not bool:
            raise ValueError('references.number 应为布尔值')
        if 'hanging_chars' in ref and not bounded(ref['hanging_chars'],0,20):
            raise ValueError('悬挂缩进字符数无效')
        style = ref.get('style',{})
        check_keys(style,{'font','size','line'},'references.style')
        if any(not bounded(v,0.1,200) for k,v in style.items() if k!='font') or ('font' in style and not isinstance(style['font'],str)):
            raise ValueError('文献样式属性无效')
    if plan.get('styles') and not plan.get('paragraph_roles'):
        raise ValueError('请依据 audit 的段落索引填写 paragraph_roles；未选择段落将保留')


def verify(original, result):
    if set(original) != set(result):
        raise ValueError('DOCX包部件清单发生变化')
    a,b = (xml(p[DOC]) for p in (original,result))
    if math_snapshot(a) != math_snapshot(b):
        raise ValueError('原生公式内容或顺序改变')
    for tag in ('drawing','object','pict','tbl'):
        def snapshots(root):
            return [canonical(n) for n in root.iter(q(tag)) if tag != 'tbl' or not managed(n)]
        if snapshots(a) != snapshots(b):
            raise ValueError(tag+' 原始对象完整性校验失败')
    for k in original:
        if k not in (STYLES, SETTINGS) and not STORY.fullmatch(k) and original[k] != result[k]:
            raise ValueError('意外改变包部件 '+k)
    return {'math':len(math_snapshot(b)), 'drawings':len(list(b.iter(q('drawing')))),
            'original_tables':sum(not managed(n) for n in b.iter(q('tbl'))),
            'unchanged_parts':sum(original[n] == result[n] for n in original)}


def process(source, plan, output):
    source, output = Path(source), Path(output)
    report_path = Path(str(output)+'.report.json')
    if output.suffix.lower() != '.docx' or output.exists() or report_path.exists() or output.resolve() == source.resolve():
        raise ValueError('请使用新的DOCX输出路径，报告和原稿须保留')
    digest = sha256(source.read_bytes()).hexdigest()
    if plan.get('source_sha256',digest) != digest:
        raise ValueError('配置所绑定的输入文件哈希不一致')
    original = read_package(source)
    original_root = xml(original[DOC])
    logical = logical_paragraphs(original_root)
    direct = original_root.find(q('body')).findall(q('p'))
    check_plan(plan,len(logical))
    opts = plan.get('structure',{})
    report = {'source_sha256':digest,'stages':{},'pending':[]}
    stages = report['stages']
    with TemporaryDirectory(dir=output.parent) as scratch:
        tmp = Path(scratch)
        current = source
        if plan.get('styles') or plan.get('page'):
            roles = {str(i):'preserve' for i in range(len(direct))}
            for idx,role in plan.get('paragraph_roles',{}).items():
                p = logical[int(idx)]
                if p in direct:
                    roles[str(direct.index(p))] = role
            config = dict(plan,paragraph_roles=roles)
            current = tmp/'styles.docx'
            stages['styles'] = format_docx(source,current,config)
            report['pending'].extend(stages['styles']['pending'])
        if plan.get('styles'):
            formatted = read_package(current)
            stages['style_audit'] = audit_styles(logical_paragraphs(xml(formatted[DOC])),xml(formatted[STYLES]),plan['paragraph_roles'],plan['styles'])
        if 'title_index' in opts:
            next_path = tmp/'layout.docx'
            idx = direct.index(logical[opts['title_index']])
            stages['layout'] = isolate(current,next_path,idx)
            current = next_path
        parts = read_package(current)
        docs = stories(parts)
        root = docs[DOC]
        guard = TextIntegrity(docs)
        paragraphs = logical_paragraphs(root)
        chapter_map = chapters(paragraphs,opts.get('chapters',{}))
        styles = xml(parts[STYLES])
        trees = {**docs, STYLES:styles}
        if 'captions' in opts:
            settings = xml(parts[SETTINGS])
            trees[SETTINGS] = settings
            stages['captions'] = captions(docs,styles,settings,match_captions(current),chapter_map,opts,guard)
        if 'equations' in opts:
            stages['equations'] = equations(docs,paragraphs,chapter_map,opts['equations'],opts.get('sync_refs',False),guard)
        if 'references' in opts:
            entries = {paragraphs[i] for i in opts['references']['entries']}
            for doc in docs.values():
                for p in doc.iter(q('p')):
                    st = p.find('./'+q('pPr')+'/'+q('pStyle'))
                    if st is not None and st.get(q('val')) == 'PaperFitBibliography' and p not in entries:
                        raise ValueError('文献样式用于范围外段落，请核对')
            if any(b.get(q('val')) == 'PaperFitBibliography' for b in styles.iter(q('basedOn'))):
                raise ValueError('文献样式影响其他派生样式')
            from native_references import owned_numbering
            from toc import val
            selected=[paragraphs[i] for i in opts['references']['entries']]
            ids={val(p,'pPr/numPr/numId') for p in selected}
            native_labels=None
            if len(ids)==1 and ids <= owned_numbering(parts):
                if [p for p in root.iter(q('p')) if val(p,'pPr/numPr/numId') in ids]==selected:
                    native_labels=list(range(1,len(selected)+1))
            stages['references'] = references(docs,styles,paragraphs,opts['references'],guard,native_labels)
        stages['fields'] = audit(docs,opts.get('sync_refs',False),guard)
        stages['text_integrity'] = guard.verify()
        for stage in stages.values():
            report['pending'].extend(stage.get('issues',[]))
        current = tmp/'structured.docx'
        save_package(parts,trees,current)
        if stages.get('captions',{}).get('changes'):
            final = tmp/'captions.docx'
            stages['caption_style'] = normalize(current,final)
            current = final
        result = read_package(current)
        report['integrity'] = verify(original,result)
        if plan.get('native_references', True):
            from native_references import convert_parts
            converted, native_report = convert_parts(result)
            stages['native_references'] = native_report
            report['pending'].extend(native_report['pending'])
            current = tmp/'native-references.docx'
            save_package(converted,{},current)
        report['output_sha256'] = sha256(current.read_bytes()).hexdigest()
        report['status'] = 'needs_review' if report['pending'] else 'passed'
        record = tmp/'report.json'
        record.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        os.link(record,report_path)
        try:
            os.link(current,output)
        except BaseException:
            report_path.unlink()
            raise
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description='PaperFit 本地结构化排版')
    commands = parser.add_subparsers(dest='cmd',required=True)
    for name in ('extract','audit','run'):
        sub = commands.add_parser(name)
        sub.add_argument('source',type=Path)
        if name=='audit':
            sub.add_argument('--rules',type=Path)
        if name=='run':
            sub.add_argument('rules',type=Path)
        sub.add_argument('output',type=Path)
    args = parser.parse_args(argv)
    try:
        if args.cmd == 'run':
            report = process(args.source,json.loads(args.rules.read_text(encoding='utf-8')),args.output)
            print(f"{report['status']}  {args.output}")
        else:
            data = extract(args.source) if args.cmd=='extract' else inventory(args.source,json.loads(args.rules.read_text(encoding='utf-8')) if args.rules else None)
            with args.output.open('x',encoding='utf-8') as f:
                json.dump(data,f,ensure_ascii=False,indent=2)
    except (ValueError, OSError, KeyError) as error:
        parser.exit(2, f'PaperFit: {error}\n')


if __name__=='__main__':
    main()

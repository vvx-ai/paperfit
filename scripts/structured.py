"""Scoped structural changes to native Word elements."""
from collections import Counter, defaultdict
import re
from lxml import etree as ET
from paperfit import M, q, setv
from caption_style import configure
from number_captions import insert_seq
from word_fields import scan

CAP = re.compile(r'^\s*([图表])\s*(\d+(?:[-－.]\d+)*)(?!\d)')
MENTION = re.compile(r'([图表])\s*(\d+(?:[-－.]\d+)*)')
EQ_MENTION = re.compile(r'((?:公式|式))\s*[（(]\s*(\d+(?:[-－.]\d+)*)\s*[)）]')
MATH = '{' + M + '}oMath'
DISPLAY = '{' + M + '}oMathPara'


def visible(p):
    return ''.join(p.itertext(q('t')))


def element(tag, **attrs):
    return ET.Element(q(tag), {q(k): str(v) for k, v in attrs.items()})


def run(value):
    node = element('r')
    text = ET.SubElement(node, q('t'))
    text.text = value
    if value != value.strip():
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return node


def seq_field(code, value):
    node = element('fldSimple', instr=' ' + code + ' ')
    node.append(run(str(value)))
    return node


def chapters(paragraphs, config):
    result = {}
    current = None
    for i, p in enumerate(paragraphs):
        if str(i) in config:
            current = config[str(i)]
        result[p] = current
    return result


def next_label(kind, mode, chapter, counters):
    if mode == 'chapter' and chapter is None:
        raise ValueError('按章编号需要通过 chapters 指定所属章节')
    prefix = f'{chapter}-' if mode == 'chapter' else ''
    key = (kind, prefix)
    counters[key] += 1
    count = counters[key]
    reset = r' \r 1' if prefix and count == 1 else ''
    return prefix, str(count), f'SEQ {kind}{reset} \\* ARABIC'


def numeric_node(p, start, end):
    offset = 0
    for node in p.iter(q('t')):
        finish = offset + len(node.text or '')
        if offset <= start < finish and end <= finish:
            return node, start-offset, end-offset
        offset = finish
    raise ValueError('编号跨文本节点，需要人工核对')


def guard_sequences(stories, selected, labels):
    """An extra SEQ in the main story would advance Word's counter on F9."""
    for f in scan(stories['word/document.xml'])[0]:
        tokens = f.code.split()
        if len(tokens) >= 2 and tokens[0].upper() == 'SEQ' and tokens[1].strip('"') in labels:
            if next(f.owner.iterancestors(q('p')),None) not in selected:
                raise ValueError('正文中还有未纳入编号范围的同名SEQ域，请核对')


def update_caption(p, match, prefix, digit, code):
    fields, issues = scan(p)
    if issues or len(fields) > 1:
        raise ValueError('题注域结构有歧义')
    if any(n.tag in {q('ins'), q('del'), q('sectPr'), q('pPrChange'), q('moveFrom'), q('moveTo')} for n in p.iter()):
        raise ValueError('题注带有修订或分节属性，请核对')
    if not fields:
        node = insert_seq(p, *match.span(2), digit, ' ' + code + ' ')
        if prefix:
            node.addprevious(run(prefix))
        return
    f = fields[0]
    if (not re.fullmatch(r'SEQ\s+(?:' + match[1] + '|' + {'图':'PaperFitFigure','表':'PaperFitTable'}[match[1]] + r')(?:\s+\\r\s+1)?(?:\s+\\\*\s+ARABIC)?', f.code, re.I)
            or f.nested or f.owner.get(q('fldLock')) in ('true','on','1') or len(f.result) != 1):
        raise ValueError('发现不支持或锁定的题注域')
    old = match[2]
    if not f.value.isdecimal() or not old.endswith(f.value):
        raise ValueError('题注显示值与域不一致')
    text_nodes = list(p.iter(q('t')))
    upto = sum(len(n.text or '') for n in text_nodes[:text_nodes.index(f.result[0])])
    if upto != match.end(2)-len(f.value):
        raise ValueError('题注域未精确覆盖编号')
    old_prefix = old[:-len(f.value)]
    if prefix != old_prefix:
        if old_prefix:
            n, left, right = numeric_node(p, match.start(2), upto)
            n.text = n.text[:left] + prefix + n.text[right:]
        elif prefix:
            anchor = f.owner if f.simple else f.owner.getparent()
            if anchor.getparent() is not p:
                raise ValueError('题注域结构不支持插入章节前缀')
            anchor.addprevious(run(prefix))
    if f.code != code:
        f.set_code(code)
    if f.value != digit:
        f.set_value(digit)


def sync_mentions(stories, selected, mapping, sync, guard, pattern=MENTION):
    changes, issues = [], []
    for part, root in stories.items():
        fields, errors = scan(root)
        if errors:
            raise ValueError('存在未闭合的域，请先修复')
        cached = {n for f in fields for n in f.result}
        for i, p in enumerate(root.iter(q('p'))):
            if p in selected:
                continue
            text = visible(p)
            for match in reversed(list(pattern.finditer(text))):
                key = ('公式' if pattern is EQ_MENTION else match[1], match[2])
                location = {'part': part, 'paragraph': i, 'mention': match[0]}
                # Word REF caches are handled by bookmark identity, never by label text.
                if cached:
                    offset = 0
                    covered = []
                    for n in p.iter(q('t')):
                        finish = offset + len(n.text or '')
                        if offset < match.end(2) and finish > match.start(2):
                            covered.append(n)
                        offset = finish
                    if covered and all(n in cached for n in covered):
                        continue
                targets = mapping.get(key, [])
                if len(targets) != 1:
                    issues.append(dict(location, reason='编号没有唯一目标'))
                    continue
                if re.match(r'\s*(?:[、,，~～–—至]|到)\s*\d',text[match.end():]):
                    issues.append(dict(location, reason='省略标签的组合引用需要逐项核对'))
                    continue
                if targets[0] == match[2]:
                    continue
                node, left, right = numeric_node(p, *match.span(2))
                if not sync or node.getparent().getparent() is not p or any(n.tag in {q('ins'),q('del'),q('fldChar')} for n in p.iter()):
                    raise ValueError(f'{part} 段落{i} 的 {match[0]} 需确认后同步')
                guard.replace(p,match.start(2),match[2],targets[0])
                node.text = node.text[:left] + targets[0] + node.text[right:]
                changes.append(dict(location, number=targets[0]))
    return changes, issues


def captions(stories, styles, settings, analysis, chapter_map, options, guard):
    root = stories['word/document.xml']
    body = root.find(q('body'))
    nodes = list(body)
    matches = analysis['matches']
    if analysis['unmatched_objects'] or any(c['status'] != 'unique_adjacent_match' for c in matches):
        raise ValueError('图表与题注不能唯一匹配，请检查 audit 报告')
    guard_sequences(stories, {nodes[c['caption_body_index']] for c in matches}, {'图','表','PaperFitFigure','PaperFitTable'})
    if not matches:
        return {'changes': [], 'issues': []}
    selected, changes, mapping, counters = set(), [], defaultdict(list), Counter()
    for item in matches:
        p = nodes[item['caption_body_index']]
        selected.add(p)
        match = CAP.match(visible(p))
        kind = item['kind']
        prefix, digit, code = next_label(kind, options['captions'], chapter_map[p], counters)
        mapping[kind, match[2]].append(prefix + digit)
        old = visible(p)
        guard.replace(p,match.start(2),match[2],prefix+digit)
        update_caption(p, match, prefix, digit, code)
        expected = old[:match.start(2)] + prefix + digit + old[match.end(2):]
        if visible(p) != expected:
            raise ValueError('题注文字完整性检查失败')
        changes.append({'before':old, 'after':expected,'field':code})
    refs, issues = sync_mentions(stories, selected, mapping, options.get('sync_refs', False), guard)
    if issues and any(c['before'] != c['after'] for c in changes):
        raise ValueError('重编号可能影响有歧义的正文引用，请解除歧义后重新执行')
    for item in matches:
        target = options.get('position', {}).get(item['kind'])
        if target and target != item['position']:
            p = nodes[item['caption_body_index']]
            obj = nodes[item['object_body_index']]
            unsafe = {q('sectPr'),q('bookmarkStart'),q('bookmarkEnd'),q('commentRangeStart'),q('commentRangeEnd'),q('fldChar')}
            # A move may also cross a range beginning or ending outside this pair.
            if any(n.tag in unsafe for e in (obj,p) for n in e.iter()):
                raise ValueError('题注位置调整涉及书签、域或分节边界，需人工检查')
            (obj.addprevious if target == 'before' else obj.addnext)(p)
    configure(root, styles, settings, sorted(body.index(p) for p in selected))
    for cap in settings.find(q('captions')).findall(q('caption')):
        target = options.get('position',{}).get(cap.get(q('name')))
        if target:
            cap.set(q('pos'),'above' if target=='before' else 'below')
    return {'changes':changes, 'plain_references':refs, 'issues':issues}


def equation_width(root):
    sections = list(root.iter(q('sectPr')))
    if len(sections) != 1:
        return None
    sect = sections[0]
    cols = sect.find(q('cols'))
    if cols is not None and (cols.get(q('num'),'1') != '1' or len(cols)):
        return None
    page, margin = sect.find(q('pgSz')), sect.find(q('pgMar'))
    if page is None or margin is None:
        return None
    return int(page.get(q('w'))) - sum(int(margin.get(q(k),'0')) for k in ('left','right','gutter'))


def equations(stories, paragraphs, chapter_map, mode, sync, guard):
    root = stories['word/document.xml']
    # Keep original paragraph and OMML intact inside a borderless three-cell row.
    # Equal outer cells keep the centre independent of number width.
    width = equation_width(root)
    changes, issues, counters = [], [], Counter()
    mapping, selected = defaultdict(list), set()
    for i, p in enumerate(paragraphs):
        formulas = list(p.iter(MATH))
        if not formulas:
            continue
        table = next(p.iterancestors(q('tbl')),None)
        if table is not None:
            cells = table.findall('./'+q('tr')+'/'+q('tc'))
            number_p = cells[2].find(q('p'))
            fields, errors = scan(number_p)
            old = re.fullmatch(r'\((\d+(?:-\d+)?)\)',visible(number_p))
            if errors or len(fields) != 1 or not old or not fields[0].code.startswith('SEQ 公式'):
                raise ValueError('已有公式编号布局需要核对')
            prefix,digit,code = next_label('公式',mode,chapter_map[p],counters)
            f = fields[0]
            if not f.value.isdecimal() or not old[1].endswith(f.value):
                raise ValueError('公式编号域结果不一致')
            if f.code != code or old[1] != prefix+digit:
                if f.nested or f.owner.get(q('fldLock')) in ('1','true','on'):
                    raise ValueError('公式编号含嵌套或锁定的域')
                prefix_node = next(number_p.iter(q('t')))
                if prefix_node.text != '('+old[1][:-len(f.value)]:
                    raise ValueError('公式编号前缀需要核对')
                guard.replace(number_p,1,old[1],prefix+digit)
                prefix_node.text = '('+prefix
                f.set_code(code)
                f.set_value(digit)
            mapping['公式',old[1]].append(prefix+digit)
            changes.append({'paragraph':i,'old_number':old[1],'number':prefix+digit,'field':code})
            selected.add(number_p)
            continue
        old_literal = visible(p).strip()
        old = re.fullmatch(r'[（(](\d+(?:-\d+)?)[)）]',old_literal)
        ordinary = p.findall(q('r'))
        math_container = next((n for n in p if n.tag in (MATH,DISPLAY)),None)
        has_label = (old is not None and bool(ordinary) and math_container is not None and
                     all(p.index(r) > p.index(math_container) and
                         all(n.tag in {q('rPr'),q('t'),q('tab')} for n in r) for r in ordinary))
        allowed = {q('pPr'), MATH, DISPLAY}
        if has_label:
            allowed.add(q('r'))
        if (not width or width < 2000 or len(formulas) != 1 or any(n.tag not in allowed for n in p)
                or p.find('.//' + q('sectPr')) is not None or (old_literal and not has_label)):
            issues.append({'paragraph':i,'issue':'行内、混合内容或复杂页面布局，保留原公式'})
            continue
        prefix, digit, code = next_label('公式', mode, chapter_map[p], counters)
        if has_label:
            guard.replace(p,0,visible(p),'')
            mapping['公式',old[1]].append(prefix+digit)
            for r in ordinary:
                p.remove(r)
        table = element('tbl')
        props = ET.SubElement(table,q('tblPr'))
        props.append(element('tblW', w=width,type='dxa'))
        borders = ET.SubElement(props,q('tblBorders'))
        for side in ('top','left','bottom','right','insideH','insideV'):
            borders.append(element(side,val='nil'))
        props.append(element('tblLayout',type='fixed'))
        props.append(element('tblCaption',val='PaperFit:equation:v1'))
        grid = ET.SubElement(table,q('tblGrid'))
        row = ET.SubElement(table,q('tr'))
        side = width//7
        p.addprevious(table)
        for pos, cell_width in enumerate((side,width-2*side,side)):
            grid.append(element('gridCol',w=cell_width))
            cell = ET.SubElement(row,q('tc'))
            cp = ET.SubElement(cell,q('tcPr'))
            cp.append(element('tcW',w=cell_width,type='dxa'))
            cp.append(element('vAlign',val='center'))
            para = p if pos == 1 else element('p')
            cell.append(para)
            if pos != 1:
                guard.new(para, '('+prefix+digit+')' if pos==2 else '')
            if pos:
                pp = para.find(q('pPr'))
                if pp is None:
                    pp = element('pPr')
                    para.insert(0,pp)
                setv(pp,'jc',{'val':'center' if pos==1 else 'right'})
            if pos == 2:
                para.extend([run('('+prefix), seq_field(code,digit), run(')')])
        changes.append({'paragraph':i, 'old_number':old[1] if old else None, 'number':prefix+digit, 'field':code})
    equation_labels = set()
    for table in root.iter(q('tbl')):
        marker = table.find('./'+q('tblPr')+'/'+q('tblCaption'))
        if marker is not None and marker.get(q('val')) == 'PaperFit:equation:v1':
            equation_labels.update(table.findall('./'+q('tr')+'/'+q('tc')+'[3]/'+q('p')))
    guard_sequences(stories,equation_labels,{'公式'})
    processed = {m for p in paragraphs for m in p.iter(MATH)}
    other_math = sum(m not in processed for doc in stories.values() for m in doc.iter(MATH))
    if other_math:
        issues.append({'issue':'正文直属段落以外的公式已保留','count':other_math})
    refs, problems = sync_mentions(stories,selected,mapping,sync,guard,EQ_MENTION)
    if problems and any(x.get('old_number') and x['old_number'] != x['number'] for x in changes):
        raise ValueError('公式重编号影响未能唯一定位的引用')
    return {'changes':changes,'issues':issues+problems,'plain_references':refs}

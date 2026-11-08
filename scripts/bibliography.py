"""Format only reviewed bibliography entries. Keep bibliographic data literal."""
import re
from lxml import etree as ET
from paperfit import q, props
from structured import element, run, visible, numeric_node

LABEL = re.compile(r'^\s*\[(\d+)\]\s*')
CITATION = re.compile(r'\[(\d+(?:\s*[-–,，、]\s*\d+)*)\]')


def citation_numbers(value):
    numbers = []
    for group in re.split(r'[,，、]', value):
        bounds = re.split(r'[-–]',group)
        if len(bounds) == 1:
            numbers.append(int(bounds[0]))
        elif len(bounds) == 2 and 0 <= int(bounds[1]) - int(bounds[0]) <= 1000:
            numbers.extend(range(int(bounds[0]),int(bounds[1])+1))
        else:
            raise ValueError('文献引用范围需要核对：'+value)
    return numbers


def references(stories, styles, paragraphs, options, guard, native_labels=None):
    root = stories['word/document.xml']
    indices = options.get('entries')
    if not indices:
        raise ValueError('请用 entries 明确指定参考文献条目段落')
    selected = [paragraphs[i] for i in indices]
    if any(p.getparent() is not root.find(q('body')) for p in selected):
        raise ValueError('参考文献范围应为正文直属段落')
    changes, issues, old = [], [], []
    labels = []
    for i,p in zip(indices,selected):
        text = visible(p)
        if not text.strip() or text.strip() in ('参考文献','References') or any(n.tag not in {q('pPr'),q('r'),q('hyperlink'),q('bookmarkStart'),q('bookmarkEnd')} for n in p):
            raise ValueError(f'文献条目段落{i}结构需核实')
        old.append(text)
        match = LABEL.match(text)
        labels.append(int(match[1]) if match else native_labels[len(labels)] if native_labels else None)
    if len({n for n in labels if n is not None}) != sum(n is not None for n in labels):
        raise ValueError('文献编号重复，请先核对')
    renumber = options.get('number',False) and native_labels is None
    # Never silently change the identity of a cited item when remapping a list.
    shifted = {n for pos,n in enumerate(labels,1) if n is not None and n != pos}
    used = set()
    for p in (p for doc in stories.values() for p in doc.iter(q('p'))):
        if p in selected:
            continue
        for match in CITATION.finditer(visible(p)):
            used.update(citation_numbers(match[1]))
    if renumber and shifted & used:
        raise ValueError('重排文献会改变已有引用含义，请核对文献映射')
    number_set = set(range(1,len(selected)+1)) if renumber else {n for n in labels if n is not None}
    for n in sorted(used-number_set):
        issues.append({'citation':n,'issue':'正文引用未找到明确编号的文献'})
    sid = 'PaperFitBibliography'
    existing = next((s for s in styles if s.get(q('styleId')) == sid), None)
    if existing is not None and (existing.find(q('name')) is None or existing.find(q('name')).get(q('val')) != '参考文献条目'):
        raise ValueError('文献样式ID冲突')
    definitions = {s.get(q('styleId')):s for s in styles.findall(q('style'))}
    default_id = next((k for k,s in definitions.items() if s.get(q('type'))=='paragraph' and s.get(q('default'))=='1'),None)
    base_ids = set()
    for para in selected:
        st = para.find('./'+q('pPr')+'/'+q('pStyle'))
        source_id = st.get(q('val')) if st is not None else default_id
        if source_id != sid:
            base_ids.add(source_id)
    style = existing
    if len(base_ids) > 1:
        raise ValueError('文献条目继承不同的样式，请先确认共同的基础样式')
    if style is None:
        style = element('style',type='paragraph',styleId=sid,customStyle='1')
        style.append(element('name',val='参考文献条目'))
        if base_ids and None not in base_ids:
            style.append(element('basedOn',val=next(iter(base_ids))))
        style.append(element('qFormat'))
        styles.append(style)
    elif base_ids:
        base = style.find(q('basedOn'))
        if base_ids != {base.get(q('val')) if base is not None else None}:
            raise ValueError('已有文献样式与新条目的基础样式不同')
    inherited = []
    base = style.find(q('basedOn'))
    visited = {sid}
    while base is not None:
        key = base.get(q('val'))
        if key in visited or key not in definitions:
            raise ValueError('文献样式的继承关系需要核对')
        visited.add(key)
        parent = definitions[key]
        inherited.append(parent)
        if renumber and parent.find('./'+q('pPr')+'/'+q('numPr')) is not None:
            raise ValueError('文献条目继承Word自动编号，请核对映射')
        base = parent.find(q('basedOn'))
    formatting = options.get('style',{})
    pp, rp = props(style,formatting)
    if 'hanging_chars' in options:
        count = options['hanging_chars']
        indent = pp.find(q('ind'))
        if indent is None:
            indent = ET.SubElement(pp,q('ind'))
        for key in ('leftChars','hangingChars'):
            indent.set(q(key),str(round(count*100)))
        font_size = style.find('./'+q('rPr')+'/'+q('sz'))
        for ancestor in inherited:
            if font_size is None:
                font_size = ancestor.find('./'+q('rPr')+'/'+q('sz'))
        if font_size is None:
            font_size = styles.find('./'+q('docDefaults')+'/'+q('rPrDefault')+'/'+q('rPr')+'/'+q('sz'))
        if font_size is not None:
            for key in ('left','hanging'):
                indent.set(q(key),str(round(count*int(font_size.get(q('val')))*10)))
    for i,(para,old_text) in enumerate(zip(selected,old)):
        idx = indices[i]
        ppr = para.find(q('pPr'))
        if ppr is None:
            ppr = element('pPr')
            para.insert(0,ppr)
        if ppr.find(q('numPr')) is not None and renumber:
            raise ValueError(f'段落{idx}使用Word列表编号，需核对列表映射')
        if renumber and labels[i] is None and re.match(r'^\s*(?:\d+[.、)]|［\d+］)',old_text):
            raise ValueError(f'段落{idx}含其他格式的已有编号')
        if ppr.find(q('sectPr')) is not None or any(n.tag in {q('fldChar'),q('instrText'),q('ins'),q('del'),q('drawing')} for n in para.iter()):
            raise ValueError(f'段落{idx}含复杂对象，请核对')
        st = ppr.find(q('pStyle'))
        if st is None:
            st = element('pStyle')
            ppr.insert(0,st)
        st.set(q('val'),sid)
        if 'hanging_chars' in options:
            direct = ppr.find(q('ind'))
            if direct is not None:
                for k in ('left','start','leftChars','startChars','firstLine','firstLineChars','hanging','hangingChars'):
                    direct.attrib.pop(q(k),None)
        if 'line' in formatting:
            spacing = ppr.find(q('spacing'))
            if spacing is not None:
                for k in ('line','lineRule'):
                    spacing.attrib.pop(q(k),None)
            snap = ppr.find(q('snapToGrid'))
            if snap is not None:
                ppr.remove(snap)
        for direct in para.iter(q('rPr')):
            for item in list(direct):
                if rp.find(item.tag) is not None:
                    direct.remove(item)
        expected = old_text
        if renumber:
            match = LABEL.match(old_text)
            if not match:
                para.insert(1,run(f'[{i+1}] '))
                expected = f'[{i+1}] '+old_text
            elif match[1] != str(i+1):
                node,left,right = numeric_node(para,*match.span(1))
                node.text = node.text[:left]+str(i+1)+node.text[right:]
                expected = old_text[:match.start(1)]+str(i+1)+old_text[match.end(1):]
        guard.replace(para,0,old_text,expected)
        if visible(para) != expected:
            raise ValueError('文献数据完整性检查失败')
        if not re.search(r'(?<!\d)(?:18|19|20)\d{2}(?!\d)', old_text):
            issues.append({'paragraph':idx,'issue':'未识别年份，请核实原条目'})
        changes.append({'paragraph':idx,'old_number':labels[i],'new_number':i+1 if renumber else labels[i], 'style':sid})
    return {'changes':changes,'issues':issues,'cited':sorted(used),'uncited':sorted(number_set-used)}

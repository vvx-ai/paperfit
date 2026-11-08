"""Compare effective, requested style properties on reviewed paragraphs."""
from paperfit import q, in_math


def audit_styles(paragraphs, styles, roles, requested):
    definitions = {n.get(q('styleId')):n for n in styles.findall(q('style'))}
    defaults = styles.find(q('docDefaults'))
    records, issues = [], []

    def chain(sid):
        result, seen = [], set()
        while sid in definitions and sid not in seen:
            seen.add(sid)
            style = definitions[sid]
            result.append(style)
            parent = style.find(q('basedOn'))
            sid = parent.get(q('val')) if parent is not None else None
        return list(reversed(result))

    def effective(layers, name, attribute='val'):
        value = None
        for props in layers:
            node = props.find(q(name)) if props is not None else None
            if node is not None:
                value = node.get(q(attribute),'1' if name == 'b' else value)
        return value

    def matches(value,expected):
        if type(expected) is bool:
            return (value not in (None,'0','off','false')) == expected
        return value == str(expected)

    for i,p in enumerate(paragraphs):
        role = roles.get(str(i))
        if role not in requested:
            continue
        rule = requested[role]
        pp = p.find(q('pPr'))
        ref = pp.find(q('pStyle')) if pp is not None else None
        sid = ref.get(q('val')) if ref is not None else None
        ancestors = chain(sid)
        style_name = None
        if sid in definitions:
            name = definitions[sid].find(q('name'))
            style_name = name.get(q('val')) if name is not None else None
        path = lambda tag:[defaults.find('./'+q(tag+'Default')+'/'+q(tag)) if defaults is not None else None] + [s.find(q(tag)) for s in ancestors]
        p_layers = path('pPr')+[pp]
        expected = []
        if 'line' in rule:
            expected.extend([('spacing','line',round(rule['line']*240)),('spacing','lineRule','auto')])
        if 'indent_chars' in rule:
            expected.append(('ind','firstLineChars',round(rule['indent_chars']*100)))
        if 'align' in rule:
            expected.append(('jc','val',rule['align']))
        for k in ('before','after'):
            if k in rule:
                expected.append(('spacing',k,round(rule[k]*20)))
        if role.startswith('heading'):
            expected.append(('outlineLvl','val',int(role[-1])-1))
        problems, overrides = set(), set()
        for tag,attr,target in expected:
            if pp is not None and pp.find(q(tag)) is not None and pp.find(q(tag)).get(q(attr)) is not None:
                overrides.add(tag+'.'+attr)
            if not matches(effective(p_layers,tag,attr),target):
                problems.add(tag+'.'+attr)
        if not ancestors:
            problems.add('缺少明确段落样式')
        for r in p.iter(q('r')):
            if in_math(r) or r.find(q('t')) is None:
                continue
            rp = r.find(q('rPr'))
            char = rp.find(q('rStyle')) if rp is not None else None
            layers = path('rPr')
            if char is not None:
                layers.extend(s.find(q('rPr')) for s in chain(char.get(q('val'))))
            layers.append(rp)
            for k,tag,attr in (('font','rFonts','eastAsia'),('font','rFonts','ascii'),('size','sz','val'),('bold','b','val')):
                if k in rule:
                    value = round(rule[k]*2) if k=='size' else rule[k]
                    if rp is not None and rp.find(q(tag)) is not None:
                        overrides.add(tag)
                    if not matches(effective(layers,tag,attr),value):
                        problems.add(tag+'.'+attr)
        record = {'paragraph':i,'role':role,'style_id':sid,'name':style_name,'direct_overrides':sorted(overrides),'differences':sorted(problems)}
        records.append(record)
        if problems:
            issues.append(record)
    return {'paragraphs':records,'issues':issues}

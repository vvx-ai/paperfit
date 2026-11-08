"""Inspect native Word fields and safely refresh plain REF display caches."""
from collections import defaultdict
from dataclasses import dataclass, field
import shlex
from paperfit import q


@dataclass
class Field:
    owner: object
    simple: bool = False
    instructions: list = field(default_factory=list)
    result: list = field(default_factory=list)
    separated: bool = False
    nested: bool = False

    @property
    def code(self):
        if self.simple:
            return self.owner.get(q('instr'), '').strip()
        return ''.join(n.text or '' for n in self.instructions).strip()

    @property
    def value(self):
        return ''.join(n.text or '' for n in self.result)

    def set_code(self, code):
        if self.simple:
            self.owner.set(q('instr'), ' ' + code + ' ')
        else:
            self.instructions[0].text = ' ' + code + ' '
            for node in self.instructions[1:]:
                node.text = ''

    def set_value(self, value):
        if self.nested or len(self.result) != 1:
            raise ValueError('域结果跨节点或嵌套，需要人工检查')
        self.result[0].text = value


def scan(root):
    fields, stack, problems = [], [], []
    for node in root.iter():
        if node.tag == q('fldSimple'):
            item = Field(node, simple=True, result=list(node.iter(q('t'))))
            item.nested = bool(stack or list(node.iterancestors(q('fldSimple'))) or node.find('.//' + q('fldSimple')) is not None or node.find('.//' + q('fldChar')) is not None)
            if stack:
                stack[-1].nested = True
            fields.append(item)
        elif node.tag == q('fldChar'):
            kind = node.get(q('fldCharType'))
            if kind == 'begin':
                if stack:
                    stack[-1].nested = True
                item = Field(node, nested=bool(stack or list(node.iterancestors(q('fldSimple')))))
                stack.append(item)
                fields.append(item)
            elif not stack:
                problems.append('域缺少起始标记')
            elif kind == 'separate':
                stack[-1].separated = True
            elif kind == 'end':
                stack.pop()
        elif stack and node.tag == q('instrText'):
            stack[-1].instructions.append(node)
        elif stack and stack[-1].separated and node.tag == q('t'):
            stack[-1].result.append(node)
    if stack:
        problems.append('域缺少结束标记')
    return fields, problems


def bookmark_ranges(root):
    opened, named = {}, defaultdict(list)
    nodes = list(root.iter())
    for index, node in enumerate(nodes):
        key = node.get(q('id'))
        if node.tag == q('bookmarkStart'):
            opened[key] = (node.get(q('name'), ''), index)
        elif node.tag == q('bookmarkEnd') and key in opened:
            name, start = opened.pop(key)
            named[name].append(nodes[start + 1:index])
    for name, _ in opened.values():
        named[name].append(None)
    return named


def audit(stories, sync=False, guard=None):
    bookmarks, fields, issues, refs, seqs = defaultdict(list), [], [], [], []
    for part, root in stories.items():
        for name, ranges in bookmark_ranges(root).items():
            bookmarks[name].extend(ranges)
        found, errors = scan(root)
        fields.extend((part, f) for f in found)
        issues.extend({'part': part, 'issue': e} for e in errors)
    for part, f in fields:
        try:
            tokens = [s.strip('"') for s in shlex.split(f.code, posix=False)]
        except ValueError:
            issues.append({'part': part, 'issue': '无法解析域指令', 'code': f.code})
            continue
        kind = tokens[0].upper() if tokens else ''
        if kind == 'SEQ':
            seqs.append({'part': part, 'code': f.code, 'value': f.value})
        if kind not in ('REF', 'PAGEREF', 'NOTEREF'):
            continue
        name = tokens[1] if len(tokens) > 1 else ''
        ranges = bookmarks.get(name, [])
        record = {'part': part, 'field': f.code, 'bookmark': name, 'cached': f.value}
        refs.append(record)
        if len(ranges) != 1 or ranges[0] is None:
            record['status'] = '目标缺失、重复或未闭合'
        elif kind == 'REF' and name.startswith('_PaperFitBib') and '\\n' in tokens[2:] and all(t.lower() in ('\\n','\\h','\\*','mergeformat') for t in tokens[2:]):
            record['status'] = '自动编号目标存在，显示值由Word计算'
        elif kind != 'REF' or any(t.lower() != '\\h' for t in tokens[2:]):
            record['status'] = '目标存在，显示值需由Word计算'
        else:
            nodes = ranges[0]
            value = ''.join(n.text or '' for n in nodes if n.tag == q('t'))
            if not value or any(n.tag in {q('p'), q('br'), q('tab'), q('drawing'), q('sym'), q('ins'), q('del')} for n in nodes):
                record['status'] = '复杂书签范围，需由Word计算'
            elif value == f.value:
                record['status'] = '一致'
            elif sync and not f.nested and len(f.result) == 1 and f.owner.get(q('fldLock')) not in ('1','true','on') and not any(n.tag in {q('ins'), q('del')} for n in f.owner.iterancestors()):
                if guard is not None:
                    result = f.result[0]
                    p = next(result.iterancestors(q('p')))
                    nodes = list(p.iter(q('t')))
                    offset = sum(len(n.text or '') for n in nodes[:nodes.index(result)])
                    guard.replace(p,offset,f.value,value)
                f.set_value(value)
                record.update(status='已同步', updated=value)
            else:
                record.update(status='显示值待更新', expected=value)
        if record['status'] not in ('一致','已同步','自动编号目标存在，显示值由Word计算'):
            issues.append({'part': part, 'issue': record['status'], 'bookmark': name})
    return {'references': refs, 'sequences': seqs, 'issues': issues}

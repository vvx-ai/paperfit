"""Read-only structural comparison, retaining source positions on both sides."""
from difflib import SequenceMatcher
from hashlib import sha256
from pathlib import Path

from lxml import etree as ET
from phase2 import DOC, STYLES, xml, canonical, read_package
from paperfit import q
from toc import val
from structured import visible


def describe(node, index):
    return {'body_index':index,'kind':ET.QName(node).localname,'text':visible(node)[:100],
            'paragraph_style':val(node,'pPr/pStyle'),'xml_sha256':sha256(canonical(node)).hexdigest()}


def compare(before, after):
    before, after = Path(before), Path(after)
    left, right = read_package(before), read_package(after)
    a, b = [list(xml(p[DOC]).find(q('body'))) for p in (left,right)]
    sig_a, sig_b = [[sha256(canonical(n)).digest() for n in nodes] for nodes in (a,b)]
    blocks = []
    for op,i,j,k,m in SequenceMatcher(None,sig_a,sig_b,autojunk=False).get_opcodes():
        if op != 'equal':
            blocks.append({'action':op,'before':[describe(a[n],n) for n in range(i,j)],
                           'after':[describe(b[n],n) for n in range(k,m)]})
    styles = []
    defs = [{s.get(q('styleId')):s for s in xml(p[STYLES]).findall(q('style'))} for p in (left,right)]
    for sid in sorted(defs[0].keys() | defs[1].keys()):
        old,new = [group.get(sid) for group in defs]
        if old is not None and new is not None and canonical(old)==canonical(new):
            continue
        styles.append({'id':sid,'before_name':val(old,'name') if old is not None else None,
                       'after_name':val(new,'name') if new is not None else None})
    text_same = [visible(n) for n in a] == [visible(n) for n in b]
    parts = {k:('added' if k not in left else 'removed' if k not in right else 'modified')
             for k in sorted(left.keys() | right.keys()) if left.get(k) != right.get(k)}
    return {'before_sha256':sha256(before.read_bytes()).hexdigest(),'after_sha256':sha256(after.read_bytes()).hexdigest(),
            'body_block_text_equal':text_same,'body_changes':blocks,'style_changes':styles,'package_changes':parts}

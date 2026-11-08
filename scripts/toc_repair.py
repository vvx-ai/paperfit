"""Recover a broken TOC link from an independently verified PAGEREF target."""
import re
from collections import defaultdict
from paperfit import q
from structured import visible
from toc import heading_level, val
from word_fields import scan


def managed_links(root):
    blocks = [s for s in root.find(q('body')).findall(q('sdt')) if val(s,'sdtPr/tag')=='PaperFitTOC']
    if len(blocks)!=1 or scan(blocks[0])[1]:
        raise ValueError('需要唯一且域边界完整的 PaperFit 目录')
    return list(blocks[0].iter(q('hyperlink')))


def repair_links(root,styles):
    links = managed_links(root)
    body = root.find(q('body'))
    bookmarks, ends = defaultdict(list), defaultdict(list)
    for b in body.iter(q('bookmarkStart')):
        bookmarks[b.get(q('name'))].append(b)
    for b in body.iter(q('bookmarkEnd')):
        ends[b.get(q('id'))].append(b)
    repairs, pending = [], []
    for index,link in enumerate(links):
        fields, errors = scan(link)
        old = link.get(q('anchor'),'')
        if len(fields)!=1 or errors:
            pending.append({'link':index,'reason':'页码域无法唯一识别'})
            continue
        ref = fields[0]
        match = re.fullmatch(r'PAGEREF\s+([\w]+)(?:\s+\\h)?',ref.code,re.I)
        if not match:
            pending.append({'link':index,'reason':'页码域格式需核对'})
            continue
        target = match[1]
        candidates = bookmarks.get(target,[])
        if len(candidates)!=1:
            pending.append({'link':index,'reason':'页码书签缺失或重复'})
            continue
        bookmark = candidates[0]
        p = bookmark.getparent()
        close = ends.get(bookmark.get(q('id')),[])
        if (p.tag!=q('p') or p.getparent() is not body or not heading_level(p,styles)
                or len(close)!=1 or close[0].getparent() is not p or p.index(close[0])<=p.index(bookmark)
                or visible(link)!=visible(p).strip()+ref.value):
            pending.append({'link':index,'reason':'标题文字或书签范围需核对'})
            continue
        if any(key in link.attrib for key in ('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',q('docLocation'))):
            pending.append({'link':index,'reason':'外部或特殊链接需核对'})
        elif old==target:
            continue
        elif bookmarks.get(old):
            pending.append({'link':index,'reason':'链接与页码分别指向现有目标，请核对'})
        else:
            link.set(q('anchor'),target)
            repairs.append({'link':index,'heading':visible(p),'before':old,'after':target})
    return {'repairs':repairs,'pending':pending,'status':'needs_review' if pending else 'passed'}

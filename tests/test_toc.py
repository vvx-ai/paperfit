import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from phase2 import DOC, STYLES, canonical, read_package, xml
from paperfit import q
from structured import element
from toc import add_toc, val
from word_fields import scan


class TocTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.source = Path(tmp.name) / 'source.docx'
        self.output = Path(tmp.name) / 'output.docx'

    def test_native_levels_links_and_integrity(self):
        doc = Document()
        doc.add_heading('总标题', 0)
        doc.add_heading('章', 1)
        style = doc.styles.add_style('自定义子节', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles['Heading 2']
        doc.add_paragraph('节', style=style)
        doc.add_heading('小节', 3)
        doc.add_heading('更深一级', 4)
        p = doc.add_paragraph('正文')
        p._p.extend([element('bookmarkStart', id=4, name='_PaperFitToc5'), element('bookmarkEnd', id=4)])
        doc.add_table(1, 1).cell(0, 0).text = '数据'
        doc.save(self.source)
        original = read_package(self.source)
        report = add_toc(self.source, self.output)
        self.assertEqual([e['level'] for e in report['entries']], [1, 2, 3])
        result = read_package(self.output)
        for part in original.keys() - {DOC, STYLES}:
            self.assertEqual(original[part], result[part])
        root = xml(result[DOC])
        block = root.find('.//' + q('sdt'))
        fields, issues = scan(block)
        self.assertFalse(issues)
        self.assertEqual([f.code.split()[0] for f in fields], ['TOC', 'PAGEREF', 'PAGEREF', 'PAGEREF'])
        for link, entry in zip(block.iter(q('hyperlink')), report['entries']):
            anchor = link.get(q('anchor'))
            matches = [b for b in root.iter(q('bookmarkStart')) if b.get(q('name')) == anchor]
            self.assertEqual(len(matches), 1)
            self.assertEqual(''.join(matches[0].getparent().itertext(q('t'))), entry['heading'])
            self.assertEqual(val(link.getparent(), 'pPr/pStyle'), f"TOC{entry['level']}")
            parent = matches[0].getparent()
            end = next(b for b in parent.iter(q('bookmarkEnd')) if b.get(q('id')) == matches[0].get(q('id')))
            parent.remove(matches[0])
            parent.remove(end)
        block.getparent().remove(block)
        self.assertEqual(canonical(root), canonical(xml(original[DOC])))
        with self.assertRaisesRegex(ValueError, '已有目录域'):
            add_toc(self.output, self.output.with_name('duplicate.docx'))
        with self.assertRaises(ValueError):
            add_toc(self.source, self.output)

    def test_manual_toc_missing_headings_and_sections(self):
        for case in ('目 录', '普通正文', '分节'):
            doc = Document()
            doc.add_paragraph(case)
            if case == '分节':
                doc.add_heading('章节', 1)
                doc.add_section(WD_SECTION.NEW_PAGE)
            doc.save(self.source)
            with self.subTest(case=case), self.assertRaises(ValueError):
                add_toc(self.source, self.output)
            self.assertFalse(self.output.exists())

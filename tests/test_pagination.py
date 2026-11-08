from pathlib import Path
import sys
import tempfile
import unittest

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from pagination import DOC, REL, RELS, TYPES, paginate, read_package, xml
from paperfit import q
from phase2 import canonical, save_package
from structured import run
from toc import add_toc, val
from word_fields import scan


class PaginationTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        self.toc = self.folder / 'toc.docx'
        self.out = self.folder / 'paginated.docx'

    def fixture(self, footer=False):
        doc = Document()
        doc.sections[0].header.paragraphs[0].text = '校名'
        if footer:
            doc.sections[0].footer.paragraphs[0].text = '已有文字'
        doc.add_heading('第一章', 1)
        doc.add_paragraph('原始内容')
        doc.add_heading('原生样式子节', 2)
        source = self.folder / 'source.docx'
        doc.save(source)
        add_toc(source, self.toc)

    def test_split_page_restart_and_preservation(self):
        self.fixture()
        old = read_package(self.toc)
        paginate(self.toc, self.out)
        new = read_package(self.out)
        root = xml(new[DOC])
        block = root.find('.//' + q('sdt'))
        front, main = list(root.iter(q('sectPr')))
        self.assertIs(front.getparent().getparent(), block.getnext())
        self.assertIsNone(block.find('.//' + q('br')))
        self.assertIsNone(front.find(q('footerReference')))
        self.assertEqual(val(main, 'type'), 'nextPage')
        self.assertEqual(main.find(q('pgNumType')).attrib, {q('start'): '1', q('fmt'): 'decimal'})
        self.assertEqual(canonical(front), canonical(xml(old[DOC]).find('.//' + q('sectPr'))))
        ref = main.find(q('footerReference')).get('{'+REL+'}id')
        target, = [r.get('Target') for r in xml(new[RELS]) if r.get('Id') == ref and r.get('Type') == REL+'/footer']
        footer = xml(new['word/'+target])
        self.assertEqual(val(footer, 'p/pPr/jc'), 'center')
        fields, errors = scan(footer)
        self.assertEqual(([f.code for f in fields], errors), (['PAGE'], []))
        self.assertTrue(any(n.get('PartName') == '/word/'+target and n.get('ContentType').endswith('footer+xml') for n in xml(new[TYPES])))
        for name in old.keys() - {DOC, RELS, TYPES}:
            self.assertEqual(old[name], new[name])
        self.assertEqual(read_package(self.toc), old)
        with self.assertRaisesRegex(ValueError, '已有分节'):
            paginate(self.out, self.folder / 'again.docx')
        with self.assertRaises(ValueError):
            paginate(self.toc, self.out)

    def test_existing_footer_and_modified_break(self):
        self.fixture(footer=True)
        with self.assertRaisesRegex(ValueError, '已有页脚'):
            paginate(self.toc, self.out)
        parts = read_package(self.toc)
        root = xml(parts[DOC])
        section = root.find('.//' + q('sectPr'))
        section.remove(section.find(q('footerReference')))
        root.find('.//' + q('sdtContent'))[-1].append(run('要保留的文字'))
        edited = self.folder / 'edited.docx'
        save_package(parts, {DOC: root}, edited)
        with self.assertRaisesRegex(ValueError, '已被编辑'):
            paginate(edited, self.out)
        self.assertFalse(self.out.exists())
        self.assertFalse(Path(str(self.out)+'.report.json').exists())

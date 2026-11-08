import contextlib
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt
from lxml import etree as ET

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from phase2 import DOC, STYLES, read_package, save_package, xml, canonical
from phase3 import main, sections, section_values, typography, definitions, layers_for
from paperfit import q
from structured import element, visible
from toc import add_toc, val
from toc_repair import managed_links
from word_fields import scan


class AdvancedFormatting(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        self.source = self.folder/'source.docx'
        self.doc = Document()
        self.doc.add_heading('一',1)
        self.doc.add_heading('同级甲',2)
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        sec.header.is_linked_to_previous = False
        sec.header.paragraphs[0].text = '第二节专用页眉'
        sec.footer.paragraphs[0].text = '页脚'
        self.doc.add_heading('二',1)
        self.target = self.doc.add_heading('同级乙',2)
        self.run = self.target.runs[0]
        self.run.italic = True
        self.run.font.superscript = True
        self.run.font.name = 'Override'
        self.doc.add_heading('下一级',3)
        self.doc.add_table(1,1).cell(0,0).text = '数据'
        self.doc.add_heading('三',1)
        self.doc.add_heading('同级丙',2)
        self.doc.save(self.source)

    def invoke(self, command, *args):
        main([command,*map(str,args)])

    def chapter(self, source, out, *options):
        self.invoke('chapter',source,out,'--chapter','二','--level',2,*options)

    def test_scoped_edit_isolated_repair_repeated_and_diff(self):
        out = self.folder/'one.docx'
        opts = ('--font','黑体','--size',14)
        self.chapter(self.source,out,*opts)
        old, new = read_package(self.source),read_package(out)
        for part in old.keys()-{DOC, STYLES}:
            self.assertEqual(old[part],new[part])
        b0,b1 = [xml(parts[DOC]).find(q('body')) for parts in (old,new)]
        changed = [(x,y) for x,y in zip(b0,b1) if canonical(x)!=canonical(y)]
        self.assertEqual(len(changed),1)
        x,y = changed[0]
        self.assertEqual(visible(y),'同级乙')
        self.assertEqual(val(y,'r/rPr/vertAlign'),'superscript')
        self.assertIsNotNone(y.find('./'+q('r')+'/'+q('rPr')+'/'+q('i')))
        self.assertIsNone(y.find('.//'+q('rFonts')))
        sty = xml(new[STYLES])
        sid = val(y,'pPr/pStyle')
        defs = definitions(sty)
        self.assertEqual(val(defs[sid],'basedOn'),val(x,'pPr/pStyle'))
        self.assertEqual(typography(layers_for(sty,defs,sid))['fonts']['eastAsia'],'黑体')
        again = self.folder/'again.docx'
        self.chapter(out,again,*opts)
        self.assertEqual(new,read_package(again))
        diff = self.folder/'diff.json'
        self.invoke('compare',self.source,out,diff)
        data = json.loads(diff.read_text())
        self.assertTrue(data['body_block_text_equal'])
        self.assertEqual(len(data['body_changes']),1)
        self.assertEqual(len(data['style_changes']),1)
        self.assertEqual(set(data['package_changes']),{DOC,STYLES})
        with self.assertRaises(SystemExit), contextlib.redirect_stderr(io.StringIO()):
            self.chapter(self.source,out,*opts)

    def test_inherited_template_style_and_section_mapping(self):
        tpl = Document()
        base = tpl.styles.add_style('学校基样式',WD_STYLE_TYPE.PARAGRAPH)
        rp = base._element.get_or_add_rPr()
        rp.append(element('rFonts',ascii='Arial',hAnsi='Arial',eastAsia='黑体',cs='Arial'))
        rp.append(element('sz',val='30'))
        base.paragraph_format.space_before = Pt(9)
        base.paragraph_format.line_spacing = 1.5
        tpl.styles.add_style('学校二级标题',WD_STYLE_TYPE.PARAGRAPH).base_style = base
        sec = tpl.add_section(WD_SECTION.NEW_PAGE)
        sec.left_margin,sec.page_width = Cm(3.1),Cm(26)
        template = self.folder/'template.docx'
        tpl.save(template)
        out, shaped = self.folder/'styled.docx', self.folder/'mapped.docx'
        self.chapter(self.source,out,'--template',template,'--style','学校二级标题')
        report = json.loads(Path(str(out)+'.report.json').read_text())
        style = report['changes'][0]['requested_format']
        self.assertEqual(style['half_points']['sz'],30)
        self.assertEqual(style['fonts']['ascii'],'Arial')
        self.assertEqual(style['paragraph']['spacing']['before'],'180')
        self.assertEqual(style['paragraph']['spacing']['line'],'360')
        repeat = self.folder/'template-repeat.docx'
        self.chapter(out,repeat,'--template',template,'--style','学校二级标题')
        self.assertEqual(read_package(out),read_package(repeat))
        self.invoke('layout',out,shaped,'--section',2,'--template',template,'--template-section',2)
        audit = self.folder/'info.json'
        self.invoke('template-info',template,audit)
        source, new = read_package(out),read_package(shaped)
        self.assertEqual([k for k in source if source[k]!=new[k]],[DOC])
        def get_sections(parts):
            return sections(xml(parts[DOC]).find(q('body')))
        self.assertEqual(canonical(get_sections(source)[0]),canonical(get_sections(new)[0]))
        self.assertEqual(section_values(get_sections(new)[1]),json.loads(audit.read_text())['sections'][1]['layout'])

    def test_reject_ambiguous_chapter_and_unsupported_template(self):
        out = self.folder/'no.docx'
        self.doc.add_heading('二',1)
        self.doc.save(self.source)
        with self.assertRaises(SystemExit),contextlib.redirect_stderr(io.StringIO()):
            self.chapter(self.source,out,'--font','黑体','--size',14)
        self.assertFalse(out.exists())
        with self.assertRaises(SystemExit),contextlib.redirect_stderr(io.StringIO()):
            self.chapter(self.source,out,'--template',self.source,'--style','Missing')
        self.assertFalse(Path(str(out)+'.report.json').exists())

    def test_revision_and_inherited_style_cycle_are_reported(self):
        parts = read_package(self.source)
        styles = xml(parts[STYLES])
        heading = definitions(styles)['Heading2']
        heading.find(q('basedOn')).set(q('val'),'Heading2')
        cyclic = self.folder/'cycle.docx'
        save_package(parts,{STYLES:styles},cyclic)
        with self.assertRaises(SystemExit),contextlib.redirect_stderr(io.StringIO()):
            self.chapter(cyclic,self.folder/'invalid.docx','--font','黑体','--size',14)
        root = xml(parts[DOC])
        root.find('.//'+q('p')).append(element('ins',id=1))
        revised = self.folder/'revised.docx'
        save_package(parts,{DOC:root},revised)
        with self.assertRaises(SystemExit),contextlib.redirect_stderr(io.StringIO()):
            self.chapter(revised,self.folder/'invalid.docx','--font','黑体','--size',14)
        self.assertFalse((self.folder/'invalid.docx').exists())

    def test_repair_only_a_uniquely_corroborated_broken_link(self):
        doc = Document()
        doc.add_heading('章一',1)
        doc.add_heading('小节',2)
        doc.save(self.source)
        toc = self.folder/'toc.docx'
        add_toc(self.source,toc)
        parts = read_package(toc)
        root = xml(parts[DOC])
        links = managed_links(root)
        expected = links[0].get(q('anchor'))
        links[0].set(q('anchor'),'Missing')
        scan(links[1])[0][0].set_code(' PAGEREF OtherMissing ')
        broken, repaired = self.folder/'broken.docx', self.folder/'fixed.docx'
        save_package(parts,{DOC:root},broken)
        self.invoke('repair-toc',broken,repaired)
        after = managed_links(xml(read_package(repaired)[DOC]))
        self.assertEqual(after[0].get(q('anchor')),expected)
        self.assertEqual(canonical(after[1]),canonical(links[1]))
        report = json.loads(Path(str(repaired)+'.report.json').read_text())
        self.assertEqual(report['status'],'needs_review')
        self.assertEqual(len(report['pending']),1)
        self.assertEqual(len(report['repairs']),1)

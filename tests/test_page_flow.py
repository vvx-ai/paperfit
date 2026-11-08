import base64
from pathlib import Path
import sys
import tempfile
import unittest

from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Inches

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from page_flow import enforce, sections, content_signature, WP, REL
from phase2 import DOC, STYLES, read_package, xml, save_package
from paperfit import q
from structured import element
from toc import val


class PageFlow(unittest.TestCase):
    def setUp(self):
        tmp=tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.folder=Path(tmp.name)
        self.source=self.folder/'source.docx'
        self.output=self.folder/'out.docx'
        self.image=self.folder/'dot.png'
        self.image.write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jh1sAAAAASUVORK5CYII='))

    def test_chapters_keep_continuous_pages_and_independent_stories(self):
        doc=Document()
        doc.add_paragraph('前置内容')
        doc.sections[0].header.paragraphs[0].text='页眉'
        doc.sections[0].footer.paragraphs[0].text='页脚'
        section=doc.add_section(WD_SECTION.NEW_PAGE)
        section._sectPr.append(element('pgNumType',start=1))
        doc.add_heading('第一章',1)
        doc.add_paragraph('正文')
        doc.add_heading('第二章',1)
        doc.add_heading('二级',2)
        doc.add_paragraph('正文二')
        doc.add_heading('第三章',1)
        doc.save(self.source)
        enforce(self.source,self.output)
        a,b=read_package(self.source),read_package(self.output)
        root=xml(b[DOC]); secs=sections(root.find(q('body')))
        self.assertEqual(len(secs),4)
        self.assertEqual([s.find(q('pgNumType')).get(q('start')) if s.find(q('pgNumType')) is not None else None for s in secs],[None,'1',None,None])
        self.assertEqual([val(s,'type') for s in secs[1:]],['nextPage']*3)
        ids=[r.get('{'+REL+'}id') for s in secs for r in s if r.tag in (q('headerReference'),q('footerReference'))]
        self.assertEqual(len(ids),24)
        self.assertEqual(len(set(ids)),24)
        self.assertEqual(content_signature(xml(a[DOC])),content_signature(root))
        styles={s.get(q('styleId')):s for s in xml(b[STYLES]).findall(q('style'))}
        self.assertEqual(val(styles['Heading1'],'pPr/pageBreakBefore'),'1')
        self.assertEqual(val(styles['Heading2'],'pPr/keepNext'),'1')
        again=self.folder/'repeat.docx'
        report=enforce(self.output,again)
        self.assertEqual(report['changed_parts'],[])
        self.assertEqual(b,read_package(again))

    def test_image_both_caption_positions_and_table_chain(self):
        doc=Document()
        doc.add_heading('章',1)
        doc.add_picture(str(self.image),height=Inches(12),width=Inches(4))
        doc.add_paragraph('图1 大图',style='Caption')
        doc.add_paragraph('间隔')
        doc.add_paragraph('图2 图上题注',style='Caption')
        doc.add_picture(str(self.image),height=Inches(1),width=Inches(1))
        doc.add_paragraph('间隔')
        table=doc.add_table(2,2)
        table.cell(0,0).text='保留表格文字'
        doc.add_paragraph('表1',style='Caption')
        doc.save(self.source)
        report=enforce(self.source,self.output)
        self.assertEqual(report['pending'],[])
        self.assertEqual(len(report['captions']),3)
        resize=report['captions'][0]['image_resize']
        self.assertLess(resize['after_emu']['cy'],resize['before_emu']['cy'])
        self.assertAlmostEqual(resize['after_emu']['cy']/resize['after_emu']['cx'],3,places=5)
        parts=read_package(self.output); root=xml(parts[DOC])
        self.assertEqual(val(xml(parts['word/settings.xml']),'doNotAutoCompressPictures'),'1')
        body=root.find(q('body'))
        captions=[p for p in body.findall(q('p')) if val(p,'pPr/pStyle')=='Caption']
        self.assertEqual([val(p,'pPr/keepNext') for p in captions],['0','1','0'])
        self.assertEqual(val(captions[0].getprevious(),'pPr/keepNext'),'1')
        self.assertEqual(val(captions[1].getnext(),'pPr/keepNext'),'0')
        self.assertTrue(all(val(p,'pPr/keepNext')=='1' for p in root.find('.//'+q('tbl')).iter(q('p'))))
        self.assertEqual(len(list(root.iter(q('cantSplit')))),2)
        source=read_package(self.source)
        for name in source:
            if name.startswith('word/media/'):
                self.assertEqual(source[name],parts[name])
        self.assertEqual(content_signature(xml(source[DOC])),content_signature(root))

    def test_ambiguous_caption_reported(self):
        doc=Document()
        doc.add_picture(str(self.image))
        doc.add_paragraph('图1')
        doc.add_picture(str(self.image))
        doc.save(self.source)
        result=enforce(self.source,self.output)
        self.assertEqual(result['status'],'needs_review')
        self.assertEqual(result['captions'],[])
        self.assertTrue(result['pending'])

    def test_inherited_break_does_not_reach_subheadings(self):
        doc=Document()
        doc.styles['Heading 2'].base_style=doc.styles['Heading 1']
        doc.styles['Heading 1'].paragraph_format.page_break_before=False
        doc.add_heading('一',1)
        doc.add_heading('二级',2)
        doc.add_paragraph('正文')
        doc.save(self.source)
        enforce(self.source,self.output)
        result=Document(self.output)
        second=result.paragraphs[1]
        self.assertFalse(second.paragraph_format.page_break_before)
        self.assertEqual(second.style.name,'Heading 2')
        again=self.folder/'again.docx'
        enforce(self.output,again)
        self.assertEqual(read_package(self.output),read_package(again))

    def test_long_table_is_flagged_for_layout_review(self):
        doc=Document()
        doc.add_paragraph('表1',style='Caption')
        doc.add_table(50,2)
        doc.save(self.source)
        report=enforce(self.source,self.output)
        self.assertEqual(report['status'],'needs_review')
        self.assertTrue(any('超过单页' in item['reason'] for item in report['pending']))

    def test_reject_revision_and_preserve_existing_outputs(self):
        doc=Document(); doc.add_heading('一',1); doc.save(self.source)
        parts=read_package(self.source); root=xml(parts[DOC])
        root.find(q('body'))[0].append(element('ins',id=1))
        revised=self.folder/'revised.docx'
        save_package(parts,{DOC:root},revised)
        with self.assertRaises(ValueError): enforce(revised,self.output)
        self.assertFalse(self.output.exists())
        self.output.write_bytes(b'existing')
        with self.assertRaises(ValueError): enforce(self.source,self.output)
        self.assertEqual(self.output.read_bytes(),b'existing')

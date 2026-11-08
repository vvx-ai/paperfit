"""Phase 2 integration fixtures use real Word packages and native editable math."""
import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from lxml import etree as ET

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from phase2 import process, inventory, read_package, xml, DOC, logical_paragraphs, canonical
from rule_sources import extract
from structured import element, run, seq_field, visible
from paperfit import q, M
from word_fields import scan


def math_node(value):
    math = ET.Element('{'+M+'}oMath')
    run = ET.SubElement(math,'{'+M+'}r')
    ET.SubElement(run,'{'+M+'}t').text = value
    return math


def complex_field(p, code, text):
    for marker,content in [('begin',None),(None,code[:3]),(None,code[3:]),('separate',None),('result',text),('end',None)]:
        r = p.add_run()._r
        if marker is None:
            ET.SubElement(r,q('instrText')).text = content
        elif marker == 'result':
            ET.SubElement(r,q('t')).text = content
        else:
            r.append(element('fldChar',fldCharType=marker))


def sample(path):
    doc = Document()
    doc.add_heading('第一章 试验',1)
    doc.add_paragraph('见表7，参见式（8），数据来源[1]。')
    doc.add_table(1,1).cell(0,0).text='原始数据'
    p = doc.add_paragraph()
    p._p.append(element('bookmarkStart',id=12,name='TableSeven'))
    p.add_run('表')
    complex_field(p,' SEQ 表 \\* ARABIC ','7')
    p._p.append(element('bookmarkEnd',id=12))
    p.add_run(' 实验数据')
    p = doc.add_paragraph('见')
    complex_field(p,' REF TableSeven \\h ','表7')
    p = doc.add_paragraph()
    p._p.append(math_node('x=y'))
    p.add_run('(8)')
    p = doc.add_paragraph('正文中的')
    p._p.append(math_node('x'))
    p.add_run('为变量。')
    doc.add_heading('第二章 讨论',1)
    p = doc.add_paragraph()
    display = ET.Element('{'+M+'}oMathPara')
    display.append(math_node('a+b=c'))
    p._p.append(display)
    doc.add_paragraph('参考文献')
    doc.add_paragraph('甲. 数据处理方法[J]. 示例期刊, 2020.')
    doc.save(path)
    return {'structure':{'captions':'chapter','equations':'chapter','chapters':{'0':1,'6':2},
                         'sync_refs':True,'references':{'entries':[9],'number':True,'hanging_chars':2}}}


class Phase2(unittest.TestCase):
    def setUp(self):
        scratch = tempfile.TemporaryDirectory()
        self.addCleanup(scratch.cleanup)
        self.dir = Path(scratch.name)

    def test_native_fields_equations_bibliography_and_repeat(self):
        src, out, again = [self.dir/name for name in ('src.docx','out.docx','repeat.docx')]
        plan = sample(src)
        report = process(src,plan,out)
        self.assertEqual(report['integrity']['math'],3)
        self.assertEqual(report['stages']['fields']['references'][0]['updated'],'表1-1')
        self.assertEqual(report['stages']['references']['cited'],[1])
        self.assertEqual(report['stages']['equations']['issues'][0]['paragraph'],5)
        data = read_package(out)
        root = xml(data[DOC])
        self.assertIn('见表1-1，参见式（1-1）',visible(root))
        self.assertIn(' 甲. 数据处理方法',visible(root))
        self.assertEqual(report['stages']['native_references']['native_bibliography_entries'],1)
        self.assertIsNotNone(logical_paragraphs(root)[9].find('./'+q('pPr')+'/'+q('numPr')))
        self.assertIn('(2-1)',visible(root))
        self.assertEqual(len(logical_paragraphs(root)),10)
        self.assertEqual(len(inventory(out)['caption_matches']['matches']),1)
        process(out,plan,again)
        self.assertEqual(read_package(again),data)

    def test_conflict_and_table_source(self):
        p = self.dir/'规范.docx'
        doc = Document()
        doc.add_paragraph('正文：宋体小四')
        table = doc.add_table(1,2)
        table.cell(0,0).text = '正文'
        table.cell(0,1).text = '黑体'
        doc.save(p)
        rule = extract(p)
        self.assertTrue(any('冲突' in issue and '表格' in issue for issue in rule['pending']))
        with self.assertRaisesRegex(ValueError,'待确认'):
            process(p,rule,self.dir/'out.docx')
        self.assertFalse((self.dir/'out.docx.report.json').exists())

    def test_pdf_source(self):
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.pdfgen.canvas import Canvas
        p = self.dir/'规范.pdf'
        c = Canvas(str(p))
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        c.setFont('STSong-Light',12)
        c.drawString(50,700,'正文：宋体小四，首行缩进2字符')
        c.showPage()
        c.setFont('STSong-Light',12)
        c.drawString(50,700,'图表全文连续编号')
        c.save()
        result = extract(p)
        self.assertEqual(result['pending'],[])
        self.assertEqual(result['styles']['body']['size'],12)
        self.assertEqual(result['structure']['captions'],'continuous')
        self.assertEqual([s['location'] for s in result['sources']],['第1页','第2页'])

    def test_orphan_ref_and_atomic_failure(self):
        d = Document()
        p = d.add_paragraph()
        p._p.append(seq_field('REF Missing123 \\h','旧值'))
        d.add_paragraph('参见表9。')
        d.add_table(1,1)
        d.add_paragraph('表7 表名')
        src,out = self.dir/'s.docx',self.dir/'o.docx'
        d.save(src)
        self.assertIn('缺失',inventory(src)['field_audit']['references'][0]['status'])
        with self.assertRaisesRegex(ValueError,'歧义'):
            process(src,{'structure':{'captions':'continuous','sync_refs':True}},out)
        self.assertFalse(out.exists())
        self.assertFalse(Path(str(out)+'.report.json').exists())

    def test_position_and_scope(self):
        d = Document()
        d.add_paragraph('段首')
        d.add_table(1,1).cell(0,0).text = '7'
        d.add_paragraph('表1 数据')
        src,out = self.dir/'s.docx',self.dir/'o.docx'
        d.save(src)
        process(src,{'structure':{'captions':'continuous','position':{'表':'before'}}},out)
        root = xml(read_package(out)[DOC])
        self.assertEqual(visible(root.find(q('body'))[1]),'表1 数据')
        self.assertEqual(root.find(q('body'))[2].tag,q('tbl'))
        with self.assertRaises(ValueError):
            process(src,{},out)

    def test_continuous_formula_and_line_breaks(self):
        d = Document()
        for val in ('a','b'):
            d.add_paragraph()._p.append(math_node(val))
        d.add_paragraph('参考文献')
        d.add_paragraph('乙. 示例条目[J].')
        d.add_paragraph('数据[2]')
        src,out = self.dir/'s.docx',self.dir/'o.docx'
        d.save(src)
        plan = {'structure':{'equations':'continuous','references':{'entries':[3],'number':True,'hanging_chars':2}}}
        report = process(src,plan,out)
        self.assertEqual([x['number'] for x in report['stages']['equations']['changes']],['1','2'])
        self.assertTrue(any(x.get('citation')==2 for x in report['pending']))
        self.assertTrue(any('年份' in x.get('issue','') for x in report['pending']))
        part = xml(read_package(out)['word/styles.xml'])
        style = next(s for s in part if s.get(q('styleId'))=='PaperFitBibliography')
        indent = style.find('./'+q('pPr')+'/'+q('ind'))
        self.assertEqual(indent.get(q('hangingChars')),'200')
        self.assertEqual(indent.get(q('hanging')),indent.get(q('left')))
        self.assertTrue(indent.get(q('hanging')))

    def test_conflicting_cited_list_and_unknown_sequence(self):
        d = Document()
        d.add_paragraph('资料[2]')
        d.add_paragraph('[2] 甲. 资料[J]. 2020.')
        d.add_paragraph('[1] 乙. 资料[J]. 2021.')
        src,out = self.dir/'s.docx',self.dir/'o.docx'
        d.save(src)
        with self.assertRaisesRegex(ValueError,'引用含义'):
            process(src,{'structure':{'references':{'entries':[1,2],'number':True}}},out)
        self.assertFalse(out.exists())
        d.add_paragraph('正文含域')._p.append(seq_field('SEQ 表','1'))
        d.add_table(1,1)
        d.add_paragraph('表1 数据')
        d.save(src)
        with self.assertRaisesRegex(ValueError,'同名SEQ'):
            process(src,{'structure':{'captions':'continuous'}},out)
        self.assertFalse(out.exists())

    def test_reviewed_styles_and_readonly_audit(self):
        d = Document()
        d.add_paragraph('论文标题')
        d.add_paragraph('正文abc')
        src,out = self.dir/'s.docx',self.dir/'o.docx'
        d.save(src)
        plan = {'styles':{'body':{'size':12,'font':'宋体','line':1,'indent_chars':2}},'paragraph_roles':{'1':'body'}}
        original = src.read_bytes()
        before = inventory(src,plan)['style_audit']
        self.assertTrue(before['issues'])
        report = process(src,plan,out)
        self.assertFalse(report['stages']['style_audit']['issues'])
        self.assertEqual(src.read_bytes(),original)
        self.assertEqual(visible(logical_paragraphs(xml(read_package(out)[DOC]))[0]),'论文标题')

    def test_scan_only_pdf_and_page_conflicts(self):
        from reportlab.pdfgen.canvas import Canvas
        p = self.dir/'扫描页.pdf'
        c = Canvas(str(p))
        c.rect(20,20,100,100)
        c.showPage()
        c.save()
        self.assertTrue(extract(p)['pending'])
        t = self.dir/'页边距.txt'
        t.write_text('上页边距2厘米；上页边距3厘米',encoding='utf-8')
        self.assertIn('冲突',extract(t)['pending'][0])

    def test_leading_native_ref_is_not_another_caption(self):
        d = Document()
        d.add_paragraph()._p.append(seq_field('REF Missing','表2'))
        d.add_table(1,1).cell(0,0).text='数据'
        d.add_paragraph('表2 真正的题注')
        src,out = self.dir/'s.docx',self.dir/'o.docx'
        d.save(src)
        info = inventory(src)
        self.assertEqual(len(info['caption_matches']['matches']),1)
        report = process(src,{'structure':{'captions':'continuous','sync_refs':True}},out)
        self.assertEqual(report['status'],'needs_review')
        self.assertIn('缺失',report['pending'][0]['issue'])

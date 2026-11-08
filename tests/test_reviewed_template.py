"""Regressions from the school-template acceptance exercise."""
import sys
from pathlib import Path
import tempfile
import unittest
from copy import deepcopy
from docx import Document
from docx.shared import Inches
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from reviewed_template import style_definition, apply_style, remove_suffix, three_line_table, format_template
from phase2 import xml, read_package, save_package, DOC, STYLES, canonical
from paperfit import q
from structured import element, run, visible
from page_flow import enforce, WP, PIC, A
from toc import val
from hashlib import sha256

class ReviewedTemplate(unittest.TestCase):
    def test_mixed_fonts_live_in_style_not_runs(self):
        styles=element('styles')
        s=style_definition(styles,'PaperFitBody','论文正文',dict(font='宋体',latin_font='Times New Roman',size=12,bold=False,line=1,indent_chars=2))
        p=element('p');r=run('中文 English 123');rp=element('rPr');rp.append(element('rFonts',eastAsia='黑体',ascii='Arial'));r.insert(0,rp);p.append(r)
        apply_style(p,'PaperFitBody')
        self.assertEqual(visible(p),'中文 English 123')
        self.assertEqual(val(p,'pPr/pStyle'),'PaperFitBody')
        self.assertIsNone(rp.find(q('rFonts')))
        rf=s.find('./'+q('rPr')+'/'+q('rFonts'))
        self.assertEqual(rf.get(q('eastAsia')),'宋体');self.assertEqual(rf.get(q('ascii')),'Times New Roman')
        self.assertEqual(s.find('./'+q('pPr')+'/'+q('ind')).get(q('firstLineChars')),'200')

    def test_exact_suffix_across_runs_preserves_other_parentheses(self):
        p=element('p');p.extend([run('正文（学术内容）'),run('（格式：宋'),run('体）')])
        remove_suffix(p,'（格式：宋体）')
        self.assertEqual(visible(p),'正文（学术内容）')
        with self.assertRaises(ValueError):remove_suffix(p,'没有这个后缀')

    def test_three_lines_override_imported_cell_borders_without_text_loss(self):
        doc=Document();table=doc.add_table(rows=3,cols=2)
        table.cell(0,0).text='表头';table.cell(1,1).text='42'
        cells=[visible(c) for c in table._tbl.iter(q('tc'))]
        for cell in table._tbl.iter(q('tc')):
            b=element('tcBorders');b.append(element('left',val='double',sz=16));cell.find(q('tcPr')).append(b)
        three_line_table(table._tbl,doc.styles.element,1)
        self.assertEqual(cells,[visible(c) for c in table._tbl.iter(q('tc'))])
        for i,row in enumerate(table._tbl.findall(q('tr'))):
            for cell in row.findall(q('tc')):
                b=cell.find('./'+q('tcPr')+'/'+q('tcBorders'))
                self.assertEqual(val(b,'left'),'nil');self.assertEqual(val(b,'right'),'nil')
                self.assertEqual(val(b,'top'),'single' if i==0 else 'nil')
                self.assertEqual(val(b,'bottom'),'single' if i in (0,2) else 'nil')
        with self.assertRaises(ValueError):three_line_table(table._tbl,doc.styles.element,3)

    def test_manifest_rejects_changed_source_without_output(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'in.docx';output=Path(td)/'out.docx';Document().save(source)
            with self.assertRaisesRegex(ValueError,'已经变化'):format_template(source,{'source_sha256':'invalid'},output)
            self.assertFalse(output.exists())

    def test_existing_table_caption_moves_without_moving_section_end(self):
        with tempfile.TemporaryDirectory() as td:
            source=Path(td)/'in.docx';output=Path(td)/'out.docx'
            doc=Document();doc.add_paragraph('封面');doc.add_section()
            doc.add_paragraph('摘要');doc.add_section()
            heading=doc.add_paragraph('第一章 测试')._p
            table=doc.add_table(2,2)._tbl;caption=doc.add_paragraph('表1')._p
            doc.add_section()
            boundary=caption.getnext();section=boundary.find('./'+q('pPr')+'/'+q('sectPr'))
            caption.insert(0,element('pPr'));caption[0].append(section);boundary.getparent().remove(boundary)
            refhead=doc.add_paragraph('参考文献')._p;ref=doc.add_paragraph('[1]甲. 原文, 2024.')._p
            doc.save(source);nodes=list(doc._element.body)
            rule=dict(font='宋体',latin_font='Times New Roman',size=12,line=1,indent_chars=0)
            styles={sid:dict(name=name,rule=rule,outline=outline) for sid,name,outline in [
                ('PaperFitBody','论文正文',9),('PaperFitTableText','表格正文',9),
                ('PaperFitFrontTitle','前置页标题',9),('PaperFitH1','标题 1',0),
                ('PaperFitBibliography','参考文献正文',9)]}
            plan=dict(source_sha256=sha256(source.read_bytes()).hexdigest(),cover_end=1,
                      body_start=nodes.index(heading),styles=styles,caption_rule=rule,
                      paragraphs=[dict(index=nodes.index(p),text=visible(p),style=sid) for p,sid in [(heading,'PaperFitH1'),(refhead,'PaperFitH1'),(ref,'PaperFitBibliography')]],
                      captions=[dict(index=nodes.index(table),caption_index=nodes.index(caption),label='表',number=1)],
                      tables=[dict(index=nodes.index(table),header_rows=1)])
            report=format_template(source,plan,output);root=xml(read_package(output)[DOC]);body=root.find(q('body'))
            cap=next(p for p in body.findall(q('p')) if visible(p)=='表1')
            self.assertEqual(cap.getnext().tag,q('tbl'))
            self.assertIsNone(cap.find('./'+q('pPr')+'/'+q('sectPr')))
            self.assertEqual(report['pending'],[])
            self.assertEqual(report['native_references']['native_bibliography_entries'],1)

    def test_right_angle_picture_keeps_caption_and_media(self):
        import base64
        with tempfile.TemporaryDirectory() as td:
            folder=Path(td);pic=folder/'a.png';source=folder/'a.docx';rot=folder/'b.docx';out=folder/'out.docx'
            pic.write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jh1sAAAAASUVORK5CYII='))
            doc=Document();doc.add_heading('章',1);doc.add_picture(str(pic),width=Inches(10),height=Inches(3));doc.add_paragraph('图1',style='Caption');doc.save(source)
            parts=read_package(source);root=xml(parts[DOC]);x=root.find('.//{'+PIC+'}spPr/{'+A+'}xfrm');x.set('rot','5400000');save_package(parts,{DOC:root},rot)
            report=enforce(rot,out);result=read_package(out);body=xml(result[DOC]).find(q('body'))
            self.assertFalse(report['pending']);self.assertEqual(val(body[1],'pPr/keepNext'),'1');self.assertEqual(val(body[2],'pPr/keepNext'),'0')
            self.assertEqual(result['word/media/image1.png'],parts['word/media/image1.png'])
            self.assertEqual(xml(result[DOC]).find('.//{'+PIC+'}spPr/{'+A+'}xfrm').get('rot'),'5400000')

if __name__=='__main__':unittest.main()

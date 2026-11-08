"""Native links must retain target identity, text and existing complex objects."""
from copy import deepcopy
from pathlib import Path
import sys
import tempfile
import unittest
from docx import Document
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from native_references import convert_parts, convert, splice, owned_numbering, CITATION_STYLE
from phase2 import read_package, save_package, xml, DOC, canonical
from structured import element, run, seq_field, visible
from paperfit import q
from toc import val
from word_fields import scan, bookmark_ranges, audit


class NativeReferences(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.folder=Path(self.tmp.name)

    def fixture(self, duplicate=False, labels=(1,2,3)):
        doc=Document();doc.add_paragraph('如图1，见表1，来源[2]及[3]；日期[2024-12-17]。')
        for label in ['图','表']+(['图'] if duplicate else []):
            p=doc.add_paragraph(label,style='Caption')
            p._p.append(seq_field('SEQ '+label+' \\* ARABIC',1));p.add_run(' 原有题目')
        doc.add_heading('参考文献',1)
        for n in labels:
            p=doc.add_paragraph();p.add_run('[')
            p._p.append(element('proofErr',type='spellStart'))
            p.add_run(str(n));p.add_run(']');p.add_run('作者'+str(n)+'. 原文[J], 2024.')
            p._p.append(element('proofErr',type='spellEnd'))
        source=self.folder/'source.docx';doc.save(source)
        return source,read_package(source)

    def test_native_lists_label_only_refs_and_repeat(self):
        _,before=self.fixture();after,report=convert_parts(before)
        self.assertEqual(report['pending'],[])
        self.assertEqual(report['native_bibliography_entries'],3)
        self.assertEqual(report['cross_references_created'],4)
        root=xml(after[DOC]);refs=[f for f in scan(root)[0] if f.code.startswith('REF ')]
        self.assertEqual([f.value for f in refs],['图1','表1','[2]','[3]'])
        for f in refs[:2]:
            nodes=bookmark_ranges(root)[f.code.split()[1]][0]
            self.assertEqual(''.join(n.text or '' for n in nodes if n.tag==q('t')),f.value)
        ids=owned_numbering(after);bibs=[p for p in root.iter(q('p')) if val(p,'pPr/numPr/numId') in ids]
        self.assertEqual(len(bibs),3)
        self.assertTrue(all(not visible(p).startswith('[') for p in bibs))
        self.assertIn('日期[2024-12-17]',visible(root))
        self.assertEqual(len(list(root.iter(q('proofErr')))),6)
        again,r=convert_parts(after);self.assertEqual(again,after);self.assertEqual(r['changes'],[])
        self.assertEqual(audit({DOC:root},sync=True)['issues'],[])
        # Package inputs are not mutated; bibliography data is unchanged.
        self.assertEqual(before,read_package(self.folder/'source.docx'))
        self.assertIn('[2]作者2',visible(xml(before[DOC])))

    def test_ambiguous_caption_and_nonsequential_bibliography_are_preserved(self):
        _,parts=self.fixture(duplicate=True,labels=(1,3))
        result,report=convert_parts(parts)
        root=xml(result[DOC])
        self.assertIn('[3]作者3',visible(root));self.assertEqual(report['native_bibliography_entries'],0)
        self.assertEqual(report['cross_references_created'],1) # only the unique table
        self.assertTrue(report['pending'])
        self.assertFalse(any(f.value=='图1' and f.code.startswith('REF ') for f in scan(root)[0]))

    def test_bookmarks_and_hyperlinks_are_not_cut(self):
        p=element('p');p.append(run('['));p.append(element('bookmarkStart',id=8,name='user'))
        p.append(run('2]'));p.append(element('bookmarkEnd',id=8))
        before=canonical(p)
        with self.assertRaisesRegex(ValueError,'书签'):splice(p,0,3)
        self.assertEqual(canonical(p),before)
        link=element('hyperlink');link.append(run('图1'));p=element('p');p.append(link)
        with self.assertRaises(ValueError):splice(p,0,2)

    def test_word_complex_fields_and_target_reordering_retain_identity(self):
        _,parts=self.fixture();parts,_=convert_parts(parts);root=xml(parts[DOC])
        # Model Word's simple-to-complex field save without replacing caches.
        for f in list(root.iter(q('fldSimple'))):
            parent=f.getparent();index=parent.index(f);sequence=[]
            for kind in ('begin','instruction','separate','result','end'):
                if kind=='result':sequence.extend(deepcopy(list(f)));continue
                r=element('r')
                if kind=='instruction':n=element('instrText');n.text=f.get(q('instr'));r.append(n)
                else:r.append(element('fldChar',fldCharType=kind))
                sequence.append(r)
            parent.remove(f)
            for n in sequence:parent.insert(index,n);index+=1
        # Insert another real list entry ahead of the cited item. Existing REF
        # identity must remain the same even while its cached number is stale.
        bib=next(p for p in root.iter(q('p')) if visible(p).startswith('作者2'))
        p=element('p');p.append(deepcopy(bib.find(q('pPr'))));p.append(run('新增条目'));bib.addprevious(p)
        file=self.folder/'complex.docx';save_package(parts,{DOC:root},file);before=read_package(file)
        result,report=convert_parts(before)
        self.assertEqual(result,before);self.assertEqual(report['native_bibliography_entries'],4)
        self.assertEqual(report['changes'],[]);self.assertEqual(scan(xml(result[DOC]))[1],[])

    def test_output_refuses_overwrite_and_malformed_fields(self):
        source,parts=self.fixture();output=self.folder/'out.docx'
        convert(source,output)
        with self.assertRaises(ValueError):convert(source,output)
        root=xml(parts[DOC]);root.find(q('body'))[0].append(element('r'))
        root.find(q('body'))[0][-1].append(element('fldChar',fldCharType='begin'))
        file=self.folder/'bad.docx';save_package(parts,{DOC:root},file)
        with self.assertRaises(ValueError):convert(file,self.folder/'bad-out.docx')
        self.assertFalse((self.folder/'bad-out.docx').exists())

    def test_orphan_ref_and_restarted_list_are_reported(self):
        _,parts=self.fixture();parts,_=convert_parts(parts);root=xml(parts[DOC])
        root.find(q('body'))[0].append(seq_field('REF Missing \\h','图9'))
        numbering=xml(parts['word/numbering.xml']);nid=next(iter(owned_numbering(parts)))
        num=next(n for n in numbering.findall(q('num')) if n.get(q('numId'))==nid)
        override=element('lvlOverride',ilvl=0);override.append(element('startOverride',val=5));num.append(override)
        file=self.folder/'restart.docx';save_package(parts,{DOC:root,'word/numbering.xml':numbering},file)
        before=read_package(file);result,report=convert_parts(before)
        self.assertEqual(result,before)
        self.assertTrue(any('未知自动列表' in p['reason'] for p in report['pending']))
        self.assertTrue(any('目标缺失' in p['reason'] for p in report['pending']))

    def test_grouped_citations_are_reported_without_changing_dates(self):
        _,parts=self.fixture();root=xml(parts[DOC])
        root.find(q('body'))[0].append(run('另见[1,2]，图1、2。'))
        file=self.folder/'groups.docx';save_package(parts,{DOC:root},file)
        result,report=convert_parts(read_package(file))
        self.assertIn('另见[1,2]，图1、2。',visible(xml(result[DOC])))
        self.assertEqual({x['label'] for x in report['pending']},{'[1,2]','图1'})

    def test_only_inline_bibliography_refs_use_superscript_character_style(self):
        _,parts=self.fixture();result,report=convert_parts(parts);root=xml(result[DOC])
        styles=xml(result['word/styles.xml'])
        style=next(s for s in styles if s.get(q('styleId'))==CITATION_STYLE)
        self.assertEqual(style.get(q('type')),'character')
        self.assertEqual(val(style,'rPr/vertAlign'),'superscript')
        refs=[f for f in scan(root)[0] if f.code.startswith('REF ')]
        for f in refs:
            for t in f.result:
                r=next(t.iterancestors(q('r')))
                self.assertEqual(val(r,'rPr/rStyle')==CITATION_STYLE,f.value.startswith('['))
            if f.value.startswith('['):self.assertIn('\\* MERGEFORMAT',f.code)
        bibs=[p for p in root.iter(q('p')) if visible(p).startswith('作者')]
        self.assertTrue(all(not list(p.iter(q('vertAlign'))) for p in bibs))
        self.assertEqual(report['citations_superscripted'],2)
        self.assertNotIn(b'vertAlign',result['word/numbering.xml'])
        again,r=convert_parts(result);self.assertEqual(again,result);self.assertEqual(r['changes'],[])

    def test_existing_ref_with_character_style_preserves_other_formatting(self):
        _,parts=self.fixture();parts,_=convert_parts(parts);root=xml(parts[DOC])
        f=next(f for f in scan(root)[0] if f.value=='[2]' and f.code.startswith('REF '))
        r=f.result[0].getparent();rp=r.find(q('rPr'));rp.find(q('rStyle')).set(q('val'),'Hyperlink')
        rp.append(element('position',val=4));rp.append(element('vertAlign',val='baseline'))
        # An old REF may have separate runs for the brackets and digit.
        r.find(q('t')).text='[';r2=deepcopy(r);r2.find(q('t')).text='2]';r.addnext(r2)
        file=self.folder/'old-ref.docx';save_package(parts,{DOC:root},file)
        result,report=convert_parts(read_package(file));after=xml(result[DOC])
        f=next(f for f in scan(after)[0] if f.value=='[2]' and f.code.startswith('REF '))
        for t in f.result:
            r=t.getparent()
            self.assertEqual(val(r,'rPr/rStyle'),'Hyperlink')
            self.assertEqual(val(r,'rPr/vertAlign'),'superscript')
            self.assertIsNone(r.find('./'+q('rPr')+'/'+q('position')))
        self.assertEqual(visible(root),visible(after))
        self.assertEqual(report['citations_superscripted'],1)

    def test_word_renamed_character_style_does_not_duplicate_or_reformat(self):
        _,parts=self.fixture();parts,_=convert_parts(parts)
        root=xml(parts[DOC]);styles=xml(parts['word/styles.xml'])
        for s in styles:
            if s.get(q('styleId'))==CITATION_STYLE:s.set(q('styleId'),'af0')
        for s in root.iter(q('rStyle')):
            if s.get(q('val'))==CITATION_STYLE:s.set(q('val'),'af0')
        file=self.folder/'word-save.docx';save_package(parts,{DOC:root,'word/styles.xml':styles},file)
        before=read_package(file);result,report=convert_parts(before)
        self.assertEqual(result,before);self.assertEqual(report['changes'],[])

if __name__=='__main__':unittest.main()

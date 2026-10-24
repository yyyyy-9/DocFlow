"""FN-01: native notes preserve identity/content without guessing Word edits."""
from pathlib import Path
import json, shutil, tempfile, unittest, zipfile
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart
from docx.opc.packuri import PackURI
from lxml import etree as E
from hrs_bridge.docx_export import convert
from hrs_bridge.docx_import import convert as restore
from hrs_bridge.ooxml import N, element
from hrs_bridge.runtime import BridgeError
from test_html import runtime


class NativeFootnotes(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp=tempfile.TemporaryDirectory(prefix='bridge-notes-');cls.root=Path(cls.tmp.name);cls.rt=runtime()
        cls.project=cls.root/'source';(cls.project/'assets').mkdir(parents=True)
        cls.raw='# 原生脚注\n\n## 计算\n\n正文[^note]，重复[^note]。\n\n[^note]: **加粗**、`代码`、[链接](https://example.com/)，公式 $\\boldsymbol{E}=mc^2$。\n    第二行中文保留。\n'
        (cls.project/'report.md').write_bytes(cls.raw.encode('utf8'))
        cls.path=cls.root/'native.docx';convert(cls.rt,cls.project,cls.path)

    @classmethod
    def tearDownClass(cls):cls.tmp.cleanup()

    def test_01_native_formula_links_and_unchanged_exact_return(self):
        with zipfile.ZipFile(self.path) as archive:
            notes=E.fromstring(archive.read('word/footnotes.xml'))
            self.assertEqual(len(notes.xpath('//m:oMath',namespaces=N)),1)
            self.assertTrue(notes.xpath('//w:b',namespaces=N))
            self.assertIn('代码',''.join(notes.xpath('//w:t/text()',namespaces=N)))
            self.assertIn('第二行中文保留',''.join(notes.xpath('//w:t/text()',namespaces=N)))
            self.assertIn(b'https://example.com/',archive.read('word/_rels/footnotes.xml.rels'))
        target=self.root/'returned';restore(self.rt,self.path,target)
        self.assertEqual((target/'report.md').read_bytes(),self.raw.encode('utf8'))

    def test_02_changed_note_is_not_silently_replaced_with_old_markdown(self):
        edited=self.root/'edited.docx'
        with zipfile.ZipFile(self.path) as old,zipfile.ZipFile(edited,'x',zipfile.ZIP_DEFLATED) as new:
            for info in old.infolist():
                data=old.read(info.filename)
                if info.filename=='word/footnotes.xml':data=data.replace('第二行中文保留'.encode(),'已修改的脚注正文'.encode())
                new.writestr(info,data)
        target=self.root/'rejected'
        with self.assertRaisesRegex(BridgeError,'FOOTNOTE_EDIT_UNSUPPORTED'):
            restore(self.rt,edited,target,provenance=Path(str(self.path)+'.hrs-bridge.json'))
        self.assertFalse(target.exists())

    def test_03_cover_existing_note_and_relationship_survive(self):
        template=self.root/'cover.docx';doc=Document(self.rt.wps/'templates/cover_template_v0311.docx')
        # A generic cover need not contain a pre-existing footnotes part.
        try:
            existing=doc.part.part_related_by(RT.FOOTNOTES)
        except KeyError:
            root=element('w:footnotes')
            existing=XmlPart(PackURI('/word/footnotes.xml'),CT.WML_FOOTNOTES,root,doc.part.package)
            doc.part.relate_to(existing,RT.FOOTNOTES)
        root=parse_xml(existing.blob)
        old_note=element('w:footnote',id=7);para=element('w:p');run=element('w:r');text=element('w:t');text.text='封面已有脚注';run.append(text);para.append(run);old_note.append(para);root.append(old_note)
        part=XmlPart(existing.partname,CT.WML_FOOTNOTES,root,doc.part.package)
        part.relate_to('https://example.com/cover',RT.HYPERLINK,is_external=True)
        for rel in list(doc.part.rels.values()):
            if rel.reltype==RT.FOOTNOTES:doc.part.rels.add_relationship(rel.reltype,part,rel.rId)
        reference=element('w:r');reference.append(element('w:footnoteReference',id=7));doc._element.xpath('//w:body//w:p')[0].append(reference)
        doc.save(template)
        before=template.read_bytes();output=self.root/'covered.docx';convert(self.rt,self.project,output,template=template)
        with zipfile.ZipFile(output) as archive:
            notes=E.fromstring(archive.read('word/footnotes.xml'))
            old=notes.xpath('w:footnote[@w:id="7"]',namespaces=N)[0]
            self.assertEqual(E.tostring(old,method='c14n'),E.tostring(old_note,method='c14n'))
            self.assertEqual(len(notes.xpath('w:footnote[not(@w:type)]',namespaces=N)),2)
            self.assertIn(b'https://example.com/cover',archive.read('word/_rels/footnotes.xml.rels'))
            self.assertIn(b'https://example.com/',archive.read('word/_rels/footnotes.xml.rels'))
        self.assertEqual(template.read_bytes(),before)
        target=self.root/'cover-return';restore(self.rt,output,target)
        self.assertEqual((target/'report.md').read_bytes(),self.raw.encode('utf8'))

    def test_04_undefined_duplicate_and_nested_notes_fail_without_output(self):
        cases={'missing':'正文[^missing]。\n',
               'duplicate':'正文[^a]。\n\n[^a]: 第一条\n\n[^a]: 第二条\n',
               'nested':'正文[^a]。\n\n[^a]: 嵌套[^b]\n\n[^b]: 第二条\n'}
        for key,text in cases.items():
            with self.subTest(case=key):
                folder=self.root/key;(folder/'assets').mkdir(parents=True)
                (folder/'report.md').write_bytes(('# 错误脚注\n\n## 说明\n\n'+text).encode('utf8'))
                output=self.root/(key+'.docx')
                with self.assertRaises(BridgeError):convert(self.rt,folder,output)
                self.assertFalse(output.exists())

    def test_05_formula_only_inside_note_still_requires_explicit_converter(self):
        original=self.rt.pandoc;self.rt.pandoc=None
        try:
            output=self.root/'missing-formula.docx'
            with self.assertRaisesRegex(BridgeError,'Pandoc'):convert(self.rt,self.project,output)
            self.assertFalse(output.exists())
        finally:self.rt.pandoc=original


if __name__=='__main__':unittest.main()

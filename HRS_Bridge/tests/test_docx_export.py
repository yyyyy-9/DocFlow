"""B3: native objects use the actual template, without inheriting body indent."""
from pathlib import Path
import json
import tempfile
import unittest
from zipfile import ZipFile
from lxml import etree as E
from docx import Document
from hrs_bridge.browser import recover
from hrs_bridge.runtime import sha
from hrs_bridge.ooxml import N, W
from test_html import ROOT, runtime


class DOCXExport(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from hrs_bridge.docx_export import convert
        cls.temp = tempfile.TemporaryDirectory(prefix='bridge-docx-test-')
        cls.folder = Path(cls.temp.name)
        cls.rt = runtime()
        cls.project = cls.folder / 'source'
        recover(cls.rt, ROOT / 'fixtures/hrs_sample.html', cls.project)
        cls.output = cls.folder / 'report.docx'
        cls.before = {str(p.relative_to(cls.project)): sha(p.read_bytes()) for p in cls.project.rglob('*') if p.is_file()}
        cls.receipt = convert(cls.rt, cls.project, cls.output)
        cls.doc = Document(cls.output)

    @classmethod
    def tearDownClass(cls):
        after = {str(p.relative_to(cls.project)): sha(p.read_bytes()) for p in cls.project.rglob('*') if p.is_file()}
        assert after == cls.before, 'Export modified its source'
        cls.temp.cleanup()

    def test_group_is_one_borderless_native_grid(self):
        """GROUP-02: each of two image rows has its own aligned caption row."""
        self.assertEqual(len(self.doc.tables), 5)
        # Group identity is recorded in bridge provenance, not a new Word table type.
        from hrs_bridge.ooxml import read_provenance, bookmark_regions
        meta = read_provenance(self.output)
        group = next(b for b in meta['blocks'] if b['type'] == 'image-group')
        nodes = bookmark_regions(self.doc)[group['bookmark']]
        table = E.fromstring(E.tostring(next(n for n in nodes if n.tag == '{' + W + '}tbl')))
        rows = table.findall('w:tr', N)
        self.assertEqual(len(rows), 4)
        self.assertEqual(len(table.findall('w:tblGrid/w:gridCol', N)), 3)
        self.assertEqual(len(table.xpath('.//w:drawing', namespaces=N)), 6)
        self.assertFalse(table.xpath('.//w:tblHeader|.//w:trHeight', namespaces=N))
        self.assertTrue(all(b.get('{' + W + '}val') == 'nil' for b in table.findall('w:tblPr/w:tblBorders/*', N)))
        self.assertEqual(sum(int(n.get('{' + W + '}w')) for n in table.findall('w:tblGrid/w:gridCol', N)), 8312)
        for image_row, caption_row in zip(rows[::2], rows[1::2]):
            self.assertEqual(len(image_row.xpath('.//w:drawing', namespaces=N)), 3)
            self.assertFalse(image_row.xpath('.//w:t', namespaces=N))
            self.assertFalse(caption_row.xpath('.//w:drawing', namespaces=N))
            self.assertEqual(len(caption_row.findall('w:tc', N)), 3)
        for cell in table.findall('w:tr/w:tc', N):
            width = int(cell.find('w:tcPr/w:tcW', N).get('{' + W + '}w'))
            for extent in cell.xpath('.//wp:extent', namespaces={**N, 'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}):
                self.assertLessEqual(int(extent.get('cx')), (width - 170) * 635)

    def test_only_object_paragraphs_clear_both_indent_units(self):
        """Zero firstLine alone is insufficient: all eight indent attributes must be explicit."""
        for p in self.doc._element.body.xpath('.//w:p'):
            role = p.find('w:pPr/w:pStyle', N)
            if role is not None and role.get('{' + W + '}val') in self.receipt['object_styles']:
                indent = p.find('w:pPr/w:ind', N)
                self.assertIsNotNone(indent)
                for key in ('left','right','firstLine','firstLineChars','hanging','hangingChars','leftChars','rightChars'):
                    self.assertEqual(indent.get('{' + W + '}' + key), '0')
        config = json.loads((self.rt.wps / 'config.json').read_text('utf-8-sig'))
        with ZipFile(self.rt.wps / config['template']) as z:
            template_styles = z.read('word/styles.xml')
        with ZipFile(self.output) as z:
            # HEAD-01: six heading styles gain native numbering bindings. All
            # other template style properties, including body indent, survive.
            actual = E.fromstring(z.read('word/styles.xml'))
        expected = E.fromstring(template_styles)
        for tree in (actual, expected):
            for style in tree.findall('w:style', N):
                if style.get('{' + W + '}styleId') in {'Heading' + str(i) for i in range(1,7)}:
                    for num in style.findall('w:pPr/w:numPr', N):
                        num.getparent().remove(num)
        self.assertEqual(E.tostring(actual, method='c14n'), E.tostring(expected, method='c14n'))
        normal = self.doc.styles['Normal'].element.find('w:pPr/w:ind', N)
        self.assertEqual(normal.get('{' + W + '}firstLineChars'), '200')

    def test_native_equations_fields_bookmarks_and_originals(self):
        """Math is editable OMML; figure references are native REF/SEQ, EMF originals are embedded."""
        xml = self.doc._element
        self.assertGreaterEqual(len(xml.xpath('.//m:oMath')), 10)  # 5 display + inline/cell math
        fields = xml.xpath('.//w:fldSimple/@w:instr')
        # CAP-01: the visible caption label and sequence name must match so that
        # normal WPS/Word caption insertion continues the existing sequence.
        self.assertEqual(sum(s.strip().startswith('SEQ 图 ') for s in fields), 10)
        self.assertEqual(sum(s.strip().startswith('SEQ 表 ') for s in fields), 4)
        self.assertTrue(any(s.strip().startswith('REF HN_') for s in fields))
        with ZipFile(self.output) as z:
            emfs = [n for n in z.namelist() if n.startswith('word/media/') and n.endswith('.emf')]
            self.assertTrue(emfs)
        self.assertTrue(self.receipt['ok'])


if __name__ == '__main__':
    unittest.main()

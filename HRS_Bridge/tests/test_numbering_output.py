"""NUMBERING-02: independent Bridge output uses HRS2's document policy."""
from pathlib import Path
import json,tempfile,unittest
from zipfile import ZipFile
from lxml import etree as E
from hrs_bridge.docx_export import convert
from test_html import runtime

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M='{http://schemas.openxmlformats.org/officeDocument/2006/math}'

def field_codes(path):
    with ZipFile(path) as archive:root=E.fromstring(archive.read('word/document.xml'))
    return root,[node.get(W+'instr','') for node in root.iter(W+'fldSimple')]+[node.text or '' for node in root.iter(W+'instrText')]


class NumberingOutput(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='bridge-numbering-');self.home=Path(self.tmp.name);self.rt=runtime()
        self.project=self.home/'source';(self.project/'assets').mkdir(parents=True)
    def tearDown(self):self.tmp.cleanup()
    def test_formula_number_switch_keeps_native_math_refs_and_source(self):
        raw='# 编号导出\n\n## 说明\n\n<!-- hrs:eq id=eq-first -->\n$$\n\\frac{x+1}{2}=\\boldsymbol{k}\n$$\n\n<!-- hrs:eq id=eq-last -->\n$$\ny=2\n$$\n\n见 @eq:eq-last。\n'.encode('utf8')
        (self.project/'report.md').write_bytes(raw)
        for show in (False,True):
            (self.project/'meta.json').write_text(json.dumps(dict(render_profile='engineering',render_options=dict(show_equation_numbers=show))),encoding='utf8')
            output=self.home/('on.docx' if show else 'off.docx');convert(self.rt,self.project,output)
            xml,codes=field_codes(output);self.assertEqual(len(list(xml.iter(M+'oMath'))),2)
            self.assertEqual(len([s for s in codes if s.strip().startswith('SEQ 公式')]),2 if show else 0)
            self.assertTrue(any(s.strip().startswith('REF ' if show else 'HYPERLINK ') for s in codes),codes)
            self.assertEqual((self.project/'report.md').read_bytes(),raw)
    def test_native_and_structured_tables_share_one_sequence(self):
        raw=('# 表格\n\n## 内容\n\n| 参数 | 数值 |\n|---|---|\n| x | 2.30 |\n\n'
             '```report-table\n'+json.dumps(dict(schema='report-table-v1',id='tbl-second',caption='第二表',rows=[[dict(text='1.5 mm',rowspan=1,colspan=1)]]),ensure_ascii=False)+'\n```\n').encode('utf8')
        (self.project/'report.md').write_bytes(raw)
        output=self.home/'tables.docx';convert(self.rt,self.project,output);xml,codes=field_codes(output)
        self.assertEqual(len(list(xml.iter(W+'tbl'))),2)
        self.assertEqual(len([s for s in codes if s.strip().startswith('SEQ 表')]),2)
        self.assertEqual((self.project/'report.md').read_bytes(),raw)


if __name__=='__main__':unittest.main()

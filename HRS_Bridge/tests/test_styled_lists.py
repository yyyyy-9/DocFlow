"""INT-01: the existing HRS list extension must become a native Word list."""
from pathlib import Path
import tempfile
import unittest
from docx import Document
from docx.oxml.ns import qn
from hrs_bridge.docx_export import convert
from test_html import runtime


class StyledLists(unittest.TestCase):
    def test_five_existing_styles_keep_native_markers_and_independent_chapter_numbering(self):
        """Use the original Marked extension, preserving all text and five display choices."""
        styles=[('bullet','bullet','•'),('decimal','decimal','%1.'),
                ('lower-alpha','lowerLetter','%1.'),('paren-decimal','decimal','（%1）'),
                ('paren-lower-alpha','lowerLetter','（%1）')]
        with tempfile.TemporaryDirectory(prefix='bridge-styled-list-') as directory:
            project=Path(directory)/'source';project.mkdir();(project/'assets').mkdir()
            source='# 列表样式专项\n\n## 章节\n\n'
            for style,_,_ in styles:
                marker='-' if style=='bullet' else '1.'
                source+=f'<!-- hrs:list style={style} -->\n{marker} {style}甲\n{marker} {style}乙\n\n'
            (project/'report.md').write_bytes(source.encode('utf8'))
            output=Path(directory)/'report.docx';convert(runtime(),project,output)
            doc=Document(output);root=doc.part.numbering_part.element
            chapter=doc.styles['Heading 1'].element.pPr.numPr.numId.val
            for style,fmt,label in styles:
                paragraphs=[p for p in doc.paragraphs if p.text in (style+'甲',style+'乙')]
                self.assertEqual(len(paragraphs),2)
                ids={p._p.pPr.numPr.numId.val for p in paragraphs}
                self.assertEqual(len(ids),1)
                nid=ids.pop();self.assertNotEqual(nid,chapter)
                num=next(n for n in root.findall(qn('w:num')) if int(n.get(qn('w:numId')))==nid)
                aid=num.find(qn('w:abstractNumId')).get(qn('w:val'))
                abstract=next(n for n in root.findall(qn('w:abstractNum')) if n.get(qn('w:abstractNumId'))==aid)
                first=abstract.find(qn('w:lvl'))
                self.assertEqual(first.find(qn('w:numFmt')).get(qn('w:val')),fmt)
                self.assertEqual(first.find(qn('w:lvlText')).get(qn('w:val')),label)
            self.assertEqual((project/'report.md').read_bytes(),source.encode('utf8'))


if __name__=='__main__':unittest.main()

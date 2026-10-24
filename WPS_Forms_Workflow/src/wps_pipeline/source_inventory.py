"""Independent, read-only package oracle. Does not import generator or scanner.

Content is compared as ordered text/control/math/image events and table cells.
Physical formatting is tracked separately. Math signatures retain the XML tree.
"""
import hashlib
import json
import posixpath
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def digest(value):
    if not isinstance(value, bytes):
        value = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode('utf8')
    return hashlib.sha256(value).hexdigest()


def tree_value(e):
    return [e.tag, sorted(e.attrib.items()), e.text if len(e) == 0 else None,
            [tree_value(c) for c in e if isinstance(c.tag, str)]]


class Oracle:
    def __init__(self, path):
        self.path = Path(path)
        with ZipFile(path) as z:
            self.parts = {p: z.read(p) for p in z.namelist()}
        self.root = E.fromstring(self.parts['word/document.xml'])
        self.body = self.root.find('w:body', NS)
        rels = E.fromstring(self.parts['word/_rels/document.xml.rels'])
        self.rels = {x.get('Id'): (x.get('Target'), x.get('TargetMode')) for x in rels}
        self.rel_types={x.get('Id'):x.get('Type') for x in rels}

    def resources(self,node):
        answer=[]
        for n in node.iter():
            for key,rid in n.attrib.items():
                if not key.startswith('{'+NS['r']+'}'):continue
                image=(n.tag=='{'+NS['a']+'}blip' and key in ('{'+NS['r']+'}embed','{'+NS['r']+'}link')) or (n.tag=='{urn:schemas-microsoft-com:vml}imagedata' and key=='{'+NS['r']+'}id')
                row=dict(relationship_id=rid,attribute=key,kind='image' if image else 'related')
                rel=self.rels.get(rid)
                if rel is None:row.update(status='MISSING_RELATIONSHIP',target=None)
                elif rel[1]=='External':row.update(status='EXTERNAL',target=rel[0],relationship_type=self.rel_types[rid])
                else:
                    part=rel[0].lstrip('/') if rel[0].startswith('/') else posixpath.normpath(posixpath.join('word',rel[0]))
                    row.update(status='PRESENT' if part in self.parts else 'MISSING_PART',target=part,relationship_type=self.rel_types[rid])
                    if part in self.parts:row['sha256']=digest(self.parts[part])
                answer.append(row)
        return answer

    def events(self, p):
        events = []

        def push(kind, value):
            if kind == 'text' and events and events[-1][0] == 'text':
                events[-1][1] += value
            else:
                events.append([kind, value])

        def walk(e):
            if e.tag in [f"{{{NS['w']}}}pPr", f"{{{NS['w']}}}rPr"]:
                return
            local = E.QName(e).localname
            space = E.QName(e).namespace
            if space == NS['m'] and local in ('oMathPara', 'oMath'):
                push('math_display' if local == 'oMathPara' else 'math_inline', digest(tree_value(e)))
                return
            if space == NS['wp'] and local in ('inline', 'anchor'):
                if not e.xpath('.//a:blip[@r:embed or @r:link]',namespaces=NS):push('raw_fragment',digest(tree_value(e)));return
                embeds = e.xpath('.//a:blip/@r:embed', namespaces=NS)
                images = []
                for rid in embeds:
                    if rid not in self.rels:continue
                    target, mode = self.rels[rid]
                    if mode == 'External':
                        continue
                    part = posixpath.normpath(posixpath.join('word', target))
                    if target.startswith('/'):
                        part = target.lstrip('/')
                    if part in self.parts:images.append(digest(self.parts[part]))
                extent = e.find('wp:extent', NS)
                push('drawing', {'layout': local, 'media': images, 'extent': dict(extent.attrib) if extent is not None else {}})
                return
            if space == NS['w'] and local == 't':
                push('text', e.text or '')
                return
            if space == NS['w'] and local in ('proofErr','lastRenderedPageBreak'):
                push('layout_marker',digest(tree_value(e)));return
            if space == NS['w'] and local == 'pict' and len(e)==1 and e[0].tag=='{urn:schemas-microsoft-com:vml}rect' and e[0].get('{urn:schemas-microsoft-com:office:office}hr')=='t':
                push('legacy_separator', digest(tree_value(e)))
                return
            if space==NS['w'] and local in ('pict','object'):
                equation=any(x.get('ProgID','').lower().startswith(('equation.','mathtype.')) or x.get('equationxml') is not None for x in e.iter())
                kind='legacy_equation' if equation else 'legacy_picture' if local=='pict' else 'embedded_object'
                push(kind,dict(tree=digest(tree_value(e)),resources=self.resources(e)));return
            if space == NS['w'] and local in ('tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen'):
                push(local, sorted(e.attrib.items()))
                return
            if space == NS['w'] and local in ('bookmarkStart', 'bookmarkEnd'):
                push(local, e.get(f"{{{NS['w']}}}id"))
                return
            if space == NS['w'] and local in ('fldChar', 'instrText', 'fldSimple'):
                push('field_evidence', tree_value(e))
                return
            for c in e:
                walk(c)

        for child in p:
            walk(child)
        return events

    def content(self, e):
        tag = E.QName(e).localname
        if tag == 'p':
            return ['paragraph', self.events(e)]
        if tag == 'tbl':
            rows = []
            for tr in e.findall('w:tr', NS):
                cells = []
                for tc in tr.findall('w:tc', NS):
                    span = tc.xpath('w:tcPr/w:gridSpan/@w:val', namespaces=NS)
                    merge = tc.find('w:tcPr/w:vMerge', NS)
                    cells.append({'span': int(span[0]) if span else 1,
                                  'vmerge': None if merge is None else merge.get(f"{{{NS['w']}}}val", 'continue'),
                                  'children': [self.content(c) for c in tc if E.QName(c).localname in ('p', 'tbl')]})
                rows.append(cells)
            return ['table', {'columns': len(e.findall('w:tblGrid/w:gridCol', NS)), 'rows': rows}]
        return ['boundary', tree_value(e)]

    def formatting(self, e):
        nodes = e.xpath('.//w:pPr | .//w:rPr | .//w:tblPr | .//w:tcPr', namespaces=NS)
        return digest([tree_value(n) for n in nodes])

    def inventories(self):
        return {i: self.content(e) for i, e in enumerate(self.body) if E.QName(e).localname != 'sectPr'}

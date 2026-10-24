"""Resolve visible style names against the current saved DOCX, never cached IDs."""
import hashlib
import json
import zipfile
from pathlib import Path
from lxml import etree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
BODY_BOOKMARK = 'WPS_ReportBody'


def resolve_style_names(docx_path, names):
    with zipfile.ZipFile(docx_path) as archive:
        data = archive.read("word/styles.xml")
    root = ET.fromstring(data)
    available = {}
    for node in root.findall(f"{{{W}}}style"):
        name = node.find(f"{{{W}}}name")
        if name is not None:
            available.setdefault(name.get(f"{{{W}}}val").casefold(), []).append({"style_id": node.get(f"{{{W}}}styleId"), "style_type": node.get(f"{{{W}}}type")})
    resolved = {}
    for name in names:
        matches = available.get(name.casefold(), [])
        if len(matches) != 1:
            raise ValueError(f"Style name {name!r}: expected exactly one match, found {len(matches)}")
        resolved[name] = matches[0]
    return {"styles_xml_sha256": hashlib.sha256(data).hexdigest(), "resolved": resolved,
            "validity": "this saved DOCX only; resolve again after each WPS save"}


def inspect_template(docx_path):
    """Validate the declared empty body slots before replacing any template text."""
    from . import resource_root
    root_dir = resource_root()
    config = json.loads((root_dir / 'config.json').read_text(encoding='utf-8-sig'))
    contract = config['template_contract']
    names = contract['required_styles']
    resolved = resolve_style_names(docx_path, names)
    with zipfile.ZipFile(docx_path) as archive:
        for part in ('[Content_Types].xml', 'word/_rels/document.xml.rels'):
            ET.fromstring(archive.read(part))
        root = ET.fromstring(archive.read('word/document.xml'))
    body = root.find(f'{{{W}}}body')
    if body is None:
        raise ValueError('Template has no document body')
    if not len(body) or body[-1].tag != f'{{{W}}}sectPr':
        raise ValueError('Template must end with its final document section')
    slots = body.xpath('.//w:sdt[w:sdtPr/w:tag[@w:val="wps-report-body-v1"]]', namespaces={'w': W})
    anchors = body.xpath('.//w:bookmarkStart[@w:name="'+BODY_BOOKMARK+'"]', namespaces={'w': W})
    if anchors:
        anchor = anchors[0];paragraph = anchor.getparent();ident = anchor.get(f'{{{W}}}id')
        starts = body.xpath('.//w:bookmarkStart[@w:id=$ident]',ident=ident,namespaces={'w':W})
        ends = body.xpath('.//w:bookmarkEnd[@w:id=$ident]',ident=ident,namespaces={'w':W})
        if (slots or len(anchors)!=1 or len(starts)!=1 or len(ends)!=1 or ident is None
                or paragraph.tag!=f'{{{W}}}p' or paragraph.getparent() is not body
                or body[-2] is not paragraph or ends[0].getparent() is not paragraph
                or list(paragraph).index(anchor)>=list(paragraph).index(ends[0])):
            raise ValueError('Template needs one paired body bookmark in its final ordinary paragraph')
        if ''.join(paragraph.xpath('.//w:t/text()',namespaces={'w':W})).strip():
            raise ValueError('Template body bookmark paragraph must be empty; preserve user content')
        paragraphs=[paragraph];slot={'kind':'bookmark','index':len(body)-2,'name':BODY_BOOKMARK}
    elif slots:
        if len(slots) != 1 or slots[0].getparent() is not body or body[-2] is not slots[0]:
            raise ValueError('Template needs exactly one top-level report body slot immediately before the final section')
        content = slots[0].find(f'{{{W}}}sdtContent')
        if content is None:
            raise ValueError('Template body slot has no content')
        paragraphs = list(content)
        slot = {'kind': 'marked', 'index': len(body)-2, 'tag': 'wps-report-body-v1'}
    else:
        paragraphs = list(body)[:-1]
        slot = {'kind': 'legacy', 'index': 0}
    if any(p.tag != f'{{{W}}}p' for p in paragraphs):
        raise ValueError('Template replacement slots must contain plain paragraphs only')
    texts = [''.join(p.xpath('.//w:t/text()',namespaces={'w':W})) for p in paragraphs]
    expected = contract['body_slot_texts']
    if slot['kind']!='bookmark' and texts != expected:
        raise ValueError('Template body must contain only the declared empty slots and final section; preserve user content and choose a compatible template')
    section = body[-1]
    size = section.find(f'{{{W}}}pgSz')
    margin = section.find(f'{{{W}}}pgMar')
    if size is None or margin is None:
        raise ValueError('Template needs explicit page size and margins')
    usable = int(size.get(f'{{{W}}}w')) - int(margin.get(f'{{{W}}}left')) - int(margin.get(f'{{{W}}}right'))
    if usable <= 0:
        raise ValueError('Template has no usable page width')
    plain_tags={f'{{{W}}}p',f'{{{W}}}r',f'{{{W}}}t',f'{{{W}}}pPr',f'{{{W}}}rPr'}
    for paragraph in paragraphs:
        for node in paragraph.iter():
            if node.tag==f'{{{W}}}sectPr':
                raise ValueError('Template replacement slots must not contain section breaks')
            if node.tag in plain_tags:continue
            if slot['kind']=='bookmark' and node in (anchors[0],ends[0]):continue
            if any(parent.tag in (f'{{{W}}}pPr',f'{{{W}}}rPr') for parent in node.iterancestors()):continue
            raise ValueError('Template replacement slots must contain plain text only; objects, fields, bookmarks and embedded content must be kept outside replacement slots')
    return {'template_version':config['template_version'],
            'styles':{name:item['style_id'] for name,item in resolved['resolved'].items()},
            'styles_xml_sha256':resolved['styles_xml_sha256'], 'usable_width_twips':usable,
            'body_slot':slot}


def open_template_body(body, profile):
    """Remove only the validated empty slot and return its title pattern."""
    from copy import deepcopy
    slot = profile['body_slot']
    if slot['kind'] == 'bookmark':
        body.remove(body[slot['index']])
        title=ET.Element(f'{{{W}}}p')
        pr=ET.SubElement(title,f'{{{W}}}pPr')
        ET.SubElement(pr,f'{{{W}}}pStyle').set(f'{{{W}}}val',profile['styles']['Title'])
        # Keep the old title-pattern API without adding visible placeholder text.
        ET.SubElement(ET.SubElement(title,f'{{{W}}}r'),f'{{{W}}}t').text=''
    elif slot['kind'] == 'marked':
        node = body[slot['index']]
        title = deepcopy(node.find(f'{{{W}}}sdtContent')[0])
        body.remove(node)
    else:
        title = deepcopy(body[0])
        for node in list(body)[:-1]:
            body.remove(node)
    return title

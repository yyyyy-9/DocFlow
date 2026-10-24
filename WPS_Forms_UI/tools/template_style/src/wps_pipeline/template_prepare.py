"""Prepare an existing DOCX shell for the existing WPS report assembler.

The caller selects a tail region against a SHA-256 snapshot. Its front matter,
headers and footers retain their own style inheritance; report roles come from
the configured mother template. This module never interprets report content.
"""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from . import resource_root

from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
N = {'w': W}
Q = lambda name: '{' + W + '}' + name
ROOT = resource_root()
SLOT_TAG = 'wps-report-body-v1'
from .resolve_styles import BODY_BOOKMARK


def xml_bytes(node):
    return E.tostring(node, encoding='UTF-8', xml_declaration=True, standalone=True)


def read_package(path):
    with ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}


def inspect_document(path):
    path = Path(path).resolve()
    raw = path.read_bytes()
    parts = read_package(path)
    root = E.fromstring(parts['word/document.xml'])
    body = root.find(Q('body'))
    styles = E.fromstring(parts['word/styles.xml'])
    return {'source': str(path), 'sha256': hashlib.sha256(raw).hexdigest(),
            'sections': len(body.findall('.//' + Q('sectPr'))),
            'styles': [{'id': n.get(Q('styleId')), 'name': n.find(Q('name')).get(Q('val'))}
                       for n in styles.findall(Q('style')) if n.find(Q('name')) is not None],
            'blocks': [{'index': i, 'type': E.QName(n).localname,
                        'text': ''.join(n.xpath('.//w:t/text()', namespaces=N))[:200],
                        'bookmarks':n.xpath('.//w:bookmarkStart/@w:name',namespaces=N),
                        'fields': n.xpath('.//w:instrText/text()', namespaces=N)}
                       for i, n in enumerate(body)],
            'note': 'Inspect only. Choose the first replaceable body index and bind it to this SHA-256.'}


def inspect_styles(path):
    """Read named style facts and inheritance; never infer a document's roles."""
    path = Path(path).resolve()
    parts = read_package(path)
    root = E.fromstring(parts['word/styles.xml'])
    definitions = {n.get(Q('styleId')): n for n in root.findall(Q('style'))}
    defaults = root.find(Q('docDefaults'))
    required = json.loads((ROOT/'config.json').read_text('utf-8-sig'))['template_contract']['required_styles']
    rows = []
    for sid, node in definitions.items():
        name = node.find(Q('name'))
        if name is None:
            continue
        chain, seen, current = [], set(), sid
        while current and current in definitions and current not in seen:
            seen.add(current)
            item = definitions[current]
            chain.append(item)
            based = item.find(Q('basedOn'))
            current = based.get(Q('val')) if based is not None else None
        def attrs(section, prop):
            result = {}
            if defaults is not None:
                default = defaults.find(Q(section+'Default')+'/'+Q(section)+'/'+'/'.join(Q(x) for x in prop.split('/')))
                if default is not None:
                    result.update({E.QName(k).localname:v for k,v in default.attrib.items()})
            for item in reversed(chain):
                element = item.find(Q(section)+'/'+'/'.join(Q(x) for x in prop.split('/')))
                if element is not None:
                    result.update({E.QName(k).localname:v for k,v in element.attrib.items()})
            return result
        fonts, size, spacing, ind = attrs('rPr','rFonts'), attrs('rPr','sz'), attrs('pPr','spacing'), attrs('pPr','ind')
        size_value = size.get('val')
        rows.append({'id':sid,'name':name.get(Q('val')),'type':node.get(Q('type')),
                     'based_on': [n.get(Q('styleId')) for n in chain[1:]],
                     'font_east_asia':fonts.get('eastAsia') or ('主题:'+fonts['eastAsiaTheme'] if fonts.get('eastAsiaTheme') else None),
                     'font_ascii':fonts.get('ascii') or ('主题:'+fonts['asciiTheme'] if fonts.get('asciiTheme') else None),
                     'size_pt':float(size_value)/2 if size_value is not None else None,
                     'paragraph':{'alignment':attrs('pPr','jc').get('val'),'spacing':spacing,'indent':ind},
                     'numbering':{'num_id':attrs('pPr','numPr/numId').get('val'),'level':attrs('pPr','numPr/ilvl').get('val')},
                     'properties_xml':{'paragraph':E.tostring(node.find(Q('pPr')),encoding='unicode') if node.find(Q('pPr')) is not None else None,
                                       'run':E.tostring(node.find(Q('rPr')),encoding='unicode') if node.find(Q('rPr')) is not None else None}})
    available = {row['name'].casefold() for row in rows}
    return {'source':str(path),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'styles':rows,
            'required_styles':required,'missing_required_styles':[name for name in required if name.casefold() not in available],
            'note':'Style inheritance and docDefaults are read. Theme font references remain explicit; no font substitution is guessed.'}


def _dependencies(definitions, seeds):
    result, pending = set(), list(seeds)
    while pending:
        sid = pending.pop()
        if sid in result or sid not in definitions:
            continue
        result.add(sid)
        node = definitions[sid]
        pending.extend(c.get(Q('val')) for c in node
                       if c.tag in (Q('basedOn'), Q('next'), Q('link')))
    return result


def _references(root, mapping, implicit=None):
    if implicit:
        for p in root.iter(Q('p')):
            pr = p.find(Q('pPr'))
            if pr is None:
                pr = E.Element(Q('pPr'))
                p.insert(0, pr)
            if pr.find(Q('pStyle')) is None:
                E.SubElement(pr, Q('pStyle')).set(Q('val'), implicit)
    for node in root.iter():
        if node.tag in (Q('pStyle'), Q('rStyle'), Q('tblStyle'),
                        Q('basedOn'), Q('next'), Q('link')):
            old = node.get(Q('val'))
            if old in mapping:
                node.set(Q('val'), mapping[old])


def _protect_frontmatter(parts, prefix, styles):
    definitions = {n.get(Q('styleId')): n for n in styles.findall(Q('style'))}
    normal = next((sid for sid, s in definitions.items()
                   if s.get(Q('type')) == 'paragraph' and s.get(Q('default')) == '1'), None)
    if normal is None:
        normal = next(sid for sid, s in definitions.items()
                      if s.find(Q('name')).get(Q('val')).casefold() == 'normal')
    furniture = {name: E.fromstring(raw) for name, raw in parts.items()
                 if name.startswith(('word/header', 'word/footer')) and name.endswith('.xml')}
    roots = prefix + list(furniture.values())
    seeds = {normal}
    for root in roots:
        seeds.update(root.xpath('.//w:pStyle/@w:val | .//w:rStyle/@w:val | .//w:tblStyle/@w:val', namespaces=N))
    closure = _dependencies(definitions, seeds)
    used_ids, used_names = set(definitions), {
        n.find(Q('name')).get(Q('val')).casefold() for n in definitions.values()
        if n.find(Q('name')) is not None}
    mapping = {}
    for sid in sorted(closure):
        ident = 'WPSKeep' + str(len(mapping) + 1)
        while ident in used_ids:
            ident += 'x'
        used_ids.add(ident)
        mapping[sid] = ident
    for sid in sorted(closure):
        clone = deepcopy(definitions[sid])
        clone.set(Q('styleId'), mapping[sid])
        clone.attrib.pop(Q('default'), None)
        name = '模板保留 ' + clone.find(Q('name')).get(Q('val'))
        while name.casefold() in used_names:
            name += ' '
        used_names.add(name.casefold())
        clone.find(Q('name')).set(Q('val'), name)
        _references(clone, mapping)
        styles.append(clone)
    for root in roots:
        _references(root, mapping, normal)
    # A refreshed Word TOC creates new entries using its canonical TOC styles.
    # Freeze those definitions too, so refreshing does not adopt report Normal.
    for sid, node in definitions.items():
        name = node.find(Q('name'))
        if name is not None and name.get(Q('val')).casefold().startswith('toc'):
            _references(node, mapping)
    for name, root in furniture.items():
        parts[name] = xml_bytes(root)
    return mapping


def _install_report_styles(styles, donor, required):
    source = {n.get(Q('styleId')): n for n in donor.findall(Q('style'))}
    source_names = {n.find(Q('name')).get(Q('val')).casefold(): sid
                    for sid, n in source.items() if n.find(Q('name')) is not None}
    target_names = {}
    for n in styles.findall(Q('style')):
        name = n.find(Q('name'))
        if name is not None:
            target_names.setdefault(name.get(Q('val')).casefold(), []).append(n)
    required_ids = {source_names[name.casefold()] for name in required}
    closure = _dependencies(source, required_ids)
    used_ids = {n.get(Q('styleId')) for n in styles.findall(Q('style'))}
    mapping = {}
    for sid in sorted(closure):
        name = source[sid].find(Q('name')).get(Q('val'))
        matches = target_names.get(name.casefold(), []) if sid in required_ids else []
        if len(matches) > 1:
            raise ValueError('Ambiguous target style name: ' + name)
        if matches:
            ident = matches[0].get(Q('styleId'))
        else:
            ident = 'WPSReport' + str(len(mapping) + 1)
            while ident in used_ids:
                ident += 'x'
        used_ids.add(ident)
        mapping[sid] = ident
    for sid in sorted(closure):
        clone = deepcopy(source[sid])
        clone.set(Q('styleId'), mapping[sid])
        if sid not in required_ids:
            clone.attrib.pop(Q('default'), None)
            clone.find(Q('name')).set(Q('val'), '报告依赖 ' + mapping[sid])
        _references(clone, mapping)
        previous = styles.find(Q('style') + '[@' + Q('styleId') + '="' + mapping[sid] + '"]')
        if previous is not None:
            styles.replace(previous, clone)
        else:
            styles.append(clone)
    return mapping


def _install_report_numbering(parts, donor_parts, styles, style_mapping):
    """Import the copied styles' dependencies through the existing numbering importer."""
    from .assemble_reports import Assembler
    state=Assembler.__new__(Assembler)
    state.numbering_root=E.fromstring(parts['word/numbering.xml']) if 'word/numbering.xml' in parts else E.Element(Q('numbering'),nsmap={'w':W})
    original=list(state.numbering_root)
    state.rels=E.fromstring(parts['word/_rels/document.xml.rels'])
    state.types=E.fromstring(parts['[Content_Types].xml'])
    state.ids={r.get('Id') for r in state.rels};state.numbering_changed=False
    numbers={}
    for sid in style_mapping.values():
        style=styles.find(Q('style')+'[@'+Q('styleId')+'="'+sid+'"]')
        for node in style.findall('.//'+Q('numId')):
            nid=node.get(Q('val'))
            if nid and nid!='0':node.set(Q('val'),state.import_numbering(donor_parts,nid,numbers,preserve_definition=True))
    if state.numbering_changed:
        for node in state.numbering_root:
            if node not in original:_references(node,style_mapping)
        parts['word/numbering.xml']=xml_bytes(state.numbering_root)
        parts['word/_rels/document.xml.rels']=xml_bytes(state.rels)
        parts['[Content_Types].xml']=xml_bytes(state.types)
    return numbers


def _reset_toc_cache(prefix):
    count = 0
    for parent in prefix:
        for sdt in parent.iter(Q('sdt')):
            if not sdt.xpath('./w:sdtPr/w:docPartObj/w:docPartGallery[@w:val="Table of Contents"]', namespaces=N):
                continue
            content = sdt.find(Q('sdtContent'))
            instructions = content.xpath('.//w:instrText', namespaces=N)
            toc = [n for n in instructions if (n.text or '').strip().upper().startswith('TOC ')]
            if len(toc) != 1:
                raise ValueError('TOC control must contain one native TOC field')
            paragraph = toc[0]
            while paragraph.getparent() is not content:
                paragraph = paragraph.getparent()
            if paragraph.tag != Q('p'):
                raise ValueError('TOC field paragraph must be directly inside its content control')
            index = list(content).index(paragraph)
            field = E.Element(Q('p'))
            pr = paragraph.find(Q('pPr'))
            if pr is not None:
                field.append(deepcopy(pr))
            for kind in ('begin', 'instruction', 'separate', 'end'):
                run = E.SubElement(field, Q('r'))
                if kind == 'instruction':
                    run.append(deepcopy(toc[0]))
                else:
                    E.SubElement(run, Q('fldChar')).set(Q('fldCharType'), kind)
            for node in list(content)[index:]:
                content.remove(node)
            content.append(field)
            count += 1
    return count


def _trim_cover_gap(prefix):
    removed = []
    for i, node in enumerate(prefix):
        if not node.xpath('.//w:br[@w:type="page"]', namespaces=N):
            continue
        j = i - 1
        while j >= 0:
            p = prefix[j]
            meaningful = p.xpath('.//w:fldChar | .//w:instrText | .//w:drawing | .//w:pict | .//w:br | .//w:tab', namespaces=N)
            if p.tag != Q('p') or meaningful or ''.join(p.xpath('.//w:t/text()', namespaces=N)).strip():
                break
            removed.append(j)
            j -= 1
        break
    for i in sorted(removed, reverse=True):
        del prefix[i]
    return sorted(removed)


def prepare_template(source, output, body_start, expected_sha256, trim_cover_gap=False,
                     *, style_source=None, expected_style_sha256=None):
    """Replace only an explicitly selected tail, in a new file, against a snapshot."""
    source, output = Path(source).resolve(), Path(output).resolve()
    receipt_path = output.with_suffix('.template.json')
    if output.exists() or receipt_path.exists() or source == output:
        raise FileExistsError('Output already exists; choose a new DOCX path')
    raw = source.read_bytes()
    actual = hashlib.sha256(raw).hexdigest()
    if actual != expected_sha256:
        raise ValueError('Source SHA-256 changed; inspect the document again before selecting a body index')
    parts = read_package(source)
    root = E.fromstring(parts['word/document.xml'])
    body = root.find(Q('body'))
    if body is not None and body.xpath('.//w:bookmarkStart[@w:name="'+BODY_BOOKMARK+'"]',namespaces=N):
        raise ValueError('This template already has a body bookmark; use it directly')
    if body is not None and body.xpath('.//w:sdt[w:sdtPr/w:tag[@w:val="wps-report-body-v1"]]', namespaces=N):
        from .resolve_styles import inspect_template
        old_profile=inspect_template(source)
        if old_profile['body_slot']['kind']!='marked' or body_start!=old_profile['body_slot']['index']:
            raise ValueError('Select the existing empty body control to upgrade the template')
    if body is None or not len(body) or body[-1].tag != Q('sectPr'):
        raise ValueError('A final document section is required')
    if type(body_start) is not int or not 0 <= body_start < len(body)-1:
        raise ValueError('Body start must select an existing top-level block before the final section')
    selected = list(body)[body_start:-1]
    if any(n.findall('.//' + Q('sectPr')) for n in selected):
        raise ValueError('The selected replacement tail contains a section break; select a compatible body region')
    prefix = list(body)[:body_start]
    removed = _trim_cover_gap(prefix) if trim_cover_gap else []
    config = json.loads((ROOT / 'config.json').read_text('utf-8-sig'))
    donor_path = Path(style_source) if style_source is not None else Path(config['template'])
    if not donor_path.is_absolute():
        donor_path = ROOT / donor_path
    donor_path = donor_path.resolve()
    donor_sha256 = hashlib.sha256(donor_path.read_bytes()).hexdigest()
    if expected_style_sha256 is not None and donor_sha256 != expected_style_sha256:
        raise ValueError('Style source SHA-256 changed; inspect the style template again')
    facts = inspect_styles(donor_path)
    if facts['missing_required_styles']:
        raise ValueError('Style source lacks required named styles: '+', '.join(facts['missing_required_styles']))
    from .resolve_styles import resolve_style_names
    resolve_style_names(donor_path, config['template_contract']['required_styles'])
    if donor_path == output:
        raise ValueError('Output cannot replace the style source')
    donor_parts = read_package(donor_path)
    styles = E.fromstring(parts['word/styles.xml'])
    protected = _protect_frontmatter(parts, prefix, styles)
    mapping = _install_report_styles(styles, E.fromstring(donor_parts['word/styles.xml']),
                                     config['template_contract']['required_styles'])
    numbering_mapping=_install_report_numbering(parts,donor_parts,styles,mapping)
    toc_reset = _reset_toc_cache(prefix)
    section = body[-1]
    for n in list(body):
        body.remove(n)
    body.extend(prefix)
    # A normal blank paragraph is the only replaceable tail. It has no visible
    # title/control and explicitly cancels inherited numbering and outline.
    donor_styles = {n.get(Q('styleId')):n for n in E.fromstring(donor_parts['word/styles.xml']).findall(Q('style'))}
    sid=next(sid for sid,n in donor_styles.items() if n.find(Q('name')).get(Q('val')).casefold()=='normal')
    node=E.SubElement(body,Q('p'));pr=E.SubElement(node,Q('pPr'))
    E.SubElement(pr,Q('pStyle')).set(Q('val'),mapping[sid])
    E.SubElement(E.SubElement(pr,Q('numPr')),Q('numId')).set(Q('val'),'0')
    E.SubElement(pr,Q('outlineLvl')).set(Q('val'),'9')
    used=[int(n.get(Q('id'))) for p in prefix for n in p.iter() if n.tag in (Q('bookmarkStart'),Q('bookmarkEnd')) and (n.get(Q('id')) or '').isdigit()]
    ident=str(max(used+[0])+1)
    E.SubElement(node,Q('bookmarkStart'),{Q('id'):ident,Q('name'):BODY_BOOKMARK})
    E.SubElement(node,Q('bookmarkEnd'),{Q('id'):ident})
    body.append(section)
    settings = E.fromstring(parts['word/settings.xml'])
    update = settings.find(Q('updateFields'))
    if update is None:
        update = E.SubElement(settings, Q('updateFields'))
    update.set(Q('val'), 'true')
    parts['word/settings.xml'] = xml_bytes(settings)
    parts['word/document.xml'] = xml_bytes(root)
    parts['word/styles.xml'] = xml_bytes(styles)
    output.parent.mkdir(parents=True, exist_ok=True)
    if hashlib.sha256(source.read_bytes()).hexdigest()!=actual or hashlib.sha256(donor_path.read_bytes()).hexdigest()!=donor_sha256:
        raise ValueError('Template input changed during preparation')
    # Validate the complete package before publishing either output file.
    import tempfile
    from .resolve_styles import inspect_template
    with tempfile.TemporaryDirectory(prefix='wps-template-',dir=output.parent) as folder:
        staged = Path(folder)/'prepared.docx'
        with ZipFile(staged, 'x', ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
        profile = inspect_template(staged)
        output_raw = staged.read_bytes()
    receipt = {'status': 'PREPARED', 'source': str(source), 'source_sha256': actual,
               'style_source':str(donor_path),'style_source_sha256':donor_sha256,
               'output': str(output), 'output_sha256': hashlib.sha256(output_raw).hexdigest(),
               'body_start_source_index': body_start, 'body_slot': profile['body_slot'],
               'preserved_style_ids': protected, 'trimmed_cover_gap_indices': removed,
               'toc_cache_reset': toc_reset, 'toc_refresh': 'REQUIRED_IN_WORD_OR_WPS',
               'numbering_policy': 'Existing WPS numbering policy unchanged; copied template style dependencies imported',
               'template_numbering_mapping':numbering_mapping,
               'changed_parts': [name for name, value in read_package(source).items() if parts.get(name) != value]}
    created = []
    try:
        for target, data in ((output,output_raw),(receipt_path,(json.dumps(receipt,ensure_ascii=False,indent=2)+'\n').encode('utf8'))):
            with target.open('xb') as stream:
                created.append(target)
                stream.write(data)
    except Exception:
        for target in created:
            target.unlink()
        raise
    return receipt

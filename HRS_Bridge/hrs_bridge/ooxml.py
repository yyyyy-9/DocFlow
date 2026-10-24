"""Bridge-owned DOCX identity/format adapters; no Markdown or numbering parser."""
import base64
from copy import deepcopy
import hashlib
import json
from lxml import etree as E
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .runtime import BridgeError, sha

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M='http://schemas.openxmlformats.org/officeDocument/2006/math'
N={'w':W,'m':M,'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
   'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
   'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
BRIDGE='https://html-report-studio.local/bridge/v1'


def element(tag,**attrs):
    e=OxmlElement(tag)
    for k,v in attrs.items():e.set(qn('w:'+k),str(v))
    return e


def zero_indent(p):
    pr=p.get_or_add_pPr()
    for old in list(pr):
        if old.tag==qn('w:ind'):pr.remove(old)
    pr.append(element('w:ind',left=0,right=0,firstLine=0,firstLineChars=0,hanging=0,hangingChars=0,leftChars=0,rightChars=0))


def field(p,instruction,value):
    f=element('w:fldSimple',instr=instruction);r=element('w:r');t=element('w:t');t.text=str(value);r.append(t);f.append(r);p._p.append(f)


def native_snapshot(nodes,parts,relationships=None):
    """Content fingerprint tolerates WPS run splitting and simple/complex fields.

    This compares observed OOXML, not source text. Cached field results and
    template layout are derived; current text, math structure, media and merges
    are included. Unsupported edits never get replaced with old source.
    """
    ignored={'proofErr','lastRenderedPageBreak','bookmarkStart','bookmarkEnd'}
    note_part=next((p for p in parts.values() if p.content_type=='application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'),None)
    notes={n.get(qn('w:id')):n for n in E.fromstring(note_part.blob)} if note_part is not None else {}
    # WPS rewrites style IDs on save. Resolve names from this saved package,
    # following WPS's existing rule that cached IDs are never portable.
    styles={};default_style='normal'
    for part in parts.values():
        if part.content_type=='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml':
            for s in E.fromstring(part.blob).findall('w:style',N):
                name=s.find('w:name',N)
                if name is not None:
                    styles[s.get(qn('w:styleId'))]=name.get(qn('w:val'),'').casefold()
                    if s.get(qn('w:default'))=='1' and s.get(qn('w:type'))=='paragraph':default_style=name.get(qn('w:val')).casefold()
    def sequence(children,rels=relationships):
        out=[];field_stack=[]
        def append(value):
            if value is None:return
            if out and value[0]=='text' and out[-1][0]=='text' and value[2:]==out[-1][2:]:out[-1][1]+=value[1]
            else:out.append(value)
        def visit(e, emphasis=()):
            name=E.QName(e).localname
            if name in ignored or name.endswith('Pr'):return
            if name=='fldChar':
                what=e.get(qn('w:fldCharType'))
                if what=='begin':field_stack.append([])
                elif what=='end' and field_stack:
                    instruction=' '.join(''.join(field_stack.pop()).split())
                    if not field_stack:append(['field',instruction])
                return
            if name=='instrText' and field_stack:field_stack[-1].append(e.text or '');return
            if name=='hyperlink' and e.get(qn('r:id')):
                rid=e.get(qn('r:id'))
                if rels is None or rid not in rels:raise BridgeError('Missing native hyperlink relationship context')
                append(['hyperlink',rels[rid].target_ref,sequence(e,rels)])
                return
            if name in ('r','hyperlink','smartTag'):
                if name=='r':
                    emphasis=tuple((key, e.find('w:rPr/w:'+key,N).get(qn('w:val'),'1')) for key in ('b','i','u','strike') if e.find('w:rPr/w:'+key,N) is not None)
                for child in e:visit(child,emphasis)
                return
            if field_stack:return
            if name=='fldSimple':append(['field',' '.join(e.get(qn('w:instr'),'').split())]);return
            if name=='t':append(['text',e.text or '',emphasis]);return
            if name=='footnoteReference':
                ident=e.get(qn('w:id'))
                if ident not in notes:raise BridgeError('Missing native footnote definition: '+str(ident))
                if any(n.tag==qn('w:footnoteReference') for n in notes[ident].iter()):raise BridgeError('Nested native footnotes are unsupported')
                append(['footnote',sequence(notes[ident],note_part.rels)]);return
            if name in ('br','cr','tab','noBreakHyphen','softHyphen'):append([name]);return
            if name=='drawing':
                media=[]
                for blip in e.iter('{'+N['a']+'}blip'):
                    rid=blip.get('{'+N['r']+'}embed')
                    if not rid or rid not in parts:raise BridgeError('Unresolved native drawing relationship')
                    media.append(sha(parts[rid].blob))
                transforms=[]
                for node in e.iter():
                    local=E.QName(node).localname
                    if local=='srcRect':
                        values=sorted((k,v) for k,v in node.attrib.items() if v not in ('0','0.0'))
                        if values:transforms.append(['crop',values])
                    elif local=='xfrm':
                        values=sorted((k,v) for k,v in node.attrib.items() if k in ('rot','flipH','flipV') and v not in ('0','false','off'))
                        if values:transforms.append(['transform',values])
                append(['drawing',media,transforms,drawing_extents([e])]);return
            if E.QName(e).namespace==M and name in ('oMath','oMathPara'):
                def math(n):
                    local=E.QName(n).localname
                    if local in ignored:return None
                    if local.endswith('Pr'):
                        if local in ('rPr','ctrlPr','oMathParaPr'):return None
                        # Delimiters, fraction type, n-ary signs and radical
                        # degree visibility are content, not disposable layout.
                        values=[(E.QName(c).localname,sorted(c.attrib.items())) for c in n
                                if E.QName(c).localname!='ctrlPr' and not (local=='fPr' and E.QName(c).localname=='type' and c.get('{'+M+'}val')=='bar')]
                        return [local,values] if values else None
                    children=[]
                    for c in n:
                        value=math(c)
                        if value is None:continue
                        if children and value[0]=='text' and children[-1][0]=='text':children[-1][1]+=value[1]
                        else:children.append(value)
                    if local in ('oMath','oMathPara'):return ['math',children]
                    if local=='t':return ['text',n.text or '']
                    if local=='r' and len(children)==1 and children[0][0]=='text':return children[0]
                    return [local,children]
                expr=e.find('m:oMath',N) if name=='oMathPara' else e
                append(math(expr));return
            if name=='tbl':append(table(e));return
            if name=='p':
                style=e.find('w:pPr/w:pStyle',N)
                sid=style.get(qn('w:val')) if style is not None else None
                append(['p',styles.get(sid,sid) if sid else default_style,sequence(e,rels)])
                return
            if name=='tblGrid':return
            append([e.tag,sequence(e),e.text or ''])
        for child in children:visit(child)
        if field_stack:raise BridgeError('Unclosed native field')
        return out
    def table(e):
        rows=[]
        for row in e.findall('w:tr',N):
            cells=[]
            for cell in row.findall('w:tc',N):
                span=cell.find('w:tcPr/w:gridSpan',N);merge=cell.find('w:tcPr/w:vMerge',N)
                cells.append([span.get(qn('w:val')) if span is not None else '1',
                              merge.get(qn('w:val'),'continue') if merge is not None else None,sequence(cell)])
            rows.append(cells)
        return ['table',rows]
    return sequence(nodes)


def semantic_hash(nodes,parts,relationships=None):
    return sha(json.dumps(native_snapshot(nodes,parts,relationships),ensure_ascii=False,sort_keys=True).encode('utf8'))


def drawing_extents(nodes):
    # Observed WPS SaveAs truncates positive extents to whole twips. Canonicalize
    # that representation exactly, without a widening tolerance that could hide
    # real native resizing. 635 EMU = one twip.
    return [(int(e.get('cx'))//635,int(e.get('cy'))//635)
            for n in nodes for e in n.iter('{'+N['wp']+'}extent')]


def provenance_path(docx):
    from pathlib import Path
    return Path(str(docx)+'.hrs-bridge.json')


def read_provenance(docx, explicit=None):
    path=explicit or provenance_path(docx)
    from pathlib import Path
    path=Path(path)
    if not path.exists():
        if explicit:raise BridgeError('Provenance file not found: '+str(path))
        return None
    data=json.loads(path.read_text('utf8'))
    if data.get('schema')!='hrs-wps-bridge-v1':raise BridgeError('Invalid bridge provenance schema')
    required={'document_key':str,'source_sha256':str,'source_b64':str,'assets':dict,'blobs':dict,'blocks':list}
    if any(not isinstance(data.get(key),kind) for key,kind in required.items()):raise BridgeError('Incomplete bridge provenance')
    return data


def bookmark_regions(doc,scope_name=None):
    """Slice native ranges after WPS moves body bookmarks into paragraphs/math.

    Preserve ancestors and their properties; never assign content by the nearest
    caption. New content outside all known ranges is an explicit error.
    """
    regions={};active=None;active_id=None
    scope_id=None;scope_seen=False;scope_open=scope_name is None;scope_closed=False
    pending_first=None;first_name='HB_'+scope_name[4:]+'_0' if scope_name else None
    def slice_node(node):
        nonlocal active,active_id,scope_id,scope_seen,scope_open,scope_closed,pending_first
        owner_before=active
        name=E.QName(node).localname
        if name=='bookmarkStart':
            label=node.get(qn('w:name'),'')
            if scope_name and label==scope_name:
                if scope_seen:raise BridgeError('Duplicate body scope bookmark')
                scope_seen=scope_open=True;scope_id=node.get(qn('w:id'))
                if pending_first:
                    active,active_id=pending_first;regions[active]=[];pending_first=None
                return {}
            if not scope_open and not scope_seen:
                # WPS may swap coincident starts (HB_0 then HBD). Accept only
                # the declared first block and no intervening content.
                if label==first_name:
                    if pending_first:raise BridgeError('Duplicate first body bookmark')
                    pending_first=(label,node.get(qn('w:id')))
                return {}
            if label.startswith('HB_'):
                if active is not None or label in regions:raise BridgeError('Overlapping or duplicate bridge bookmark')
                active=label;active_id=node.get(qn('w:id'));regions[label]=[]
            return {}
        if name=='bookmarkEnd':
            if scope_name and node.get(qn('w:id'))==scope_id:
                if not scope_open:raise BridgeError('Invalid body scope end')
                scope_open=False;scope_closed=True;return {}
            if not scope_open and not scope_seen:return {}
            if node.get(qn('w:id'))==active_id:active=active_id=None
            return {}
        if name=='sectPr':return {}
        if not scope_open and not scope_seen and not len(node):
            if pending_first and name in ('t','br','tab','drawing','oMath','fldChar','instrText') and (node.text or name!='t'):
                raise BridgeError('BRIDGE_BODY_SCOPE_ORDER: content precedes body scope')
            return {}
        if scope_closed and active is not None and not len(node) and name in ('t','br','tab','drawing','oMath','fldChar','instrText') and (node.text or name!='t'):
            raise BridgeError('BRIDGE_BODY_SCOPE_ORDER: content follows body scope')
        if name=='blip' and active is None:raise BridgeError('UNTRACKED_CONTENT drawing outside bridge object ranges')
        if not len(node):
            if active is None:
                if name in ('t','br','tab','drawing','oMath','fldChar','instrText') and (node.text or name!='t'):
                    raise BridgeError('UNTRACKED_CONTENT outside bridge object ranges')
                return {}
            return {active:deepcopy(node)}
        pieces={};properties=[]
        for child in node:
            local=E.QName(child).localname
            if local.endswith('Pr') or local=='tblGrid':
                properties.append(deepcopy(child));continue
            for owner,part in slice_node(child).items():
                if owner not in pieces:
                    clone=deepcopy(node)
                    for old in list(clone):clone.remove(old)
                    clone.extend(deepcopy(properties));clone.text=None
                    pieces[owner]=clone
                pieces[owner].append(part)
        if not pieces and owner_before is not None and owner_before==active and name in ('p','r','tc','tr','tbl'):
            pieces[active]=deepcopy(node)
        return pieces
    body=E.fromstring(E.tostring(doc._element.body))
    for node in body:
        for owner,part in slice_node(node).items():regions[owner].append(part)
    if active is not None:raise BridgeError('Unclosed bridge bookmark')
    if scope_name and (not scope_seen or not scope_closed):raise BridgeError('BRIDGE_BODY_SCOPE_LOST')
    return regions

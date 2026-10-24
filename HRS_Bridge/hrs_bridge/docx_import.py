"""Restricted native DOCX recovery, using WPS facts and the existing HRS contract."""
import base64
from collections import Counter
from copy import deepcopy
import json
from pathlib import Path
import re
import tempfile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree as E

from .browser import call
from .math_adapter import MathAdapter
from .ooxml import N, W, bookmark_regions, read_provenance, semantic_hash, drawing_extents
from .runtime import BridgeError, new_path, sha


def escaped(text):
    # Serialization of observed Word text, not an alternative Markdown parser.
    return re.sub(r'([\\`*_\[\]<>|@$=#!~])', r'\\\1', text)


def fenced(kind, data):
    return '```' + kind + '\n' + json.dumps(data, ensure_ascii=False, indent=2) + '\n```\n'


class NativeReader:
    def __init__(self, rt, docx, work, assets, identities=None):
        from wps_pipeline.scan_docx import Scanner
        self.rt, self.path, self.work, self.assets = rt, Path(docx), Path(work), Path(assets)
        self.scan = Scanner(self.path, self.work/'scan').run()
        self.diagnostics = []
        # These empty Word editing/rendering markers have no document content.
        # Keep the original scan on disk; adapt only the reader's event stream.
        ignored = Counter();nonempty_markers=[]
        for block in self.scan['blocks']:
            segments = []
            for segment in block.get('segments', []):
                if segment['kind'] in ('unsupported','layout_marker'):
                    node = self.object_xml(segment['object_id'])
                    if node.tag in (qn('w:proofErr'), qn('w:lastRenderedPageBreak')) and not len(node) and not (node.text or '').strip():
                        local = E.QName(node).localname
                        if segment['kind']=='unsupported':ignored[(block['block_id'], local)] += 1
                        self.diagnostics.append({'code':'WORD_EDITOR_MARKER_IGNORED', 'block':block['block_id'], 'marker':local})
                        continue
                    if segment['kind']=='layout_marker':
                        nonempty_markers.append({'block_id':block['block_id'],'reason':E.QName(node).localname,'object_id':segment['object_id']})
                segments.append(segment)
            if 'segments' in block: block['segments'] = segments
        unsupported = []
        for item in self.scan['unsupported']:
            key = (item['block_id'], item['reason'])
            if ignored[key]: ignored[key] -= 1
            else: unsupported.append(item)
        self.scan['unsupported'] = unsupported+nonempty_markers
        if self.scan['unsupported']:
            raise BridgeError('UNSUPPORTED_NATIVE_CONTENT: ' + json.dumps(self.scan['unsupported'], ensure_ascii=False))
        self.facts = {b['block_id']:b for b in self.scan['blocks']}
        self.top = [self.facts[i] for i in self.scan['top_level_ids'] if self.facts[i]['kind'] != 'boundary']
        self.math = None
        self.identities = identities or {}

    def object_xml(self, ident):
        return E.fromstring((self.work/'scan'/self.scan['objects'][ident]['payload_ref']).read_bytes())

    def tex(self, ident):
        if self.math is None: self.math = MathAdapter(self.rt)
        return self.math.tex(self.object_xml(ident))

    def field(self, instruction, caption=False):
        from wps_pipeline.assemble_reports import REF_FIELD
        from wps_pipeline.native_output import SEQ_FIELD
        # The WPS validator restricts the first identifier character to ASCII.
        # Adapt only our two native Chinese caption labels, then reuse its
        # existing switch grammar; unknown fields remain unsupported.
        seq_instruction = re.sub(r'^(\s*SEQ\s+)(图|表)(?=\s|$)',
                                 lambda m: m[1] + {'图':'HRSFigure','表':'HRSTable'}[m[2]],
                                 instruction, flags=re.I)
        if SEQ_FIELD.fullmatch(seq_instruction):
            if caption: return ''
            raise BridgeError('SEQ field outside a known caption')
        match = REF_FIELD.fullmatch(instruction)
        if match and match[1] in self.identities:
            kind, ident = self.identities[match[1]]
            return '@'+kind+':'+ident
        raise BridgeError('UNSUPPORTED_NATIVE_FIELD: ' + instruction.strip())

    def paragraph(self, b, caption=False, plain=False, text_skip=0, formatting=True):
        text=[];stack=[]
        for s in b.get('segments', []):
            kind=s['kind']
            if kind=='field_evidence':
                node=self.object_xml(s['object_id']);local=E.QName(node).localname
                if local=='fldSimple':
                    text.append(self.field(node.get(qn('w:instr'),''),caption));continue
                if local=='fldChar':
                    mode=node.get(qn('w:fldCharType'))
                    if mode=='begin':stack.append([])
                    elif mode=='end' and stack:
                        instruction=''.join(stack.pop())
                        if not stack:text.append(self.field(instruction,caption))
                    continue
                if local=='instrText' and stack:stack[-1].append(node.text or '')
                continue
            if stack:continue
            if kind=='text':
                value=s['text']
                if text_skip:
                    count=min(text_skip,len(value));value=value[count:];text_skip-=count
                value=value if plain else escaped(value)
                props=s.get('run_properties_raw')
                props={E.QName(p[0]).localname:dict(p[1]).get(qn('w:val'),'1') for p in props[3]} if props else {}
                if not plain and formatting:
                    if value and props.get('b') not in (None,'0','false','off'):value='**'+value+'**'
                    if value and props.get('i') not in (None,'0','false','off'):value='*'+value+'*'
                    if value and props.get('strike') not in (None,'0','false','off'):value='~~'+value+'~~'
                text.append(value)
            elif kind in ('math_inline','math_display'):
                if plain:raise BridgeError('Native equation in a plain HRS caption needs an explicit mapping')
                tex=self.tex(s['object_id'])
                text.append(('$$\n'+tex+'\n$$') if kind=='math_display' else '$'+tex+'$')
            elif kind in ('bookmarkStart','bookmarkEnd'):continue
            elif kind in ('br','cr'):text.append('\n')
            elif kind=='tab':text.append('\t')
            elif kind in ('noBreakHyphen','softHyphen'):text.append('-')
            elif kind=='drawing':raise BridgeError('Mixed text/drawing paragraph requires an explicit structure mapping')
            else:raise BridgeError('Unsupported native paragraph event: '+kind)
        if stack:raise BridgeError('Unclosed native field')
        if b.get('hyperlinks'):
            # Do not quietly discard destinations when converting visible text.
            raise BridgeError('Native hyperlink import needs a destination-preserving mapping')
        return ''.join(text)

    def caption(self, b, marker):
        from wps_pipeline.numbered_candidates import caption_prefix
        # Captions now accept safe inline Markdown. Serialize observed Word text
        # with the same escaping as body text; only native OMML creates math.
        text=self.paragraph(b,caption=True,formatting=False).strip()
        parsed=caption_prefix(b.get('plain_text','').strip())
        prefix=escaped(parsed['prefix']) if parsed and parsed['label']==marker else None
        if prefix and text.startswith(prefix):
            self.diagnostics.append({'code':'STATIC_CAPTION_NUMBER_NORMALIZED','block':b['block_id'],
                                     'message':'Original caption number is display text; HRS derives object numbering.'})
            return text[len(prefix):].strip()
        # Known generated caption has a removed SEQ field and a literal label.
        if any(s['kind']=='field_evidence' for s in b.get('segments',[])):
            return text.removeprefix(marker).strip()
        return text

    def image(self, drawing):
        self.check_drawing(drawing)
        obj=self.scan['objects'][drawing]
        if len(obj['media'])!=1 or obj.get('layout')!='inline':
            raise BridgeError('Only one embedded inline image per drawing is supported')
        media=obj['media'][0]
        data=(self.work/'scan'/media['payload_ref']).read_bytes()
        ext=Path(media['part']).suffix.lower()
        if ext not in ('.png','.jpg','.jpeg','.gif','.webp','.svg'):
            raise BridgeError('Edited/external '+ext+' image requires a verified PNG rendition; original source is not substituted')
        name=sha(data)[:16]+ext
        (self.assets/name).write_bytes(data)
        return {'src':'assets/'+name,'caption':''}

    def check_drawing(self, drawing):
        node=self.object_xml(drawing)
        for child in node.iter():
            local=E.QName(child).localname
            if local=='srcRect' and any(v not in ('0','0.0') for v in child.attrib.values()):
                raise BridgeError('CROPPED_IMAGE_UNSUPPORTED: preserve the edited DOCX; no uncropped source is substituted')
            if local=='xfrm' and any(v not in ('0','false','off') for k,v in child.attrib.items() if k in ('rot','flipH','flipV')):
                raise BridgeError('TRANSFORMED_IMAGE_UNSUPPORTED: native rotation/flip requires an explicit HRS representation')

    def table(self, b):
        grid=[[None for _ in range(b['columns'])] for _ in range(b['rows'])]
        if not grid or not grid[0]:raise BridgeError('Native table has no grid')
        anchors={}
        for ident in b['cells']:
            c=self.facts[ident];ri,ci=c['row'],c['col'];span=c['col_span']
            if ci+span>b['columns']:raise BridgeError('Native table grid is inconsistent')
            if c['v_merge_raw']=='continue':
                anchor=anchors.get(ci)
                if anchor is None:raise BridgeError('Native vertical merge has no restart')
                grid[anchor[0]][anchor[1]]['rowspan']+=1
                continue
            texts=[]
            for child in c['children']:
                p=self.facts[child]
                if p['kind']!='paragraph':raise BridgeError('Nested native data table is not supported by the HRS table schema')
                if any(s['kind']=='drawing' for s in p.get('segments',[])):
                    raise BridgeError('IMAGE_IN_DATA_TABLE_UNSUPPORTED: table-captioned image cells require a verified HRS data-table mapping')
                texts.append(self.paragraph(p))
            grid[ri][ci]={'text':'\n'.join(texts),'rowspan':1,'colspan':span}
            for col in range(ci,ci+span):
                if c['v_merge_raw']=='restart':anchors[col]=(ri,ci)
                else:anchors.pop(col,None)
        return grid

    def group(self, b, original):
        d=deepcopy(original)
        cols=d['columns']
        rows=d.get('rows') or (len(d['images'])+cols-1)//cols
        separate_captions=b['rows']==rows*2
        if b['rows'] not in (rows,rows*2) or b['columns']!=cols or len(b['cells'])!=b['rows']*cols:
            raise BridgeError('Edited group grid geometry is not yet supported; source is not restored')

        def content(ident):
            c=self.facts[ident]
            if c['col_span']!=1 or c['v_merge_raw'] is not None:raise BridgeError('Group grid merges are unsupported')
            paragraphs=[self.facts[i] for i in c['children']]
            if any(p['kind']!='paragraph' for p in paragraphs):raise BridgeError('Nested group cell content is unsupported')
            drawings=[s['object_id'] for p in paragraphs for s in p.get('segments',[]) if s['kind']=='drawing']
            return paragraphs,drawings

        for index in range(rows*cols):
            row,column=divmod(index,cols)
            cell_index=(row*2 if separate_captions else row)*cols+column
            paragraphs,drawings=content(b['cells'][cell_index])
            if separate_captions:
                caption_paragraphs,caption_drawings=content(b['cells'][cell_index+cols])
                if caption_drawings:raise BridgeError('Group caption row must not contain drawings')
                if any(p.get('plain_text','').strip() or any(s['kind'] in ('math_inline','math_display','field_evidence','unsupported') for s in p.get('segments',[])) for p in paragraphs):
                    raise BridgeError('Unexpected text in group image row; move it to the caption row')
            else:
                # Existing exports placed each caption below its picture in
                # the same cell. Preserve that bounded legacy read path.
                caption_paragraphs=[p for p in paragraphs if not any(s['kind']=='drawing' for s in p.get('segments',[]))]
            if index>=len(d['images']):
                if drawings or any(p.get('plain_text','').strip() for p in paragraphs+caption_paragraphs):raise BridgeError('Unexpected group content in empty cell')
                continue
            if len(drawings)!=1:raise BridgeError('Group cell requires exactly one current drawing')
            self.check_drawing(drawings[0])
            item=d['images'][index]
            media=self.scan['objects'][drawings[0]]['media']
            if len(media)!=1:raise BridgeError('Group image relationship is ambiguous')
            observed=media[0]['sha256']
            candidates=[item.get('sourceAsset'),item['src']]
            unchanged=any(name and (self.assets/Path(name).name).is_file() and sha((self.assets/Path(name).name).read_bytes())==observed for name in candidates)
            if not unchanged:
                item=self.image(drawings[0]);d['images'][index]=item
            captions=[self.paragraph(p,formatting=False).strip() for p in caption_paragraphs]
            text='\n'.join(s for s in captions if s)
            if d.get('showSubfigureLabels',True):text=text.removeprefix('('+chr(97+index)+')').strip()
            item['caption']=text
        return d


def scan_range(rt, path, nodes, work, assets, identities):
    doc=Document(path);body=doc._element.body;section=deepcopy(body[-1])
    for child in list(body):body.remove(child)
    for node in nodes:body.append(deepcopy(node))
    body.append(section)
    work.mkdir()
    scoped=work/'range.docx';doc.save(scoped)
    return NativeReader(rt,scoped,work,assets,identities)


def restore_provenance(rt, path, meta, work, staged):
    raw=base64.b64decode(meta['source_b64'],validate=True)
    if sha(raw)!=meta['source_sha256']:raise BridgeError('Provenance source hash mismatch')
    if meta.get('document_key')!=sha(raw)[:16]:raise BridgeError('Provenance document binding mismatch')
    if len(raw)>10*1024*1024:raise BridgeError('Provenance source exceeds HRS limit')
    total=0
    for name,digest in meta['assets'].items():
        target=(staged/name).resolve()
        if target.parent!=(staged/'assets').resolve():raise BridgeError('Unsafe provenance asset path')
        data=base64.b64decode(meta['blobs'][digest],validate=True)
        total+=len(data)
        if total>30*1024*1024 or sha(data)!=digest:raise BridgeError('Provenance asset limit/hash mismatch')
        target.write_bytes(data)
    ir=rt.analyze(raw,staged/'assets')
    if not ir['ok']:raise BridgeError('Provenance source fails the current HRS contract')
    blocks=ir['blocks'];mapping=meta['blocks']
    if len(mapping)!=len(blocks):raise BridgeError('Provenance block map is incomplete')
    scope=meta.get('body_scope')
    if scope is not None and scope!={'bookmark':'HBD_'+meta['document_key'],'kind':'wps-report-body-v1'}:raise BridgeError('Invalid provenance body scope')
    doc=Document(path);regions=bookmark_regions(doc,(scope or {}).get('bookmark'))
    derived=meta.get('derived_sections',[])
    if not isinstance(derived,list) or len(derived)>1:raise BridgeError('Invalid derived section map')
    for section in derived:
        expected='HB_'+meta['document_key']+'_references'
        if not isinstance(section,dict) or section.get('kind')!='references' or section.get('bookmark')!=expected:
            raise BridgeError('Invalid references identity')
        nodes=regions.pop(expected,None)
        if nodes is None or semantic_hash(nodes,doc.part.related_parts,doc.part.rels)!=section.get('native_sha256'):
            raise BridgeError('REFERENCE_EDIT_UNSUPPORTED: 文末参考文献已修改或身份丢失，不能用旧源文覆盖当前内容')
    if set(regions)!={b['bookmark'] for b in mapping}:raise BridgeError('BRIDGE_IDENTITY_LOST or unknown block bookmarks')
    identities={}
    for b in blocks:
        if b.get('id'):
            from .docx_export import Writer
            identities[Writer.number_name(b['id'])]=({'table':'tbl','figure':'fig','image-group':'fig','heading':'sec','equation':'eq'}[b['type']],b['id'])
    result=[];changed=[];serializers=[];diagnostics=[]
    for entry in mapping:
        index=entry['index']
        if not isinstance(index,int) or not 0<=index<len(blocks):raise BridgeError('Invalid provenance block index')
        original=blocks[index]
        if entry['bookmark']!='HB_'+meta['document_key']+'_'+str(index) or original['type']!=entry['type'] or original.get('id')!=entry.get('id'):
            raise BridgeError('Provenance identity does not match original HRS source')
    byname={entry['bookmark']:entry for entry in mapping}
    for name,nodes in regions.items():
        entry=byname[name];index=entry['index'];original=blocks[index]
        if entry.get('native_drawing_extents')!=[list(e) for e in drawing_extents(nodes)]:
            raise BridgeError('IMAGE_LAYOUT_EDIT_UNSUPPORTED: current image dimensions differ from the mapped native layout')
        if semantic_hash(nodes,doc.part.related_parts,doc.part.rels)==entry['native_sha256']:
            result.append(original['source']);continue
        if any(n.tag==qn('w:footnoteReference') or (n.tag==qn('w:fldSimple') and 'NOTEREF' in n.get(qn('w:instr'),'').split()) for node in nodes for n in node.iter()):
            raise BridgeError('FOOTNOTE_EDIT_UNSUPPORTED: Word 中的脚注或含脚注块已经修改，当前不能准确回编；原文件保留。')
        changed.append({'index':index,'type':original['type'],'id':original.get('id')})
        reader=scan_range(rt,path,nodes,work/('block-'+str(index)),staged/'assets',identities)
        kind=original['type']
        paragraphs=[b for b in reader.top if b['kind']=='paragraph']
        tables=[b for b in reader.top if b['kind']=='table']
        if kind=='table':
            if len(tables)!=1:raise BridgeError('Edited data table requires one native table')
            data=deepcopy(original['data']);data['rows']=reader.table(tables[0])
            caption=next((p for p in paragraphs if p.get('style_name')=='表题'),None)
            if caption is None or any(p is not caption and p.get('display_text','').strip() for p in paragraphs):
                raise BridgeError('Edited table caption or extra surrounding content needs an explicit mapping')
            data['caption']=reader.caption(caption,'表')
            if data['schema']=='gfm-pipe-table-v1':
                serializers.append({'type':'gfm','data':{'caption':data['caption'],'rows':data['rows'],'align':data['align']},'id':original.get('id')})
                result.append(('serialize',len(serializers)-1))
            else:result.append(fenced('report-table',data))
        elif kind=='image-group':
            if len(tables)!=1:raise BridgeError('Edited group requires its native layout table')
            data=reader.group(tables[0],original['data'])
            caption=next((p for p in paragraphs if p.get('style_name')=='图题'),None)
            if caption is None or any(p is not caption and p.get('display_text','').strip() for p in paragraphs):
                raise BridgeError('Edited group caption or extra surrounding content needs an explicit mapping')
            data['caption']=reader.caption(caption,'图')
            serializers.append({'type':'image-group','data':data});result.append(('serialize',len(serializers)-1))
        elif kind=='equation':
            math=[s for p in paragraphs for s in p.get('segments',[]) if s['kind'] in ('math_display','math_inline')]
            if len(math)!=1 or tables or any(p.get('plain_text','').strip() for p in paragraphs):
                raise BridgeError('Edited block equation requires one native equation without discarded prose')
            serializers.append({'type':'equation','tex':reader.tex(math[0]['object_id']),'id':original.get('id')})
            result.append(('serialize',len(serializers)-1))
        elif kind=='heading':
            if len(paragraphs)!=1 or tables:raise BridgeError('Edited heading requires one native paragraph')
            p=paragraphs[0]
            if 'word_heading_level' in entry and entry['word_heading_level'] is None:
                if (p.get('style_name') or '').casefold()!='title':
                    raise BridgeError('Changing the document title role requires explicit organization')
                level=original['level']
            else:
                outline=p.get('outline_level_raw')
                if outline is None or not 0<=outline<=5:raise BridgeError('Edited heading has no supported real outline style')
                level=outline+1
                if 'word_heading_level' in entry:
                    target=entry['word_heading_level']
                    if type(target) is not int or not 1<=target<=6:raise BridgeError('Invalid mapped Word heading level')
                    level+=original['level']-target
                if not 1<=level<=6:raise BridgeError('Edited heading exceeds the source heading range')
            result.append('#'*level+' '+reader.paragraph(p)+((' {#'+original['id']+'}') if original.get('id') else '')+'\n')
        elif kind=='markdown':
            if tables:raise BridgeError('A new table in an ordinary block needs an explicit source object mapping')
            if any(p.get('numbering',{}).get('num_id') for p in paragraphs):raise BridgeError('Edited native list import is not yet supported')
            result.append('\n\n'.join(reader.paragraph(p) for p in paragraphs)+'\n\n')
        elif kind=='figure':
            drawings=[s['object_id'] for p in paragraphs for s in p.get('segments',[]) if s['kind']=='drawing']
            if len(drawings)!=1:raise BridgeError('Edited figure requires one native drawing')
            reader.check_drawing(drawings[0])
            d=deepcopy(original['data']);obj=reader.scan['objects'][drawings[0]]
            observed=obj['media'][0]['sha256'] if len(obj['media'])==1 else None
            unchanged=any(name and (staged/name).is_file() and sha((staged/name).read_bytes())==observed for name in (d.get('sourceAsset'),d.get('src')))
            if not unchanged:
                if d['kind']=='mermaid':raise BridgeError('Edited Mermaid rendition cannot be reverse-engineered as source')
                image=reader.image(drawings[0]);d={k:v for k,v in d.items() if k not in ('sourceAsset','rendition','rasterFallback')};d['src']=image['src']
            caption=next((p for p in paragraphs if p.get('style_name')=='图题'),None)
            if caption is None or any(p is not caption and p.get('plain_text','').strip() for p in paragraphs):
                raise BridgeError('Edited figure caption or extra surrounding content needs an explicit mapping')
            d['caption']=reader.caption(caption,'图')
            result.append(fenced('report-figure',d))
        else:raise BridgeError('Edited native block type is unsupported: '+kind)
        diagnostics.extend(reader.diagnostics)
    if serializers:
        serialized=call(rt,'serialize',items=serializers)
        result=[serialized[v[1]] if isinstance(v,tuple) else v for v in result]
    text=''.join(result)
    # Preserve BOM and all source bytes on a genuinely unchanged ordered document.
    if not changed and list(regions)==[entry['bookmark'] for entry in mapping]:candidate=raw
    else:candidate=text.encode('utf8')
    origin=meta.get('word_origin') if not changed and candidate==raw else None
    return candidate,changed,diagnostics,origin


def external_groups(reader, plan, recognition):
    """Serialize confirmed cross-row groups and layout-supported single rows with HRS."""
    groups = plan.get('groups',[])
    positions = {b['block_id']:i for i,b in enumerate(reader.top)}
    consumed = set(); pending = []
    for group in groups:
        if not isinstance(group, dict): raise BridgeError('Invalid Word image-group mapping')
        blocks = group.get('blocks'); columns = group.get('columns'); rows = group.get('rows')
        if not isinstance(blocks, list) or not blocks or not all(isinstance(v,str) and v in positions for v in blocks):
            raise BridgeError('Word group references unknown paragraph blocks')
        indices = [positions[v] for v in blocks]
        if indices != list(range(indices[0], indices[0]+len(indices))) or consumed.intersection(blocks):
            raise BridgeError('Word group paragraphs must be consecutive, ordered and used once')
        if type(columns) is not int or not 1 <= columns <= 4 or type(rows) is not int or not 1 <= rows <= 3:
            raise BridgeError('Word group geometry exceeds the existing HRS group contract')
        caption = group.get('caption',''); labels = group.get('showSubfigureLabels',False)
        if not isinstance(caption,str) or type(labels) is not bool:
            raise BridgeError('Invalid Word group caption or labels')
        images = []
        for ident in blocks:
            block = reader.facts[ident]
            if block['kind'] != 'paragraph': raise BridgeError('Word group must reference picture-only paragraphs')
            for segment in block['segments']:
                kind = segment['kind']
                if kind == 'drawing': images.append(reader.image(segment['object_id']))
                elif kind in ('bookmarkStart','bookmarkEnd'): continue
                elif kind == 'text' and not segment['text'].strip(): continue
                else: raise BridgeError('Word group would consume non-image content; decision rejected')
        if not 2 <= len(images) <= 12 or not (rows-1)*columns < len(images) <= rows*columns:
            raise BridgeError('Word group image count does not match the declared grid')
        data = {'schema':'report-image-group-v1','id':'fig-'+reader.scan['source_sha256'][:16]+'-'+blocks[0],
                'columns':columns,'rows':rows,'showSubfigureLabels':labels,'caption':caption,'images':images}
        pending.append((blocks[0],data))
        consumed.update(blocks)
        reader.diagnostics.append({'code':'EXPLICIT_WORD_GROUP_APPLIED','blocks':blocks,'columns':columns,'rows':rows,
                                   'message':'Source-bound caller decision applied; following paragraphs remain unchanged.'})
    for group in recognition['groups']:
        if consumed.intersection(group['blocks']):continue
        images=[reader.image(v) for v in group['drawings']]
        if group['labels']:
            for image,label in zip(images,group['labels']):image['caption']=escaped(label)
        caption=reader.caption(reader.facts[group['caption_block']],'图') if group['caption_block'] and not group['labels'] else ''
        data={'schema':'report-image-group-v1','id':'fig-'+reader.scan['source_sha256'][:16]+'-'+group['blocks'][0],
              'columns':group['columns'],'rows':1,'showSubfigureLabels':False,'caption':caption,'images':images}
        pending.append((group['blocks'][0],data));consumed.update(group['blocks'])
        reader.diagnostics.append({'code':'WORD_INLINE_ROW_AS_GROUP','blocks':group['blocks'],
                                   'message':'Inline pictures fit one row without forced breaks; serialized with the existing HRS group module.'})
    if not pending:return {},set()
    sources = call(reader.rt,'serialize',items=[{'type':'image-group','data':data} for _,data in pending])
    return {ident:text for (ident,_),text in zip(pending,sources)}, consumed


def external_picture_table(reader, block, role, captions):
    """Project a WPS-classified figure container through the existing HRS schema.

    All image instances and cell text are consumed once. Grids outside that
    schema fail explicitly; they are never silently converted into data tables.
    """
    layout=role['figure_layout'];drawings=layout['drawings'];ident='fig-'+reader.scan['source_sha256'][:16]+'-'+block['block_id']
    caption=reader.caption(reader.facts[captions[block['block_id']]['block']],'图') if block['block_id'] in captions else ''
    def text(cell):
        paragraphs=[]
        for bid in cell['paragraphs']:
            p=deepcopy(reader.facts[bid]);p['segments']=[s for s in p['segments'] if s['kind']!='drawing']
            value=reader.paragraph(p,formatting=False).strip()
            if value:paragraphs.append(value)
        if len(paragraphs)>1:reader.diagnostics.append({'code':'WORD_CELL_CAPTION_PARAGRAPHS_JOINED','block':cell['block_id'],'paragraphs':len(paragraphs)})
        return ' '.join(paragraphs)
    if len(drawings)==1:
        image=reader.image(drawings[0]);notes=[text(c) for c in layout['cells']];notes=[v for v in notes if v]
        image.update(schema='report-figure-v1',id=ident,kind='image',alt='',caption=' '.join(notes+([caption] if caption else [])))
        reader.diagnostics.append({'code':'WORD_IMAGE_TABLE_AS_FIGURE','block':block['block_id'],'images':1})
        return fenced('report-figure',image)
    columns=layout['columns'];images=[];image_rows=0;last_row={}
    for ri in range(layout['rows']):
        cells=[c for c in layout['cells'] if c['row']==ri]
        if any(c['col_span']!=1 or c['v_merge_raw'] is not None for c in cells):
            raise BridgeError('HRS_IMAGE_TABLE_LAYOUT_UNSUPPORTED: merged figure cells require an explicit representable mapping')
        if any(c['drawings'] for c in cells):
            if image_rows and len(images)%columns:
                raise BridgeError('HRS_IMAGE_TABLE_LAYOUT_UNSUPPORTED: an incomplete intermediate image row cannot be shifted into another row')
            image_rows+=1;last_row={};gap=False
            for cell in cells:
                notes=text(cell)
                if not cell['drawings']:
                    gap=True
                    if notes:raise BridgeError('HRS_IMAGE_TABLE_LAYOUT_UNSUPPORTED: text-only cell beside pictures has no corresponding HRS slot')
                    continue
                if gap or len(cell['drawings'])!=1:
                    raise BridgeError('HRS_IMAGE_TABLE_LAYOUT_UNSUPPORTED: noncontiguous or multiple images per cell exceed the current HRS grid contract')
                image=reader.image(cell['drawings'][0]);image['caption']=notes;images.append(image);last_row[cell['col']]=image
        else:
            for cell in cells:
                notes=text(cell)
                if not notes:continue
                if cell['col'] not in last_row:raise BridgeError('HRS_IMAGE_TABLE_LAYOUT_UNSUPPORTED: cell text cannot be assigned without losing its layout')
                image=last_row[cell['col']];image['caption']=' '.join(v for v in (image['caption'],notes) if v)
    if not 1<=columns<=4 or not 1<=image_rows<=3 or not 2<=len(images)<=12 or not (image_rows-1)*columns<len(images)<=image_rows*columns:
        raise BridgeError('HRS_IMAGE_TABLE_LAYOUT_UNSUPPORTED: figure grid exceeds existing HRS 3-row / 4-column / 12-image limits')
    if len(images)!=len(drawings):raise BridgeError('HRS_IMAGE_TABLE_CONTENT_MISMATCH: not every source image was consumed')
    data=dict(schema='report-image-group-v1',id=ident,columns=columns,rows=image_rows,showSubfigureLabels=False,caption=caption,images=images)
    reader.diagnostics.append({'code':'WORD_IMAGE_TABLE_AS_GROUP','block':block['block_id'],'images':len(images),'rows':image_rows,'columns':columns})
    return call(reader.rt,'serialize',items=[{'type':'image-group','data':data}])[0]


def import_external(rt,path,work,staged,structure_map=None):
    from .recognition import analyze,load_decisions,picture_row
    reader=NativeReader(rt,path,work,staged/'assets')
    decisions=load_decisions(reader,structure_map)
    recognition=analyze(reader,decisions)
    roles={r['block_id']:r for r in recognition['roles']}
    (staged/'word-recognition.json').write_text(json.dumps(recognition,ensure_ascii=False,indent=2),encoding='utf8')
    chunks=[];skip=set();style_fragments=[];image_layouts=[]
    def heading(text, style, level, prefix=''):
        projection='heading-list' if level>6 else 'heading'
        source=('- ' if level>6 else '#' * level+' ')+text+'\n'
        style_fragments.append((len(chunks),style,projection,prefix))
        chunks.append(source)
    groups,group_blocks=external_groups(reader,decisions,recognition)
    explicit_blocks={ident for g in decisions.get('groups',[]) for ident in g['blocks']}
    captions=recognition['captions']
    # Only native number bookmarks of identified captions acquire HRS object
    # semantics. Literal @fig / [[...]] text remains ordinary Word text.
    from wps_pipeline.native_output import sequence_field,number_bookmarks
    native_body=E.fromstring(E.tostring(Document(path)._element.body))
    for target,cap in captions.items():
        p=native_body[reader.facts[cap['block']]['source_locator']['body_child_index']]
        field=sequence_field(p)
        if field is not None:
            for name in number_bookmarks(p,field['cache']):
                kind='fig' if cap['kind']=='figure' else 'tbl'
                reader.identities[name]=(kind,kind+'-'+reader.scan['source_sha256'][:16]+'-'+target)
    skip.update(v['block'] for target,v in captions.items() if target not in explicit_blocks)
    top=reader.top
    for i,b in enumerate(top):
        if b['block_id'] in skip:continue
        if b['block_id'] in group_blocks:
            if b['block_id'] in groups: chunks.append(groups[b['block_id']])
            continue
        if b['kind']=='table':
            if roles[b['block_id']].get('figure_layout'):
                chunks.append(external_picture_table(reader,b,roles[b['block_id']],captions));continue
            data={'schema':'report-table-v1','id':'tbl-'+reader.scan['source_sha256'][:16]+'-'+b['block_id'],
                  'caption':'','rows':reader.table(b)}
            if b['block_id'] in captions:data['caption']=reader.caption(reader.facts[captions[b['block_id']]['block']],'表')
            chunks.append(fenced('report-table',data));continue
        if b['kind']!='paragraph':raise BridgeError('Unknown native top-level block')
        kinds=[s['kind'] for s in b['segments'] if s['kind'] not in ('bookmarkStart','bookmarkEnd')]
        role=(b.get('style_name') or '').casefold()
        if role=='表题':
            if i+1>=len(top) or top[i+1]['kind']!='table':raise BridgeError('Unpaired native table caption')
            continue
        if role=='图题':raise BridgeError('Unpaired native figure caption')
        if 'drawing' in kinds:
            if len(picture_row(b))!=1:raise BridgeError('Mixed or multiple native drawing paragraph')
            drawing=next(s['object_id'] for s in b['segments'] if s['kind']=='drawing')
            data=reader.image(drawing)
            data.update(schema='report-figure-v1',id='fig-'+reader.scan['source_sha256'][:16]+'-'+b['block_id'],kind='image',alt='')
            obj=reader.scan['objects'][drawing]
            image_layouts.append({'id':data['id'],'source_sha256':obj['media'][0]['sha256'],
                'width_emu':int(obj['extent']['cx']),'height_emu':int(obj['extent']['cy'])})
            if b['block_id'] in captions:data['caption']=reader.caption(reader.facts[captions[b['block_id']]['block']],'图')
            chunks.append(fenced('report-figure',data));continue
        text=reader.paragraph(b)
        resolved=roles[b['block_id']]
        if resolved['role']=='document_title':
            heading(text,'Title',1);continue
        level=resolved.get('level',1)-1 if resolved['role']=='heading' and resolved['status']!='review' else None
        if level is not None and 0<=level<=5:
            prefix=''
            if level==5:
                from wps_pipeline.numbered_candidates import numeric_expression
                plain=b.get('plain_text','');parts=plain.split(maxsplit=1)
                path=parts[0].split('.') if len(parts)==2 else []
                # WPS's ordinary candidate parser covers up to five components.
                # This adapter handles only the already confirmed sixth-level
                # projection, never general prose, values, or Markdown numbering.
                if len(path)==6 and all(v.isascii() and v.isdecimal() and int(v)>0 for v in path) and numeric_expression(plain) is None:
                    count=len(plain)-len(parts[1]);prefix=plain[:count]
                    text=reader.paragraph(b,text_skip=count)
                    reader.diagnostics.append({'code':'OVERFLOW_CHAPTER_PREFIX_REMOVED','block':b['block_id'],'prefix':prefix,
                        'message':'Confirmed six-level chapter path removed from the list; Word native role is retained.'})
            heading(text,'Heading '+str(level+1),level+2,prefix)
            if level==5:reader.diagnostics.append({'code':'WORD_HEADING_OVERFLOW_AS_LIST','block':b['block_id'],
                'message':'Body Heading 6 exceeds HRS H6 under the report title; emitted as a list with a source-bound role hint.'})
            continue
        if b.get('numbering',{}).get('num_id'):raise BridgeError('External native lists require a verified list style mapping')
        # A literal number in a Normal paragraph is not native list evidence.
        # Escape its punctuation so the existing Markdown parser keeps the text.
        literal=re.sub(r'^([ \t]*\d+)([.)])(?=[ \t])',lambda m:m[1]+'\\'+m[2],text)
        if literal!=text:
            reader.diagnostics.append({'code':'LITERAL_NUMBERED_PARAGRAPH_ESCAPED','block':b['block_id'],
                                       'message':'Typed number retained as body text; no list or heading inferred.'})
            text=literal
        chunks.append(text+'\n')
    candidate='\n'.join(chunks).encode('utf8')
    ir=rt.analyze(candidate,staged/'assets')
    origin={'schema':'hrs-word-origin-v1','markdown_sha256':sha(candidate),'word_styles':[],'image_layouts':image_layouts,'source_docx_sha256':reader.scan['source_sha256']}
    if recognition['review']:reader.diagnostics.append({'code':'WORD_STRUCTURE_REVIEW_RETAINED','review':recognition['review'],
        'message':'Unresolved candidate text retained; inspect word-recognition.json before treating its structure as accepted.'})
    offsets=[];offset=0
    for source in chunks:
        offsets.append(offset);offset+=len(source.encode('utf-16-le'))//2+1
    for index,style,projection,prefix in style_fragments:
        source=chunks[index];start=offsets[index]
        origin['word_styles'].append({'start':start,'end':start+len(source.encode('utf-16-le'))//2,
            'source_sha256':sha(source.encode('utf8')),'style':style,'projection':projection,
            **({'source_number_prefix':prefix} if prefix else {})})
    return candidate,[],reader.diagnostics,origin


def convert(rt,source,output,provenance=None,structure_map=None):
    source=Path(source).resolve();output=new_path(output)
    before=sha(source.read_bytes())
    meta=read_provenance(source,provenance)
    if meta is not None and structure_map is not None:
        raise BridgeError('External structure maps cannot override existing bridge provenance')
    if meta is None:
        doc=Document(source)
        if any(b.get(qn('w:name'),'').startswith('HB_') for b in doc._element.iter(qn('w:bookmarkStart'))):
            raise BridgeError('PROVENANCE_REQUIRED: provide the original .hrs-bridge.json for this bridge DOCX')
    with tempfile.TemporaryDirectory(prefix='bridge-import-',dir=output.parent) as temp:
        work=Path(temp);staged=work/'project';(staged/'assets').mkdir(parents=True)
        if meta:
            candidate,changed,diagnostics,origin=restore_provenance(rt,source,meta,work,staged)
        else:candidate,changed,diagnostics,origin=import_external(rt,source,work,staged,structure_map)
        ir=rt.analyze(candidate,staged/'assets')
        from .runtime import source_contract_valid
        if not ir['ok'] or not source_contract_valid(ir):
            raise BridgeError('Recovered source fails HRS validation: '+json.dumps(ir['diagnostics'],ensure_ascii=False))
        if meta:
            original=rt.analyze(base64.b64decode(meta['source_b64'],validate=True))
            ids=lambda model:{b['id'] for b in model['blocks'] if b.get('id')}
            if ids(original)!=ids(ir):raise BridgeError('BRIDGE_IDENTITY_CHANGED: edited content cannot preserve all source identities')
        (staged/'report.md').write_bytes(candidate)
        from render_profile import normalize as normalize_profile, normalize_options
        profile=normalize_profile(meta.get('source_render_profile') if meta else 'engineering')
        (staged/'meta.json').write_text(json.dumps({'render_profile':profile,'render_options':normalize_options(meta.get('source_render_options') if meta else None)})+'\n',encoding='utf8')
        if origin:(staged/'bridge-origin.json').write_text(json.dumps(origin,ensure_ascii=False,indent=2),encoding='utf8')
        if sha(source.read_bytes())!=before:raise BridgeError('Source DOCX changed during import')
        staged.rename(output)
    result={'ok':True,'output':str(output),'markdown_sha256':sha(candidate),'changed_blocks':changed,
            'mode':'provenance-aware' if meta else 'external-structured','diagnostics':diagnostics+ir['diagnostics']}
    if not meta:
        record=json.loads((output/'word-recognition.json').read_text('utf8'))
        result['recognition']={'path':str(output/'word-recognition.json'),'unresolved':record['unresolved'],'model_calls':0}
        result['status']='NEEDS_REVIEW' if record['unresolved'] else 'COMPLETE'
    return result

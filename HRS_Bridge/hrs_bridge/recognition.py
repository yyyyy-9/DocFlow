"""Bridge relations over WPS facts/candidates, before Markdown serialization."""
from copy import deepcopy
import json
from pathlib import Path
import re
from .runtime import BridgeError


def load_decisions(reader, path):
    if path is None:return {}
    plan=json.loads(Path(path).read_text('utf-8-sig'))
    if not isinstance(plan,dict) or plan.get('schema')!='hrs-word-structure-map-v1':
        raise BridgeError('Unsupported Word structure-map schema')
    if plan.get('source_sha256')!=reader.scan['source_sha256']:
        raise BridgeError('STRUCTURE_MAP_SOURCE_MISMATCH: confirm this decision against the current DOCX')
    if not isinstance(plan.get('roles',{}),dict) or not isinstance(plan.get('groups',[]),list):
        raise BridgeError('Word decisions require roles object / groups array')
    return plan


def picture_row(block):
    from wps_pipeline.figure_layout import picture_row as shared_picture_row
    return shared_picture_row(block)


def analyze(reader, decisions=None):
    from wps_pipeline.recognize_reports import predict
    from wps_pipeline.numbered_candidates import manual_prefix,numeric_expression,list_introduction
    decisions=decisions or {};overrides=decisions.get('roles',{})
    initial=predict(reader.scan,overrides=overrides)
    rows={r['block_id']:deepcopy(r) for r in initial['blocks']}
    top=reader.top
    captions={};groups=[];consumed=set();relations=[]
    for i,b in enumerate(top):
        ident=b['block_id'];r=rows[ident]
        if ident in overrides:continue
        if (b.get('style_name') or '').casefold()=='title':
            r.update(role='document_title',level=None,status='auto_rule');continue
        outline=b.get('outline_level_raw')
        if outline is not None and 0<=outline<=5:
            if r.get('status')!='review' or outline==5:r.update(role='heading',level=outline+1,status='auto_rule')
            continue
    # WPS owns picture-table classification, adjacent caption evidence and
    # same-row grouping. This layer only adapts its confirmed result to HRS.
    for ident,r in rows.items():
        if r['role'] not in ('figure_caption','table_caption') or not r.get('relates_to') or r['status'] in ('review','retained_unpaired'):continue
        target=rows[r['relates_to']];kind='figure' if r['role']=='figure_caption' else 'table'
        descriptor=target.get('figure_layout',{})
        if r['relates_to'] in captions:
            raise BridgeError('MULTIPLE_CAPTIONS_FOR_OBJECT: current HRS object mapping requires one caption; source remains unchanged')
        captions[r['relates_to']]={'block':ident,'kind':kind,'labels':descriptor.get('labels')}
        consumed.add(ident)
        relations.append({'object':r['relates_to'],'caption':ident,'subcaptions':bool(r.get('subcaptions')),'evidence':r['evidence']})
    for b in top:
        layout=rows[b['block_id']].get('figure_layout',{})
        if layout.get('kind')!='inline-row':continue
        cap=captions.get(b['block_id'],{})
        groups.append({'blocks':[b['block_id']],'drawings':layout['drawings'],'columns':layout['columns'],'rows':1,
                       'caption_block':cap.get('block'),'labels':cap.get('labels')})
    # Reuse WPS numeric parsing/guards; require an established peer and a real
    # intervening section, rather than promoting every Normal numeric line.
    for i,b in enumerate(top):
        ident=b['block_id'];r=rows[ident]
        if ident in overrides or ident in consumed or b['kind']!='paragraph' or r['role']=='heading':continue
        text=b.get('plain_text','').strip();prefix=manual_prefix(text)
        if not prefix or numeric_expression(text) or len(text)>80 or b.get('numbering',{}).get('num_id'):continue
        if re.search(r'[。；;！？!?：:]$',text):continue
        recent=[v for v in top[max(0,i-3):i] if v.get('plain_text','').strip()]
        if any(list_introduction(v.get('plain_text','')) for v in recent):continue
        numbers=prefix['numbers'];peer=None
        for j in range(i-1,-1,-1):
            prev=top[j];pr=rows[prev['block_id']];pn=manual_prefix(prev.get('plain_text',''))
            if pr['role']=='heading' and pr.get('status')!='review' and pn and pn['numbers']==numbers[:-1]+[numbers[-1]-1]:
                peer=(j,pr);break
        if peer is None:continue
        between=top[peer[0]+1:i]
        content=any(v['kind']=='table' or picture_row(v) or len(v.get('plain_text',''))>20
                    or re.search('[。；;！？!?]$',v.get('plain_text','')) for v in between)
        if not content:continue
        r.update(role='heading',level=peer[1]['level'],status='auto_rule',title=prefix['content'])
        r['evidence'].append('next_numbered_peer_after_section_content')
    # Unresolved candidates are explicitly retained in the receipt. Their text
    # remains present; it is never treated as a confirmed structural decision.
    review=[{'block':b['block_id'],'text':b.get('plain_text',''),'role':rows[b['block_id']]['role'],
             'evidence':rows[b['block_id']]['evidence']} for b in top if rows[b['block_id']].get('status')=='review'
            and b['block_id'] not in consumed]
    return {'schema':'hrs-word-recognition-v1','source_sha256':reader.scan['source_sha256'],
            'recognizer_version':initial['recognizer_version'],'model_calls':0,'initial_status':initial['status'],
            'roles':list(rows.values()),'relations':relations,'groups':groups,'captions':captions,
            'consumed':sorted(consumed),'review':review,'unresolved':len(review)}

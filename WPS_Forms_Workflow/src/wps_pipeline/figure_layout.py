"""Shared figure/layout relations over scan facts; no Markdown, model or source edits."""
from collections import Counter
import re

from .numbered_candidates import manual_prefix, numeric_expression, list_introduction, caption_prefix


def paragraph_shape(block):
    """Effective source indentation; zero spellings/units compare equally."""
    f=block.get('paragraph_format_effective',{})
    def value(key):return float(f.get(key) or 0)
    ind=[]
    for key in ('firstLine','hanging','left','right'):
        chars=value(key+'Chars');length=value(key)
        ind.append(('chars',chars) if chars else ('twips',length) if length else ('zero',0))
    return (f.get('alignment') or 'left',*ind)


def regular_two_character_indent(block):
    f=block.get('paragraph_format_effective',{})
    if any(float(f.get(k) or 0) for k in ('hanging','hangingChars','left','leftChars','right','rightChars')):return False
    chars=float(f.get('firstLineChars') or 0)
    if chars:return chars==200
    size=(block.get('format_group') or [None,None])[1]
    return bool(size and abs(float(f.get('firstLine') or 0)-2*size*20)<1)


def picture_row(block):
    drawings=[]
    for segment in block.get('segments',[]):
        kind=segment['kind']
        if kind=='drawing':drawings.append(segment['object_id'])
        elif kind in ('bookmarkStart','bookmarkEnd','tab'):continue
        elif kind=='text' and not segment['text'].strip():continue
        else:return []
    return drawings


def descendants(byid, block):
    """Physical order, including nested cells/tables, without duplicating aliases."""
    yield block
    for ident in block.get('cells',block.get('children',[])):
        yield from descendants(byid,byid[ident])


def table_layout(scan, block):
    byid={b['block_id']:b for b in scan['blocks']}
    drawings=[s['object_id'] for b in descendants(byid,block) for s in b.get('segments',[]) if s['kind'] in ('drawing','legacy_picture')]
    if not drawings:return None
    cells=[]
    for ident in block['cells']:
        cell=byid[ident];content=list(descendants(byid,cell))
        cells.append(dict(block_id=ident,row=cell['row'],col=cell['col'],col_span=cell['col_span'],
                          v_merge_raw=cell['v_merge_raw'],
                          paragraphs=[b['block_id'] for b in content if b['kind']=='paragraph'],
                          drawings=[s['object_id'] for b in content for s in b.get('segments',[]) if s['kind'] in ('drawing','legacy_picture')]))
    return dict(kind='table',rows=block['rows'],columns=block['columns'],drawings=drawings,cells=cells)


def apply(scan, candidate, overrides=None):
    """Complete associations before final review/number authorization is generated.

    An image-bearing table without a linked table caption is a figure container.
    Classification does not discard text or require cells to be image-only.
    """
    overrides=overrides or {};byid={b['block_id']:b for b in scan['blocks']}
    rows={r['block_id']:r for r in candidate['blocks']}
    top=[byid[i] for i in scan['top_level_ids'] if rows[i].get('status')!='excluded_source_frontmatter']
    protected=set(overrides)|{r['block_id'] for r in rows.values() if r.get('decision')}
    positions={b['block_id']:i for i,b in enumerate(top)}
    def heading(b):
        r=rows[b['block_id']]
        # A caption may have inherited an outline level in the source.
        # Its visible label and object context must still be assessed.
        if r['role'] in ('figure_caption','table_caption'):return False
        return (b.get('outline_level_raw') in range(6) or r['role'] in ('document_title','front_heading')
                or (r['role']=='heading' and (r.get('status')!='review' or r.get('number_candidate'))))
    def nearby(index, direction):
        distance=0;j=index+direction
        while 0<=j<len(top) and distance<3:
            b=top[j]
            if heading(b):break
            if rows[b['block_id']]['role']!='spacer':
                distance+=1;yield b
            j+=direction
    def nearest(index,targets):
        unique={b['block_id']:b for b in targets};ranked=[]
        for b in unique.values():
            j=positions[b['block_id']];a,z=sorted((index,j))
            distance=sum(rows[top[k]['block_id']]['role']!='spacer' for k in range(a+1,z+1))
            ranked.append(((distance,abs(index-j)),b))
        ranked.sort(key=lambda v:v[0])
        return [b for score,b in ranked if score==ranked[0][0]] if ranked else []
    def adjacent(index,direction):
        j=index+direction
        while 0<=j<len(top):
            b=top[j]
            if rows[b['block_id']]['role']!='spacer':return None if heading(b) else b
            j+=direction
        return None
    ordinary=[paragraph_shape(b) for b in top if b['kind']=='paragraph'
              and b.get('outline_level_raw') not in range(6) and not picture_row(b)
              and not caption_prefix(b.get('display_text','').strip())
              and len(re.sub(r'\s+','',b.get('plain_text','')))>40]
    normal=scan.get('normal_paragraph_layout',('left',)+('0',)*8)
    normal_block=dict(paragraph_format_effective=dict(zip(('alignment','firstLine','firstLineChars','hanging','hangingChars','left','leftChars','right','rightChars'),normal)))
    baseline=Counter(ordinary).most_common(1)[0][0] if ordinary else paragraph_shape(normal_block)
    fields={f['block_id']:f for f in scan.get('body_layout_facts',[])}
    assessments={}
    for b in top:
        r=rows[b['block_id']];text=b.get('display_text','').strip()
        if b['kind']!='paragraph' or b['block_id'] in protected or heading(b) or not text or picture_row(b):continue
        prefix=caption_prefix(text);style=(b.get('style_name') or '').casefold()
        styled=style in ('图题','表题','caption','题注','图注','表注')
        shape=paragraph_shape(b)
        distinct=shape[1:]!=baseline[1:] and not regular_two_character_indent(b)
        positive=(styled or shape[0]=='center' or distinct)
        # A field may corroborate a compact title, never override running prose.
        seq=bool(prefix and any(re.match(r'\s*SEQ\s+',x,re.I) for x in fields.get(b['block_id'],{}).get('field_instructions',[])))
        positive=positive or seq
        remainder=text[prefix['span'][1]:].lstrip() if prefix else text
        explanatory=bool(prefix and (re.match(r'^(?:[中内上下](?:可见|显示|所示)|所示(?:为|的)|显示(?:了|出|不同|该|这|其|一个|一种)|说明了|表明|给出(?:了|的是)|展示(?:了|出)|反映(?:了|出)|可见|可以|分别为|是|为|对应于|包含了|采用了)',remainder) or
            (re.search(r'[。；;！？!?]$',text) and (re.search(r'[，,；;]',remainder) or re.match(r'^的.*(?:一致|不同|相同|表明|显示|说明|可见|不得)',remainder)))))
        prose=len(text)>100 or (len(text)>30 and bool(re.search(r'[。；;！？!?]',text))) or explanatory
        family='table' if (prefix and prefix['label']=='表') or style in ('表题','表注') else 'figure'
        # Compact labels on their conventional adjacent side are useful
        # evidence even when both body and caption happen to be flush left.
        i=positions[b['block_id']];near=adjacent(i,-1 if family=='figure' else 1)
        conventional=bool(prefix and not prose and near and rows[near['block_id']]['role']==family)
        positive=positive or conventional
        assessments[b['block_id']]=dict(positive=positive,prose=prose,styled=styled,seq=seq,prefix=prefix,family=family)
        if not prefix and not styled:continue
        r['caption_evidence']=dict(contract='source-layout-v1',source_shape=shape,body_shape=baseline,
                                   distinct_indent=distinct,normal_two_character_indent=regular_two_character_indent(b),
                                   caption_style=styled,sequence_field=seq,prose=prose,
                                   explanatory_sentence=explanatory,compact_label_on_conventional_side=conventional)
        r.update(level=None,title=None,relates_to=None,association_candidates=[])
        r.pop('level_pending',None)
        if prose:
            r.update(role='body',status='review' if positive else 'auto_rule',possible_roles=['body',family+'_caption'])
            r['evidence'].append('caption_layout_conflicts_with_prose' if positive else 'caption_prefix_in_normal_body_prose')
        else:
            r.update(role=family+'_caption',status='auto_rule' if positive else 'review',possible_roles=[family+'_caption','body'])
            r['evidence'].append('source_caption_layout_supported' if positive else 'caption_prefix_without_distinct_source_layout')
    # Explicit table-caption evidence is resolved before the no-table-caption
    # rule. A caption still awaiting a unique table keeps those tables as data.
    for i,b in enumerate(top):
        r=rows[b['block_id']]
        if b['kind']!='paragraph' or b['block_id'] in protected or heading(b):continue
        if r['role']=='table_caption':
            evidence=assessments.get(b['block_id'],{})
            targets=[v for direction in (1,-1) for v in nearby(i,direction) if v['kind']=='table'] if evidence.get('styled') or evidence.get('seq') else [v for v in (adjacent(i,1),adjacent(i,-1)) if v and v['kind']=='table']
            targets=nearest(i,targets)
            if targets:r.update(relates_to=targets[0]['block_id'] if len(targets)==1 and r['status']!='review' else None,
                                status='auto_rule' if len(targets)==1 and r['status']!='review' else 'review',association_candidates=[t['block_id'] for t in targets])
    # Once a caption has a unique table, discarded alternatives must not turn
    # nearby picture grids into data tables. Only a genuinely pending pair
    # keeps its competing physical tables reserved.
    declared_tables={ident for r in rows.values() if r['role']=='table_caption'
                     for ident in (([r.get('relates_to')]+r.get('association_candidates',[]))
                                   if r.get('status')=='review' else [r.get('relates_to')]) if ident}
    for b in top:
        r=rows[b['block_id']]
        if b['kind']=='table' and b['block_id'] in declared_tables:
            r.pop('figure_layout',None);r.update(role='table',level=None,title=None)
        elif b['kind']=='table':
            layout=table_layout(scan,b)
            if layout:
                r.update(role='figure',level=None,title=None,status='auto_rule',figure_layout=layout)
                r['evidence'].append('image_table_without_table_caption')
        elif b['kind']=='paragraph':
            drawings=picture_row(b)
            if not 2<=len(drawings)<=4 or b.get('hyperlinks'):continue
            objects=[scan['objects'][v] for v in drawings]
            if any(o.get('layout')!='inline' for o in objects):continue
            width=sum(int(o.get('extent',{}).get('cx',0)) for o in objects)
            usable=scan.get('body_width_emu')
            if not width or not usable or width>usable:continue
            r.update(role='figure',level=None,title=None,status='auto_rule',
                     figure_layout=dict(kind='inline-row',rows=1,columns=len(drawings),drawings=drawings))
            r['evidence'].append('inline_pictures_fit_single_row')
    for i,b in enumerate(top):
        r=rows[b['block_id']];text=b.get('display_text','').strip()
        if b['kind']!='paragraph' or b['block_id'] in protected or heading(b) or not text or picture_row(b):continue
        a=assessments.get(b['block_id'])
        if not a or a['prose']:continue
        explicit=a['styled'] or a['seq']
        prefixed=a['prefix'] is not None
        if not explicit and not prefixed and (manual_prefix(text) or numeric_expression(text) or list_introduction(text) or re.search(r'[。；;！？!?：:]$',text)):continue
        if not a['positive'] and not prefixed:continue
        family=a['family']
        if prefixed and not a['positive']:
            conventional=adjacent(i,-1 if family=='figure' else 1)
            if conventional and rows[conventional['block_id']]['role']==family:
                a['positive']=True;r.setdefault('caption_evidence',{})['compact_label_on_conventional_side']=True
        matches=[]
        if explicit:
            for direction in ((1,-1) if family=='table' else (-1,1)):
                for target in nearby(i,direction):
                    if rows[target['block_id']]['role']==family:
                        matches.append(target);break
        else:
            before=adjacent(i,-1);after=adjacent(i,1)
            if prefixed:
                matches=[v for v in (before,after) if v and rows[v['block_id']]['role']==family]
            elif before and rows[before['block_id']]['role']=='figure':matches=[before]
            elif after and rows[after['block_id']]['role']=='table':matches=[after];family='table'
        if not matches:
            if prefixed or explicit:r.update(status='retained_unpaired',relates_to=None,association_candidates=[])
            continue
        if explicit or prefixed:
            choices=nearest(i,matches)
            if len(choices)>1 or not a['positive']:
                r.update(role=family+'_caption',level=None,title=None,status='review',relates_to=None,
                         possible_roles=[family+'_caption','body'],association_candidates=[v['block_id'] for v in choices])
                r['evidence'].append('ambiguous_object_pair' if len(choices)>1 else 'caption_role_requires_layout_or_local_decision');continue
            matches=choices
        current=r.get('relates_to')
        target=next((v for v in matches if v['block_id']==current),matches[0])
        target_row=rows[target['block_id']]
        r.update(role=family+'_caption',level=None,title=None,status='auto_rule',relates_to=target['block_id'],association_candidates=[v['block_id'] for v in matches])
        r.pop('level_pending',None)
        r['evidence'].append('adjacent_physical_object_and_'+('caption_style_or_sequence' if explicit else 'distinct_paragraph_layout'))
        descriptor=target_row.get('figure_layout')
        if descriptor:
            descriptor['caption_block']=b['block_id']
            labels=re.split(r'\t+| {2,}',text)
            if descriptor['kind']=='inline-row' and len(labels)==len(descriptor['drawings']) and all(v.strip() for v in labels):
                descriptor['labels']=[v.strip() for v in labels];r['subcaptions']=True
    # Resolve a dense series as a whole: a terminal unique caption can make
    # earlier two-sided candidates unique. Never force a genuinely isolated
    # tie, and never let two captions silently consume the same object.
    eligible={bid for bid,a in assessments.items() if a['positive'] and not a['prose'] and bid not in protected}
    for _ in range(len(top)+1):
        changed=False;claims={}
        for r in rows.values():
            if r['role'] in ('figure_caption','table_caption') and r.get('relates_to') and r.get('status') not in ('review','retained_unpaired'):
                claims.setdefault(r['relates_to'],[]).append(r['block_id'])
        occupied={obj:ids[0] for obj,ids in claims.items() if len(ids)==1}
        for obj,ids in claims.items():
            if len(ids)<2:continue
            for bid in ids:
                if bid not in eligible:continue
                r=rows[bid];r.update(status='review',relates_to=None,association_candidates=[obj])
                r['evidence'].append('competing_titles_for_one_object')
        proposals={}
        for bid in eligible:
            r=rows[bid]
            if r['role'] not in ('figure_caption','table_caption') or r.get('status')!='review' or r.get('relates_to'):continue
            original=r.get('association_candidates',[]);choices=[x for x in original if x not in occupied]
            if len(choices)==1 and len(original)>1:
                proposals.setdefault(choices[0],[]).append(bid)
        for target,ids in proposals.items():
            if len(ids)!=1:continue
            r=rows[ids[0]];r.update(status='auto_rule',relates_to=target,association_candidates=[target])
            r['evidence'].append('caption_series_unique_after_neighbor_binding');changed=True
        if not changed:break
    candidate['figure_layouts']=[dict(block_id=r['block_id'],**r['figure_layout']) for r in candidate['blocks'] if r.get('figure_layout')]
    return candidate

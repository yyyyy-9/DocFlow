"""Small rule baseline. Input is one fact scan; no corpus or truth imports."""
import argparse
from collections import Counter,defaultdict
import hashlib
import json
import re
from pathlib import Path
from .numbered_candidates import infer_numbered,rebuild_structure,authorize,list_introduction,caption_prefix,CONTRACT,VERSION

ARABIC=re.compile(r'^(\d+(?:\.\d+){0,3})\s+(?=\S)')
CHINESE=re.compile(r'^([一二三四五六七八九十百]+)、\s*')
SUB=re.compile(r'^[（(]([一二三四五六七八九十百]+)[）)]\s*')
CAP=re.compile(r'^(式)\s*\d+(?:[-－.]\d+)?(?:[.．、]\s*|\s+|[：:])')


def recognize(scan,overrides=None,short_content_chars=20):
    overrides=overrides or {}
    for bid,d in overrides.items():
        if d.get('source_sha256')!=scan['source_sha256'] or not d.get('reason') or d.get('role') not in ('heading','body','list_item'):
            raise ValueError('Override must bind source hash, role and reason')
        if d['role']=='heading' and d.get('level') not in range(1,6):raise ValueError('Confirmed heading level missing')
    all_blocks={b['block_id']:b for b in scan['blocks']};blocks=[all_blocks[i] for i in scan['top_level_ids']]
    if any(bid not in all_blocks for bid in overrides):raise ValueError('Unknown override block')
    numbered,groups=infer_numbered(blocks,short_content_chars,overrides)
    results=[];byid={};votes=defaultdict(Counter)
    for b in blocks:
        text=b.get('display_text','').strip();plain=b.get('plain_text','').strip();kinds=[s['kind'] for s in b.get('segments',[])]
        row={'block_id':b['block_id'],'source_locator':b['source_locator'],'role':'body','level':None,'title':None,
             'parent_id':None,'relates_to':None,'status':'auto_rule','evidence':[]}
        if b['kind']=='table':row['role']='table';row['evidence']=['physical_table']
        elif b['kind']!='paragraph':row['role']='boundary';row['evidence']=['physical_boundary']
        elif any(k in ('math_display','math_inline','legacy_equation') for k in kinds) and (not plain or re.fullmatch(r'[（(]\s*\d+\s*[）)]',plain)):
            row['role']='equation';row['evidence']=['standalone_native_math_with_optional_number_tail']
        elif any(k in ('drawing','legacy_picture') for k in kinds) and not plain:row['role']='figure';row['evidence']=['drawing_without_ordinary_text']
        elif not text and not kinds:row['role']='spacer';row['evidence']=['empty_paragraph']
        elif not text:row['role']='boundary';row['evidence']=['non_text_events']
        elif caption_prefix(text) or CAP.match(text):
            prefix=caption_prefix(text);marker=prefix['label'] if prefix else CAP.match(text).group(1)
            row['role']={'图':'figure_caption','表':'table_caption','式':'equation_description'}[marker]
            row['evidence']=['caption_prefix_pending_object_pair']
            if prefix:
                row['source_caption_prefix']=dict(prefix,text_basis='display_text')
                row.update(status='review',possible_roles=[row['role'],'body'])
                row['evidence']=['caption_prefix_is_lexical_candidate_only']
        elif re.match(r'^\[\d+\]',text):
            row['role']='reference';row['evidence']=['reference_marker']
        elif b['block_id'] in numbered:
            decision=numbered[b['block_id']];row.update(decision)
            o=row['number_candidate'];row['source_number_prefix']=o.get('prefix','')
            row['evidence']=[item for values in row['evidence_detail'].values() for item in values]
            if row['role']=='review':
                row.update(role='body',status='review',possible_roles=['heading','list_item','body'],possible_levels=list(range(1,6)))
            elif row['role']=='heading':
                row['title']=o['content']
                if row.get('level_pending'):
                    row.update(status='review',possible_roles=['heading'])
        else:
            level=None;prefix=''
            if re.match(r'^附录\s+[A-ZＡ-Ｚ]\b',text):level=1;row['evidence']=['appendix_marker']
            elif b.get('outline_level_raw') is not None and b['outline_level_raw']<5:
                level=b['outline_level_raw']+1;row['evidence']=['effective_outline_level']
            if level:
                row.update(role='heading',level=level,title=text[len(prefix):],source_number_prefix=prefix)
                votes[json.dumps(b.get('format_group'))][level]+=1
        results.append(row);byid[row['block_id']]=row
    # Cross-document format evidence can help unnumbered headings, but remains a rule.
    for i,(b,r) in enumerate(zip(blocks,results)):
        if r['role']!='body' or b['block_id'] in numbered:continue
        text=b.get('display_text','').strip()
        if list_introduction(text):continue
        if not 2<=len(text)<=32 or re.search('[，。；：！？,;!?]',text) or '[数学:' in text:continue
        following=next((x for x in results[i+1:] if x['role']!='spacer'),None)
        if following is None or following['role'] in ('equation','figure','table'):continue
        profile=votes[json.dumps(b.get('format_group'))]
        if (b.get('style_name') or '').lower() not in ('normal','正文') and len(profile)==1 and sum(profile.values())>=2:
            r.update(role='heading',level=next(iter(profile)),title=text,source_number_prefix='',evidence=['format_group_supported_by_numbered_headings'])
        else:
            r.update(role='heading',level=2,title=text,source_number_prefix='',status='review',
                     evidence=['short_standalone_text_without_sufficient_heading_evidence'],possible_roles=['heading','body'],possible_levels=list(range(1,6)))
    # Caption style alone never classifies a caption; compatible object is required.
    expected={'figure_caption':'figure','table_caption':'table','equation_description':'equation'}
    for i,r in enumerate(results):
        if r['role'] not in expected:continue
        if r.get('source_caption_prefix'):continue # Layout/role evidence is resolved by figure_layout.
        compatible=[]
        for direction in (-1,1):
            nonempty=0;j=i+direction
            while 0<=j<len(results) and nonempty<3:
                other=results[j]
                if other['role']=='heading':break
                if other['role']!='spacer':nonempty+=1
                if other['role']==expected[r['role']]:compatible.append((nonempty,abs(i-j),other['block_id']));break
                j+=direction
        r['association_candidates']=[x[2] for x in sorted(compatible)]
        if compatible:
            compatible.sort()
            if len(compatible)>1 and compatible[0][:2]==compatible[1][:2]:r['status']='review';r['evidence'].append('ambiguous_object_pair')
            else:r['relates_to']=compatible[0][2];r['evidence'].append('nearby_compatible_object')
        else:r['status']='review';r['evidence'].append('missing_compatible_object')
    for r in results:
        if r['block_id'] in overrides:
            d=overrides[r['block_id']]
            r.update(role=d['role'],level=d.get('level') if d['role']=='heading' else None,status='confirmed_override',override=dict(d))
            r.pop('level_pending',None)
            r['title']=r.get('number_candidate',{}).get('content',all_blocks[r['block_id']].get('display_text')) if r['role']=='heading' else None
            r['evidence'].append('source_bound_override_recomputed')
    rebuild_structure(results)
    authorize(results,scan)
    review=[{'block_id':r['block_id'],'text':all_blocks[r['block_id']].get('display_text'),
             'suggested_role':r['role'],'suggested_level':r['level'],'reason':r['evidence']} for r in results if r['status']=='review']
    return {'schema_version':'2.0-experimental','recognizer_version':VERSION,'source_sha256':scan['source_sha256'],
            'scan_version':scan['scan_version'],'blocks':results,'review':review,'model_calls':0,'number_groups':groups,'numbering_contract':CONTRACT,'overrides':overrides,
            'status':'REVIEW_REQUIRED' if review else 'RULE_CANDIDATES_READY','confirmed_ir':False,
            'note':'Candidate predictions only; no user confirmations or hidden truth inputs.'}

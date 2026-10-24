"""Contextual numbered-unit decisions and conservative section boundaries.

No fixture/truth imports, heading vocabulary, model calls or document writes.
Scores are evidence combinations, never pseudo probabilities.
"""
from copy import deepcopy
import hashlib
import json
import re

VERSION='0.5.7-numbered-units'
CONTRACT='numbered-units-2'
DATE=re.compile(r'^\s*\d{4}[-./年]\d{1,2}[-./月]\d{1,2}(?:日|\b|\s)')
QUANTITY=re.compile(r'^\s*[+−-]?\d+(?:\.\d+)?(?:[eE][+−-]?\d+)?\s*(?:[munpμµkMGT]?V|[munpμµkM]?A|mm|cm|nm|μm|µm|m|s|ms|Hz|kHz|MHz|GHz|MPa|Pa|W|kW|K|°C|℃|%|kg|g)(?=$|\s|[，。；,;])',re.I)
# Punctuation is presentation, not the source of list identity.
INTRO=re.compile(r'(?:^(?:下面|以下|下列|下述).*(?:步骤|清单|事项|操作|列举|顺序)|(?:按|依照).*(?:步骤|顺序)|(?:清单|步骤|事项|操作).*(?:如下|包括)$|(?:依次|列举|包括|如下)$)')
ACTION=re.compile(r'^(?:安装|连接|接通|打开|关闭|启动|停止|检查|记录|设置|调整|拆卸|执行|准备)(?!的)')
SENTENCE=re.compile(r'[。；;！？!?]\s*$')
CN='一二三四五六七八九十百'


def numeric_expression(text):
    """Inspect the complete leading numeric token BEFORE trying a heading span.

    Unknown Latin suffixes are ambiguous, not an expanding unit whitelist and
    not permission to consume a decimal. Ordinary Chinese titles still enter
    contextual recognition. This is deliberately a bounded lexical guard.
    """
    s=text.strip()
    if DATE.match(s):return 'date'
    # Protect a complete scalar + Chinese unit before splitting chapter paths.
    # A unit-derived title (毫米级测量方法) remains a reviewable ambiguity.
    units='平方米|平方厘米|平方毫米|立方米|立方厘米|立方毫米|摄氏度|开尔文|帕斯卡|吉赫兹|兆赫兹|千赫兹|赫兹|千米|毫米|厘米|微米|纳米|毫秒|微秒|纳秒|分钟|小时|毫伏|伏特|毫安|安培|千帕|兆帕|千瓦|瓦特|焦耳|牛顿|千克|毫克|毫升|分贝|弧度|米|秒|伏|安|帕|瓦|克|吨|升|度'
    scalar=re.match(r'^[+−-]?(?:\d+(?:\.\d*)?|\.\d+)',s)
    if scalar:
        tail=s[scalar.end():].lstrip()
        unit=re.match(r'(?:'+units+r')(?:\s*(?:/|每)\s*(?:'+units+r'))?',tail)
        if unit:
            remainder=tail[unit.end():].strip()
            if not remainder or re.match(r'^[，。；,:：!?！？]|^(?:为|是|约为|左右|以上|以下|以内|以外)(?=.|$)',remainder):
                return 'chinese_numeric_quantity'
            return 'ambiguous_numeric_text'

    atom=r'[+−-]?(?:\d+(?:\.\d*)?|\.\d+)'
    if re.match(r'^'+atom+r'(?:[eE][+−-]?\d+|\s*[×x·]\s*10\s*(?:\^\s*)?[+−-]?\d+)(?=$|\s|[^\d.])',s):
        return 'scientific_expression'
    if re.fullmatch(atom,s):return 'numeric_literal'
    if re.match(r'^'+atom+r'\s*(?:±|\+/-|[~～–—-]|至|到)\s*'+atom+r'(?=$|\s|[^\d.])',s):
        return 'numeric_range_or_uncertainty'
    if QUANTITY.match(s):return 'numeric_quantity'
    # Generic unit-like tokens, including compound units, must never be eaten
    # as a chapter prefix. Unknown words remain reviewable instead of guessing.
    m=re.match(r'^'+atom+r'\s*([A-Za-zμµΩ°℃%][A-Za-zμµΩ°℃%0-9/^·*⁻¹²³-]*)(?=$|\s|[，。；,;])',s)
    if m:
        if m.group(1) in ('dB','dBm','dBc') or re.search(r'[/^·*⁻¹²³]',m.group(1)):
            return 'numeric_quantity'
        return 'ambiguous_numeric_text'
    # The independent edit checker examines any Latin/unit initial after a
    # scalar, including a term joined directly to Chinese (FFT频谱分析).
    # Missing a token-ending boundary is not evidence that a number is safe
    # to replace. Use the existing review/resolution path; keep the stronger
    # quantity/date/scientific guards above and do not guess a heading term.
    if re.match(r'^'+atom+r'\s*[A-Za-zμµΩ°℃%]',s):
        return 'ambiguous_numeric_text'
    return None


def list_introduction(text):
    s=re.sub(r'[\s：:。.!！?？;；]+$','',text)
    return bool(INTRO.search(s)) and not manual_prefix(s) and len(s)<120


def digest(value):
    return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()


def cn_int(s):
    if s.isdigit():return int(s)
    digits={c:i for i,c in enumerate('零一二三四五六七八九')}
    total=0;current=0
    for c in s:
        if c in digits:current=digits[c]
        elif c in '十百':total+=(current or 1)*({'十':10,'百':100}[c]);current=0
        else:return None
    return total+current


def manual_prefix(text):
    """Span is in ordinary text, including leading spacing; decimals stay intact."""
    patterns=[
        (r'^\s*(\d+(?:\.\d+){1,4})\s*(?=[^\d\s.])','decimal_path'),
        (r'^\s*(\d+)([、．.)）])\s*(?=[^\d\s])','integer_delimited'),
        (r'^\s*(\d+)\s+(?=\S)','integer_space'),
        (r'^\s*(\d+)(?=[\u4e00-\u9fff])','integer_compact'),
        (r'^\s*(['+CN+r']+)、\s*','chinese'),
        (r'^\s*(['+CN+r']+)\s+(?=\S)','chinese_space'),
        (r'^\s*[（(](['+CN+r'\d]+)[）)]\s*','parenthesized'),
        (r'^\s*第(['+CN+r'\d]+)章\s*','chapter'),
    ]
    for pattern,kind in patterns:
        m=re.match(pattern,text)
        if not m or not text[m.end():].strip():continue
        nums=[int(x) for x in m.group(1).split('.')] if kind=='decimal_path' else [cn_int(m.group(1))]
        if any(n is None for n in nums):continue
        form=kind+(':'+m.group(2) if kind=='integer_delimited' else '')
        return {'kind':'manual','prefix':text[:m.end()],'span':[0,m.end()],'numbers':nums,
                'form':form,'content':text[m.end():]}
    return None


def observation(block):
    text=block.get('plain_text','');display=block.get('display_text',text)
    protected=numeric_expression(text)
    if protected:return {'protected':protected,'manual':None,'content':display}
    manual=manual_prefix(text);auto=block.get('numbering_effective',{})
    if auto.get('status') not in (None,'none'):
        return {'kind':'automatic','manual':manual,'prefix':'','span':[0,0],
                'numbers':auto.get('values') or [],'form':'auto:'+str(auto.get('num_id'))+':'+str(auto.get('ilvl')),
                'content':display,'automatic':auto,'conflict':'automatic_and_manual' if manual else 'unresolved_automatic_definition' if auto.get('status')!='resolved' else None}
    if manual:
        # A math/drawing placeholder is not included in the title length feature.
        manual['content']=display[len(manual['prefix']):] if display.startswith(manual['prefix']) else display
        return manual
    return None


def content_summary(blocks):
    result={'nonempty_blocks':0,'text_chars':0,'objects':0,'tables':0,'sentences':0}
    for b in blocks:
        text=b.get('plain_text','').strip()
        objects=sum(s['kind'] in ('drawing','math_inline','math_display') for s in b.get('segments',[]))
        table=b['kind']=='table'
        if not text and not objects and not table:continue
        result['nonempty_blocks']+=1;result['text_chars']+=len(re.sub(r'\s','',text))
        result['objects']+=objects;result['tables']+=int(table);result['sentences']+=int(bool(SENTENCE.search(text)))
    return result


def infer_numbered(blocks,short_chars=20,overrides=None):
    observations={};groups=[];active=[];previous=None;overrides=overrides or {};scope=None
    for i,b in enumerate(blocks):
        if b['kind']!='paragraph':continue
        obs=observation(b)
        decision=overrides.get(b['block_id'],{})
        if obs and obs.get('protected')=='ambiguous_numeric_text' and decision.get('role')=='heading' and decision.get('resolve_numeric_ambiguity') is True:
            prefix=manual_prefix(b.get('plain_text',''))
            if prefix and decision.get('origin') in ('model','user') and decision.get('reason') and decision.get('evidence_ids'):
                obs=prefix
                obs['ambiguity_resolution']=dict(decision)
        if not obs:
            outline=b.get('outline_level_raw')
            if outline is not None and 0<=outline<5 and decision.get('role') not in ('body','list_item'):
                # Unnumbered outline headings are real scope boundaries too.
                scope={'block_id':b['block_id'],'level':outline+1,'index':i}
                active=[];previous=None
            continue
        obs.update(block_id=b['block_id'],index=i,outline=b.get('outline_level_raw'),group_id=None,group_options=[])
        observations[b['block_id']]=obs
        if obs.get('protected'):continue
        nums=obs.get('numbers') or [0]
        before=next((x for x in reversed(blocks[:i]) if x.get('plain_text','').strip() or x['kind']=='table'),{})
        intro=list_introduction(before.get('plain_text','')) and before.get('outline_level_raw') is None
        signature=(obs['form'],tuple(nums[:-1]));matches=[]
        for g in active:
            if g['signature']!=signature:continue
            if obs['kind']=='automatic' or g['last']+1==nums[-1]:matches.append(g)
        anchor=decision.get('group_anchor_id')
        forced=decision.get('group_id')
        if anchor:
            matches=[g for g in active if any(m['block_id']==anchor for m in g['members'])]
            if len(matches)!=1:raise ValueError('Confirmed group anchor no longer resolves: '+anchor)
        elif forced:
            matches=[g for g in active if g['id']==forced]
            if not matches:raise ValueError('Confirmed group no longer exists: '+forced)
        elif intro:matches=[]
        if len(matches)>1 and matches[-1]['introduced_list']:
            # Compare non-crossing parses: finish the active nested list, then
            # resume its outer peer. A repeated outer ordinal immediately after
            # this candidate distinguishes that parse from closing the list now.
            next_obs=next((observation(x) for x in blocks[i+1:] if x['kind']=='paragraph' and observation(x) and not observation(x).get('protected')),None)
            if next_obs and next_obs.get('form')==obs['form'] and next_obs.get('numbers')==nums:
                obs['group_hypotheses']=[{'group_id':g['id'],'accepted':g is matches[-1],
                                        'reason':'complete_nested_sequence_before_outer_resumption' if g is matches[-1] else 'would_close_inner_sequence_before_its_next_item'} for g in matches]
                matches=[matches[-1]]
        if len(matches)>1:
            obs.update(group_options=[g['id'] for g in matches],conflict='competing_number_groups')
            for g in matches:g['conflicts'].append(b['block_id'])
            previous=obs
            continue
        if matches:
            group=matches[0]
            if obs['kind']=='manual':active=active[:active.index(group)+1]
        else:
            owner=None
            if nums[:-1]:
                owner=next((o['block_id'] for o in reversed(list(observations.values())[:-1]) if o.get('numbers')==nums[:-1] and (not scope or o['index']>scope['index'])),None)
            elif intro and previous and not re.search(r'新(?:的|一组|一套)|另(?:一组|一套)|重新',before.get('plain_text','')):owner=previous['block_id']
            elif previous and active and previous.get('form')!=obs['form'] and nums[-1]==1:
                # A hypothesis only: an outer peer must later close the interval
                # before a change of marker can establish a nested group.
                owner=previous['block_id']
            group={'id':f'g{len(groups)+1:04d}','signature':signature,'members':[],'introduced_list':intro,
                   'intro_block_id':before.get('block_id') if intro else None,'owner_candidate_id':owner,
                   'scope_anchor':deepcopy(scope),'last':nums[-1],'conflicts':[], 'reason':'list_introduction_and_restart' if intro else 'new_sequence'}
            groups.append(group);active.append(group)
        group['members'].append(obs);group['last']=nums[-1];obs['group_id']=group['id'];obs['group_options']=[group['id']]
        previous=obs
    for g in groups:
        owner=observations.get(g['owner_candidate_id'],{})
        parent=next((p for p in groups if p['id']==owner.get('group_id')),None)
        following=next((o for o in parent['members'] if o['index']>owner['index']),None) if parent else None
        g['closed_parent_interval']=bool(following and g['members'][-1]['index']<following['index'])
        g['explicit_path_parent']=bool(owner and len(g['members'][0].get('numbers',[]))>1 and owner.get('numbers') and g['members'][0]['numbers'][:-1]==owner['numbers'])
    # Evidence is measured using assigned peers, with alternative assignments explicit.
    decisions={}
    for group in groups:
        members=group['members'];spans=[]
        for j,o in enumerate(members):
            nxt=members[j+1] if j+1<len(members) else None
            limit=nxt['index'] if nxt else next((x['index'] for x in observations.values() if x['index']>o['index'] and x.get('group_id')!=group['id'] and not x.get('protected')),len(blocks))
            content=content_summary(blocks[o['index']+1:limit])
            o['following_content']=content;o['next_same_group']=nxt['block_id'] if nxt else None
            o['previous_same_group']=members[j-1]['block_id'] if j else None
            o['next_distance']=content if nxt else None
            o['next_boundary_reason']=None if nxt else 'no_confirmed_next_peer'
            o['short_content']=len(re.sub(r'\s','',o['content']))<short_chars
            o['text_projection_complete']=not any(s['kind'] in ('math_inline','math_display','drawing') for s in blocks[o['index']].get('segments',[]))
            spans.append(content)
        expanded=lambda d:bool(d['sentences'] or d['objects'] or d['tables'] or d['text_chars']>=40)
        child_evidence=any(g['owner_candidate_id'] in [o['block_id'] for o in members] and not g['introduced_list'] and (g['explicit_path_parent'] or g['closed_parent_interval']) for g in groups)
        peer_expansion=len(members)>=2 and all(expanded(s) for s in spans[:-1]) and not group['conflicts']
        dense=len(members)>=2 and all(s['nonempty_blocks']==0 for s in spans[:-1])
        short_peer_support=any(o['short_content'] for o in members)
        action_sequence=len(members)>=2 and all(ACTION.match(o['content']) for o in members)
        # Explicit-outline islands with a weak middle item are not sufficient to
        # confirm that middle item's role. The boundary dependency is retained.
        mixed_outline=any(o['outline'] is not None and o['outline']<5 for o in members) and any(o['outline'] is None for o in members)
        for o in members:
            outline=o['outline'];supported_outline=outline is not None and 0<=outline<5
            signs={'heading':[],'list':[],'conflicts':[]}
            if o['short_content']:signs['heading'].append('short_numbered_content')
            if expanded(o['following_content']):signs['heading'].append('content_expansion')
            if child_evidence:signs['heading'].append('compatible_child_numbering')
            if peer_expansion:signs['heading'].append('peer_content_units')
            if supported_outline:signs['heading'].append('explicit_effective_outline')
            if group['introduced_list']:signs['list'].append('explicit_list_introduction')
            if dense:signs['list'].append('adjacent_numbered_peers')
            if action_sequence:signs['list'].append('parallel_action_phrases')
            if o.get('conflict'):signs['conflicts'].append(o['conflict'])
            if group['conflicts']:signs['conflicts'].append('unresolved_group_membership')
            if supported_outline and group['introduced_list']:signs['conflicts'].append('outline_vs_list_context')
            if group['introduced_list'] and group['explicit_path_parent']:
                signs['conflicts'].append('full_numbered_path_parent_vs_list_introduction')
            if mixed_outline and not supported_outline:signs['conflicts'].append('mixed_outline_and_unmarked_peer')
            role='review';level=None
            if not signs['conflicts']:
                if group['introduced_list']:role='list_item'
                elif supported_outline:role='heading';level=outline+1
                elif any(SENTENCE.search(x['content']) for x in members):role='list_item'
                elif action_sequence and not child_evidence:
                    role='list_item' if dense else 'review'
                    signs['conflicts'].append('expanded_actions_without_explicit_scope')
                elif child_evidence or (peer_expansion and (short_peer_support or group['scope_anchor'])):
                    role='heading';level=len(o.get('numbers') or [1])
                elif group['owner_candidate_id'] in decisions and decisions[group['owner_candidate_id']]['role']=='heading' and len(o.get('numbers',[]))>1 and expanded(o['following_content']):
                    role='heading';level=len(o['numbers']);signs['heading'].append('confirmed_numbered_parent_and_expansion')
                elif dense:role='list_item'
            if o['block_id'] in overrides:
                confirmed=overrides[o['block_id']]
                role=confirmed['role'];level=confirmed.get('level') if role=='heading' else None
                signs['confirmation']=['source_bound_override']
            decisions[o['block_id']]={'role':role,'level':level,'evidence_detail':signs,'number_candidate':o,
                                       'number_group_id':o['group_id'],'group_options':o['group_options']}
    for bid,o in observations.items():
        if o.get('protected'):
            decisions[bid]={'role':'review' if o['protected']=='ambiguous_numeric_text' else 'body','level':None,'number_candidate':o,'evidence_detail':{'heading':[],'list':[],'conflicts':[],'protected':[o['protected']]},'number_group_id':None,'group_options':[]}
        elif bid not in decisions:
            decisions[bid]={'role':'review','level':None,'number_candidate':o,'evidence_detail':{'heading':[],'list':[],'conflicts':[o.get('conflict','insufficient_group_evidence')]},'number_group_id':None,'group_options':o['group_options']}
    # Resolve relative levels independently of roles. Marker shape is not level.
    for g in groups:
        owner=decisions.get(g['owner_candidate_id']);anchor=g['scope_anchor']
        explicit={o['outline']+1 for o in g['members'] if o['outline'] is not None and 0<=o['outline']<5}
        explicit.update(overrides[o['block_id']]['level'] for o in g['members'] if overrides.get(o['block_id'],{}).get('role')=='heading')
        for o in g['members']:
            d=decisions[o['block_id']]
            if d['role']!='heading' or o['block_id'] in overrides or (o['outline'] is not None and 0<=o['outline']<5):continue
            level=None;why=None
            if len(explicit)==1:level=next(iter(explicit));why='explicit_peer_level'
            elif owner:
                owner_outline=observations.get(g['owner_candidate_id'],{}).get('outline')
                anchored_owner=owner_outline is not None and 0<=owner_outline<5
                if owner['role']=='heading' and owner['level'] is not None and (g['explicit_path_parent'] or g['closed_parent_interval'] or anchored_owner):
                    level=owner['level']+1;why='confirmed_parent_group_interval'
            elif anchor:level=anchor['level']+1;why='effective_outline_scope'
            elif len(o.get('numbers',[]))==1:level=1;why='root_peer_group'
            if level is not None and 1<=level<=5:
                d['level']=level;d['evidence_detail']['heading'].append(why)
            else:
                d.update(level=None,level_pending=True,possible_levels=list(range(1,6)))
                d['evidence_detail']['conflicts'].append('relative_level_unresolved')
    summary=[{k:v for k,v in g.items() if k not in ('members','last','signature')} | {'members':[o['block_id'] for o in g['members']],'pattern':list(g['signature'])} for g in groups]
    return decisions,summary


def rebuild_structure(rows):
    """Uncertain future nodes affect ends; uncertain earlier nodes affect parents."""
    stack=[];uncertain=[]
    for r in rows:
        unresolved=r['status']=='review' and bool(set(r.get('possible_roles',[])) & {'heading'})
        r['role_status']='confirmed_rule' if r.get('level_pending') else 'pending' if r['status']=='review' else 'confirmed_rule' if r['status']=='auto_rule' else 'confirmed_override'
        r['level_status']='pending' if unresolved else 'confirmed' if r['role']=='heading' else 'not_applicable'
        if r['role']=='heading' and r['status']!='review':
            while stack and stack[-1]['level']>=r['level']:stack.pop()
            if r['level']==1:uncertain=[]
        r['parent_id']=stack[-1]['block_id'] if stack else None
        r['parent_status']='pending' if uncertain else 'confirmed'
        r['structure_requires_review']=bool(uncertain) or r['status']=='review'
        if unresolved:uncertain.append(r['block_id'])
        if r['role']=='heading' and r['status']!='review':
            stack.append(r)
            if r['level']==1:uncertain=[]
        r.pop('section_end_id',None);r.pop('section_boundary',None)
    for i,r in enumerate(rows):
        if r['role']=='heading' and r['status']=='review':
            r['section_boundary']={'status':'pending','end_id':None,'possible_end_ids':[], 'blocked_by':[r['block_id']],'reason':'role_or_level_pending'}
            r['section_end_id']=None
        if r['role']!='heading' or r['status']=='review':continue
        stop=len(rows);blocking=[]
        for j in range(i+1,len(rows)):
            other=rows[j]
            if other['status']=='review' and 'heading' in other.get('possible_roles',[]):
                # A lower-level possibility can alter descendant ranges/parentage;
                # only possible levels at or above this level can end this section.
                levels=other.get('possible_levels',[1,2,3,4,5])
                if any(l<=r['level'] for l in levels):blocking.append(other['block_id'])
            if other['role']=='heading' and other['status']!='review' and other['level']<=r['level']:
                stop=j;break
        end=stop-1
        while end>i and rows[end]['role'] in ('spacer','boundary','excluded_toc'):end-=1
        options=[rows[end]['block_id']]
        for bid in blocking:
            k=next(n for n in range(i+1,stop) if rows[n]['block_id']==bid)-1
            while k>i and rows[k]['role'] in ('spacer','boundary'):k-=1
            options.append(rows[k]['block_id'])
        boundary={'status':'pending' if blocking else 'confirmed','end_id':None if blocking else rows[end]['block_id'],
                  'possible_end_ids':list(dict.fromkeys(options)),'blocked_by':blocking,'stop_before_id':rows[stop]['block_id'] if stop<len(rows) else None}
        r['section_boundary']=boundary;r['section_end_id']=boundary['end_id'];r['structure_requires_review']|=bool(blocking)


def select_section(pred,block_id):
    row=next(r for r in pred['blocks'] if r['block_id']==block_id)
    boundary=row.get('section_boundary',{})
    if row['role']!='heading' or row['status']=='review' or boundary.get('status')!='confirmed' or row.get('parent_status')=='pending':
        raise ValueError('Section selection blocked by unresolved role, level, parent or boundary')
    start=pred['blocks'].index(row);end=next(i for i,r in enumerate(pred['blocks']) if r['block_id']==boundary['end_id'])
    return [r['block_id'] for r in pred['blocks'][start:end+1]]


def authorize(rows,scan):
    source={b['block_id']:b for b in scan['blocks']}
    for r in rows:
        r['list_edit']={'allowed':False,'reason':'manual_list_group_not_confirmed'}
        b=source[r['block_id']];o=r.get('number_candidate') or {};text=b.get('plain_text','')
        if r['role']=='list_item' and r['status']!='review' and o.get('kind')=='manual' and r.get('number_group_id') and not o.get('protected') and not o.get('conflict') and not numeric_expression(text):
            members=[x['block_id'] for x in rows if x.get('number_group_id')==r['number_group_id'] and x['role']=='list_item']
            auth=dict(allowed=True,source_sha256=scan['source_sha256'],block_id=r['block_id'],source_locator=r['source_locator'],text_sha256=hashlib.sha256(text.encode('utf8')).hexdigest(),span=o['span'],old=o['prefix'],numbers=o['numbers'],form=o['form'],group_anchor_id=members[0],group_members=members,decision_source=r['status'])
            auth['binding']=digest(auth);r['list_edit']=auth
        r['number_edit']={'allowed':False,'reason':'role_level_or_parent_pending'}
        if r['role']!='heading' or r['status']=='review' or r.get('level_status')!='confirmed' or r.get('parent_status')=='pending':continue
        b=source[r['block_id']];o=r.get('number_candidate') or {};text=b.get('plain_text','')
        numeric=numeric_expression(text)
        resolved=o.get('ambiguity_resolution')
        explicit_resolution=isinstance(resolved,dict) and resolved.get('resolve_numeric_ambiguity') is True
        if resolved is not None and not explicit_resolution:
            r['number_edit']['reason']='invalid_numeric_ambiguity_resolution'
            continue
        if o.get('protected') or o.get('conflict') or (numeric and not (numeric=='ambiguous_numeric_text' and explicit_resolution)):
            r['number_edit']['reason']=o.get('protected') or o.get('conflict') or numeric
            continue
        kind=o.get('kind','none');span=o.get('span',[0,0]);prefix=o.get('prefix','')
        if kind=='automatic' and o.get('manual'):continue
        r['number_edit']={'allowed':True,'kind':kind,'operation':'replace_text_prefix' if kind=='manual' else 'detach_automatic_numbering' if kind=='automatic' else 'insert_target_number',
                          'source_sha256':scan['source_sha256'],'block_id':r['block_id'],'source_locator':r['source_locator'],
                          'text_sha256':hashlib.sha256(text.encode('utf8')).hexdigest(),'span':span,'old':prefix,
                          'automatic':b.get('numbering_effective'),'level':r['level'],'decision_source':r['status'],'rules_version':VERSION}
        if resolved:r['number_edit']['ambiguity_resolution']=resolved
        r['number_edit']['binding']=digest(r['number_edit'])


def validate_edit(row,source_hash,source_text,automatic):
    a=row.get('number_edit',{})
    if not a.get('allowed'):raise ValueError('Heading has no separately bound number-edit authorization')
    check=dict(a);binding=check.pop('binding',None)
    if binding!=digest(check):raise ValueError('Number-edit authorization altered')
    if a.get('rules_version')!=VERSION:raise ValueError('Old number-edit rules; rescan and recognize source')
    if row.get('level_status')!='confirmed' or row.get('parent_status')=='pending':raise ValueError('Number-edit level or parent unresolved')
    if a['source_sha256']!=source_hash or a['block_id']!=row['block_id'] or a['source_locator']!=row['source_locator'] or a['level']!=row['level'] or a['decision_source']!=row['status']:
        raise ValueError('Number-edit decision binding mismatch')
    if a['text_sha256']!=hashlib.sha256(source_text.encode('utf8')).hexdigest():raise ValueError('Number-edit text fingerprint mismatch')
    lo,hi=a['span']
    if lo!=0 or source_text[lo:hi]!=a['old']:raise ValueError('Number-edit prefix span mismatch')
    numeric=numeric_expression(source_text)
    resolution=a.get('ambiguity_resolution')
    explicit_resolution=isinstance(resolution,dict) and resolution.get('resolve_numeric_ambiguity') is True
    if resolution is not None and not explicit_resolution:raise ValueError('Numeric ambiguity resolution requires explicit JSON true')
    if numeric and not (numeric=='ambiguous_numeric_text' and explicit_resolution):raise ValueError('Technical value/date or unresolved numeric text cannot be renumbered')
    if automatic!=a.get('automatic'):raise ValueError('Automatic numbering facts changed')
    if a['kind']=='automatic' and (hi or manual_prefix(source_text)):raise ValueError('Automatic number cannot authorize text deletion')
    return a


def validate_list_edit(row,source_hash,text):
    a=row.get('list_edit',{});copy=dict(a);binding=copy.pop('binding',None)
    observed=manual_prefix(text)
    if not a.get('allowed') or binding!=digest(copy):raise ValueError('List prefix has no bound migration authorization')
    if row['role']!='list_item' or row['status']=='review' or a['source_sha256']!=source_hash or a['block_id']!=row['block_id'] or a['source_locator']!=row['source_locator'] or a['decision_source']!=row['status'] or a['text_sha256']!=hashlib.sha256(text.encode('utf8')).hexdigest():raise ValueError('List authorization binding changed')
    if numeric_expression(text) or not observed:raise ValueError('List migration cannot consume a technical value')
    if observed['span']!=a['span'] or observed['prefix']!=a['old'] or observed['numbers']!=a['numbers'] or observed['form']!=a['form'] or text[:a['span'][1]]!=a['old']:raise ValueError('List prefix facts changed')
    return a

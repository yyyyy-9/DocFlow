"""Source-bound local-agent decisions; no invented human confirmations."""
from copy import deepcopy
import hashlib
import json
from .recognize_reports import predict
from .recognize_structure import recognize
from .numbered_candidates import rebuild_structure, authorize, manual_prefix, numeric_expression
from .make_review_packets import build_packets
from .job_support import VERSION, read, write, sha, stamp, check_inputs

ROLES = {'heading', 'body', 'list_item', 'document_title', 'front_heading',
         'figure_caption', 'table_caption', 'reference', 'equation_description'}


def validate_numeric_flag(decision):
    if 'resolve_numeric_ambiguity' in decision and type(decision['resolve_numeric_ambiguity']) is not bool:
        raise ValueError('resolve_numeric_ambiguity 必须是 JSON 布尔值 true/false（不加引号），不能是字符串、数字或 null；'
                         '请对该节点重新明确确认，不要自动转换旧值：' + str(decision.get('block_id')))


def validate_relation(decision, byid, known):
    target=decision.get('relates_to')
    if target is None:return
    if target not in known or target==decision.get('block_id'):raise ValueError('Caption target must be a different top-level source block')
    b=byid[target];role=decision['role']
    from .figure_layout import descendants
    valid=(role=='figure_caption' and b['kind'] in ('paragraph','table') and any(s['kind']=='drawing' for node in descendants(byid,b) for s in node.get('segments',[]))) or (role=='table_caption' and b['kind']=='table') or (role=='equation_description' and b['kind']=='paragraph' and any(s['kind'] in ('math_inline','math_display') for s in b.get('segments',[])))
    if not valid:raise ValueError('Caption relation does not match the physical source object type')


def resolve(scan, decisions):
    # Rebuild current rules and edit permissions from the source facts.
    byid = {b['block_id']: b for b in scan['blocks']}
    known = set(scan['top_level_ids'])
    seen=set()
    for d in decisions:
        validate_numeric_flag(d)
        bid=d.get('block_id')
        if bid not in known or bid in seen or byid[bid]['kind']!='paragraph':
            raise ValueError('Decision must address a unique top-level source paragraph')
        seen.add(bid)
        if d.get('role') not in ROLES or d.get('origin') not in ('model','user'):
            raise ValueError('Decision must retain a supported role and explicit model/user origin')
        if not d.get('reason') or not d.get('decision_id'):
            raise ValueError('Decision needs a reason and decision_id')
        if not d.get('evidence_ids') or not set(d['evidence_ids']) <= set(byid):
            raise ValueError('Decision needs actual evidence block IDs')
        if d['role']=='heading':
            if type(d.get('level')) is not int or not 1<=d['level']<=5:
                raise ValueError('Heading level must be 1..5')
            if numeric_expression(byid[bid].get('plain_text','')) not in (None,'ambiguous_numeric_text'):
                raise ValueError('A technical value/date cannot be turned into a chapter')
        validate_relation(d,byid,known)
    candidate = predict(scan)
    excluded={b['block_id'] for b in candidate['blocks'] if b.get('status')=='excluded_source_frontmatter'}
    if any(d['block_id'] in excluded for d in decisions):raise ValueError('Decision addresses excluded source front matter; revise the source-scope decision explicitly')
    overrides = {}
    for d in decisions:
        if d['role'] in ('heading', 'body', 'list_item'):
            overrides[d['block_id']] = dict(source_sha256=scan['source_sha256'],
                role=d['role'], level=d.get('level'), reason=d['reason'],
                origin=d['origin'], decision_id=d['decision_id'],
                evidence_ids=d['evidence_ids'],resolve_numeric_ambiguity=d.get('resolve_numeric_ambiguity',False),group_id=d.get('group_id'),group_anchor_id=d.get('group_anchor_id'))
    if overrides:
        candidate = predict(scan, overrides=overrides)
    rows = {b['block_id']: b for b in candidate['blocks']}
    for d in decisions:
        row = rows[d['block_id']]
        row.update(role=d['role'], level=d.get('level') if d['role']=='heading' else None,
                   status='model_review' if d['origin']=='model' else 'confirmed_user',
                   decision=dict(d))
        row.pop('level_pending', None)
        row['evidence'].append('explicit_source_bound_' + d['origin'] + '_decision')
        if d['role']=='heading':
            prefix = manual_prefix(byid[d['block_id']].get('plain_text', ''))
            row['title'] = (prefix or {}).get('content', byid[d['block_id']].get('display_text'))
        if 'relates_to' in d:
            row['relates_to'] = d['relates_to']
        if d['role'] not in ('figure_caption','table_caption','equation_description'):row['relates_to']=None
        elif not row.get('relates_to'):
            row['status']='retained_unpaired'
            row['evidence'].append('explicit_role_without_resolved_object_pair')
    for row in rows.values():
        if row.get('decision') and row['role']=='heading' and (row.get('number_candidate') or {}).get('conflict')=='competing_number_groups':
            row.update(status='review',possible_roles=['heading'],possible_levels=[row['level']])
            row['evidence'].append('number_group_requires_explicit_selection')
    for row in rows.values():validate_relation(row,byid,known)
    from .figure_layout import apply as apply_figure_layouts
    apply_figure_layouts(scan,candidate,overrides)
    rebuild_structure(candidate['blocks'])
    authorize(candidate['blocks'], scan)
    candidate['review'] = [dict(block_id=b['block_id'], reason=b['evidence'])
                           for b in candidate['blocks'] if b['status']=='review']
    candidate['status'] = 'REVIEW_REQUIRED' if candidate['review'] else 'RULE_CANDIDATES_READY'
    candidate['confirmed_ir'] = False
    return candidate


def request_bundle(run, sources):
    bundle = dict(version=VERSION, created=stamp(), sources=[],
        instruction='仅判断结构，不改原文。返回每个source的source_sha256、candidate_sha256、decisions。'
                    '每条包含block_id、role、level、reason、evidence_ids、origin(model或user)、decision_id。'
                    '不能把模型判断写成人工确认；证据不足可用role=defer并说明原因。')
    for src in sources:
        folder = run / 'sources' / src['id']
        if not (folder/'candidate.json').exists():
            continue
        candidate=read(folder/'candidate.json')
        if not isinstance(candidate.get('blocks'),list) or not isinstance(candidate.get('review'),list):
            continue
        scan=read(folder/'scan/scan.json')
        if candidate.get('source_sha256')!=src['sha256'] or scan.get('source_sha256')!=src['sha256']:
            raise ValueError('Review candidate does not match its frozen source: '+src['id'])
        packets = build_packets(scan, candidate)
        bundle['sources'].append(dict(source_id=src['id'], source_sha256=src['sha256'],
            candidate_sha256=sha(folder/'candidate.json'), requests=packets))
    write(run/'review_requests.json', bundle)
    return sum(len(s['requests']) for s in bundle['sources'])


def accept(run, state, response_path):
    check_inputs(run,state)
    response_raw = response_path.read_bytes()
    response = json.loads(response_raw.decode('utf-8-sig'))
    if not isinstance(response.get('sources'),list) or not response['sources']:
        raise ValueError('Review response must contain at least one source')
    original = {s['id']: s for s in state['sources']}
    prepared = []
    seen_sources = set()
    for src in response['sources']:
        sid = src['source_id']
        if sid in seen_sources or sid not in original:
            raise ValueError('Duplicate or unknown source in model response')
        seen_sources.add(sid)
        folder = run/'sources'/sid
        scan = read(folder/'scan/scan.json')
        if src['source_sha256'] != scan['source_sha256'] or src['candidate_sha256'] != sha(folder/'candidate.json'):
            raise ValueError('Model response belongs to a different source/candidate revision')
        byid = {b['block_id']: b for b in scan['blocks']}
        current=read(folder/'candidate.json')
        current_rows={b['block_id']:b for b in current['blocks']}
        current_groups={g['id']:g for g in current.get('number_groups',[])}
        known = set(scan['top_level_ids'])
        old = read(folder/'decisions.json') if (folder/'decisions.json').exists() else []
        merged = {d['block_id']: d for d in old}
        seen = set()
        deferred = []
        for d in src['decisions']:
            validate_numeric_flag(d)
            bid = d['block_id']
            if bid in seen or bid not in known or not d.get('reason'):
                raise ValueError('Unknown/duplicate node or missing reason')
            seen.add(bid)
            if current_rows[bid].get('status')=='excluded_source_frontmatter':raise ValueError('Excluded source front matter is not an active structure question')
            if d.get('origin') not in ('model', 'user') or not d.get('decision_id'):
                raise ValueError('Decision needs explicit origin and decision_id')
            if not d.get('evidence_ids') or not set(d['evidence_ids']) <= set(byid):
                raise ValueError('Decision needs actual evidence block IDs')
            if d['role']=='defer':
                deferred.append(d)
                continue
            if d['role'] not in ROLES or byid[bid]['kind']!='paragraph':
                raise ValueError('Unsupported role or non-paragraph role change')
            if d['role']=='heading' and (type(d.get('level')) is not int or not 1<=d['level']<=5):
                raise ValueError('Heading level must be 1..5')
            if d['role']=='heading' and numeric_expression(byid[bid].get('plain_text','')) not in (None,'ambiguous_numeric_text'):
                raise ValueError('A technical value/date cannot be turned into a chapter')
            validate_relation(d,byid,known)
            chosen=deepcopy(d)
            # Group ordinals can change when an earlier false heading is
            # corrected. Bind the choice to an existing source member instead.
            if d.get('group_id'):
                gid=d['group_id']
                if gid not in current_rows[bid].get('group_options',[]) or gid not in current_groups:
                    raise ValueError('Group selection was not offered for this candidate')
                peers=[m for m in current_groups[gid]['members'] if m!=bid]
                if not peers:raise ValueError('Group choice requires an existing peer anchor')
                chosen['group_anchor_id']=peers[0]
            elif d.get('group_anchor_id'):
                raise ValueError('Submit an offered group_id; anchor is derived by the receiver')
            merged[bid] = chosen
        candidate = resolve(scan, list(merged.values()))
        prepared.append((folder, list(merged.values()), candidate, deferred))
    # Validate the complete submitted batch before mutating any candidate.
    for folder, decisions, candidate, deferred in prepared:
        write(folder/'decisions.json', decisions)
        write(folder/'candidate.json', candidate)
        write(folder/'deferred.json', deferred)
    return dict(status='DONE', source_count=len(prepared),
                response_sha256=hashlib.sha256(response_raw).hexdigest(),accepted_response=response)

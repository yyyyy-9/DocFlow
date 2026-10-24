"""Locate user feedback in a hash-bound output and create a fresh revision."""
from pathlib import Path
import math
from uuid import uuid4
from .job_support import read, write, sha, stamp, contained


def locate(run, document, quote=None, locator=None, problem='', wanted='', before=None, origin=None):
    run=Path(run).resolve()
    if origin not in ('model','user'):
        raise ValueError('Feedback needs its actual origin: model or user')
    state=read(run/'state.json')
    doc=next((d for d in state['deliverables'] if d['id']==document),None)
    if not doc:raise ValueError('Unknown output document')
    if sha(contained(run,doc['file']))!=doc['sha256']:
        raise ValueError('Output was edited outside the workflow. Preserve it and re-index before automatic repair.')
    indexpath=contained(run,doc['locator_index'])
    if doc.get('locator_index_sha256') and sha(indexpath)!=doc['locator_index_sha256']:
        raise ValueError('Locator index changed outside the workflow')
    rows=read(indexpath)
    if not isinstance(rows,list):
        raise ValueError('This output has no successful locator index; preserve its failure receipt')
    if not quote and not locator:raise ValueError('Provide copied text or an existing locator')
    matches=[]
    for i,row in enumerate(rows):
        if locator and row['locator']!=locator:continue
        if quote and quote not in row['text']:continue
        if before and not any(before in x['text'] for x in rows[max(0,i-3):i]):continue
        matches.append(dict(row,previous=[x['text'][:160] for x in rows[max(0,i-2):i]],
                            next=[x['text'][:160] for x in rows[i+1:i+3]]))
    return dict(feedback_id='FB-'+uuid4().hex[:10],created=stamp(),run=str(run.resolve()),
        document=document,output_sha256=doc['sha256'],quote=quote,problem=problem,wanted=wanted,
        origin=origin,scope='this_occurrence',status='LOCATED' if len(matches)==1 else 'AMBIGUOUS' if matches else 'NOT_FOUND',
        matches=matches,action=None,
        instruction='只有唯一匹配可自动修订。页码不作主键。范围默认仅此处；扩为同类或长期规则需明确记录。')


def apply(feedback_path,dest):
    from .workflow import revise, save, finish, eligible_sources, refresh_preparation_status
    from .model_review import resolve, ROLES
    from .numbered_candidates import numeric_expression
    from .job_support import event
    dest=Path(dest).resolve()
    f=read(feedback_path);parent=Path(f['run']).resolve();old=read(parent/'state.json')
    document=next(d for d in old['deliverables'] if d['id']==f['document'])
    if (sha(contained(parent,document['file']))!=f['output_sha256'] or document['sha256']!=f['output_sha256']
            or f['status']!='LOCATED' or len(f['matches'])!=1):
        raise ValueError('Feedback is stale or not uniquely located')
    if f.get('scope')!='this_occurrence':
        raise ValueError('Expand scope as an explicit reviewed change plan; no automatic global replacement')
    action=f.get('action') or {}
    if action.get('kind') not in ('structure','layout'):
        raise ValueError('Set action.kind to structure or layout; math/content/code issues need a documented repair task')
    if not f.get('problem') or not f.get('wanted'):raise ValueError('Problem and wanted result are required')
    if f.get('origin') not in ('model','user'):raise ValueError('Feedback origin must be explicit')
    hit=f['matches'][0]
    # The edited output has to map to the same source node in this revision.
    indexpath=contained(parent,document['locator_index'])
    if document.get('locator_index_sha256') and sha(indexpath)!=document['locator_index_sha256']:
        raise ValueError('Locator index changed outside the workflow')
    live=next((r for r in read(indexpath) if r['locator']==hit['locator']),None)
    if live is None or any(hit.get(k)!=v for k,v in live.items()):
        raise ValueError('Feedback locator was altered')
    sid,bid=hit['source_id'],hit['block_id']
    if sid not in {s['id'] for s in old['sources']}:
        raise ValueError('Feedback source is not part of this run')
    scan=read(parent/'sources'/sid/'scan/scan.json');blocks={b['block_id']:b for b in scan['blocks']}
    if bid not in scan['top_level_ids']:raise ValueError('Feedback node is not a top-level source block')
    if blocks[bid].get('kind')!='paragraph':raise ValueError('Feedback currently supports whole source paragraphs')
    # Validate all caller-controlled action fields before creating a revision.
    if action['kind']=='structure':
        if set(action)-{'kind','role','level'}:raise ValueError('Unsupported structure correction field')
        role=action.get('role');level=action.get('level')
        if role not in ROLES or (role=='heading' and (type(level) is not int or not 1<=level<=5)):
            raise ValueError('Unsupported structure correction')
        if role=='heading' and numeric_expression(blocks[bid].get('plain_text','')) not in (None,'ambiguous_numeric_text'):
            raise ValueError('Technical quantities cannot be renumbered as headings')
    else:
        validate_layout(action)
    state=revise(parent,dest)
    if sid not in {s['id'] for s in eligible_sources(dest,state)}:
        raise ValueError('Source preparation failed in new revision; review its receipts before applying feedback')
    folder=dest/'sources'/sid
    scan=read(folder/'scan/scan.json')
    if action['kind']=='structure':
        decisions=read(folder/'decisions.json') if (folder/'decisions.json').exists() else []
        decisions=[d for d in decisions if d['block_id']!=bid]
        decisions.append(dict(block_id=bid,role=role,level=level,origin=f['origin'],
            decision_id=f['feedback_id'],evidence_ids=[bid],reason=f['problem']+' → '+f['wanted']))
        candidate=resolve(scan,decisions)
        write(folder/'decisions.json',decisions)
        write(folder/'candidate.json',candidate)
        state['stages']['recognize:'+sid]['evidence_sha256']=sha(folder/'candidate.json')
        state['stages']['recognize:'+sid]['status']='FEEDBACK_APPLIED'
    else:
        state.setdefault('layout_overrides',[]).append(dict(document=f['document'],source_id=sid,block_id=bid,
            feedback_id=f['feedback_id'],origin=f['origin'],settings=action,reason=f['wanted']))
    refresh_preparation_status(dest,state)
    state['feedback_applied']=dict(id=f['feedback_id'],parent_output_sha256=f['output_sha256'],
                                  locator=hit['locator'],action=action)
    write(dest/'feedback.json',f)
    save(dest,state)
    event(dest,'feedback','APPLIED_TO_NEW_REVISION',origin=f['origin'],feedback_id=f['feedback_id'],locator=hit['locator'])
    result=finish(dest,state,retain=True)
    result['feedback_result']=dict(feedback_id=f['feedback_id'],
        status='NEEDS_REPAIR' if result['status']=='NEEDS_REPAIR' else 'REBUILT_CHECK_RESULTS_ATTACHED',
        run_status=result['status'],changed_locator=hit['locator'],
        affected=('structure/numbering/dependent boundaries and all outputs containing the source'
                  if action['kind']=='structure' else 'layout in selected document: '+f['document']))
    save(dest,result)
    return result


def validate_layout(action):
    allowed={'kind','alignment','first_line_pt','space_before_pt','space_after_pt','line_multiple','font_pt'}
    if set(action)-allowed:raise ValueError('Unsupported layout field')
    if action.get('alignment','left') not in ('left','center','right','both'):raise ValueError('Invalid alignment')
    for key in ('first_line_pt','space_before_pt','space_after_pt','font_pt','line_multiple'):
        if key in action and (type(action[key]) not in (int,float) or not math.isfinite(action[key]) or not 0<=action[key]<=144):
            raise ValueError('Invalid layout value: '+key)
    if 'font_pt' in action and not 5<=action['font_pt']<=48:raise ValueError('Font size outside 5..48 pt')
    if 'line_multiple' in action and not 0.5<=action['line_multiple']<=3:raise ValueError('Line spacing outside range')


def apply_layout(asm,overrides,document=None):
    from .docx_formatting import q,elt,NS
    from copy import deepcopy
    from lxml import etree as E
    applied=[]
    for correction in overrides:
        if correction.get('document')!=document:continue
        settings=correction['settings'];validate_layout(settings)
        rows=[r for r in asm.records if r['source_id']==correction['source_id'] and r['block_id']==correction['block_id']]
        for row in rows:
            node=asm.body[row['output_index']]
            if node.tag!=q('w:p'):raise ValueError('Layout feedback currently supports whole paragraphs')
            before=deepcopy(node)
            pp=node.find(q('w:pPr'))
            if pp is None:pp=elt('w:pPr');node.insert(0,pp)
            def prop(name,**attrs):
                e=pp.find(q('w:'+name))
                if e is None:e=elt('w:'+name);pp.append(e)
                for k,v in attrs.items():e.set(q('w:'+k),str(v))
            if 'alignment' in settings:prop('jc',val=settings['alignment'])
            if 'first_line_pt' in settings:
                prop('ind',firstLine=round(settings['first_line_pt']*20),firstLineChars=0,hanging=0,hangingChars=0)
                if settings['first_line_pt']:
                    pp.find(q('w:ind')).attrib.pop(q('w:firstLineChars'),None)
            for key,attr in (('space_before_pt','before'),('space_after_pt','after')):
                if key in settings:prop('spacing',**{attr:round(settings[key]*20)})
            if 'line_multiple' in settings:prop('spacing',line=round(settings['line_multiple']*240),lineRule='auto')
            if 'font_pt' in settings:
                for r in node.xpath('.//w:r[not(ancestor::m:oMath)]',namespaces=NS):
                    rp=r.find(q('w:rPr'))
                    if rp is None:rp=elt('w:rPr');r.insert(0,rp)
                    for key in ('sz','szCs'):
                        for e in list(rp.findall(q('w:'+key))):rp.remove(e)
                        rp.append(elt('w:'+key,val=round(settings['font_pt']*2)))
            def content(n):
                n=deepcopy(n)
                for e in list(n.xpath('.//w:pPr|.//w:rPr',namespaces=NS)):e.getparent().remove(e)
                return E.tostring(n)
            if content(before)!=content(node):raise ValueError('Layout correction changed content')
            applied.append(dict(correction,output_index=row['output_index'],content_unchanged=True))
    return applied

"""Draft first, then deliver cumulative review batches through the original CLI.

No document parser, semantic judgments or model API calls live here.
"""
from pathlib import Path
from datetime import datetime, timezone
import argparse, hashlib, json, os, subprocess, sys, time

ROOT=Path(__file__).resolve().parent
sys.dont_write_bytecode=True

def read(path):return json.loads(Path(path).read_text('utf-8-sig'))
def digest(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def review_history(run,scans,candidates):
    """Recover accepted judgments/explicit body fallbacks from verified lineage.

    Read actual accepted responses, not an editable progress counter. A fallback
    stays unresolved in the core candidate, so assembly preserves its text/number.
    """
    seen=set();visited=set();fallbacks=[];resolved=0;current=Path(run);accepted_items=[];lineage=[]
    while True:
        current=current.resolve()
        if current in visited:raise ValueError('Revision lineage cycle')
        visited.add(current);state=read(current/'state.json');lineage.append((current,state))
        events=[json.loads(line) for line in (current/'events.jsonl').read_text('utf8').splitlines() if line.strip()]
        accepted={e['round']:e for e in events if e.get('stage')=='local_review' and e.get('status')=='APPLIED'}
        for n in range(state.get('review_rounds',0),0,-1):
            path=current/f'response_{n:03d}.json'
            event=accepted.get(n,{})
            expected=event.get('stored_response_sha256',event.get('response_sha256'))
            actual=digest(path)
            if expected!=actual:
                # Legacy CLI stored normalized JSON but recorded the input byte
                # hash. Only accept a concrete known serialization whose hash
                # matches that event; never waive payload binding.
                text=path.read_text('utf8')
                variants=[(text+'\n').encode('utf8'),(text+'\n').replace('\n','\r\n').encode('utf8')]
                if event.get('stored_response_sha256') or not any(hashlib.sha256(v).hexdigest()==expected for v in variants):
                    raise ValueError('Accepted response differs from its event: '+str(path))
            for source in read(path)['sources']:
                scan=scans.get(source['source_id'])
                if not scan or scan['source_sha256']!=source['source_sha256']:continue
                blocks={b['block_id']:b for b in scan['blocks']}
                for d in source['decisions']:
                    key=(source['source_id'],d['block_id'])
                    if key in seen:continue
                    seen.add(key)
                    if d['role']!='defer':
                        resolved+=1;accepted_items.append(dict(source_id=key[0],block_id=key[1],source_sha256=source['source_sha256'],
                            disposition=d['role'],accepted_run=str(current),response_sha256=actual));continue
                    if d.get('fallback')!='body':continue
                    b=blocks[d['block_id']]
                    row=next(r for r in candidates[key[0]]['blocks'] if r['block_id']==key[1])
                    if row['status']!='review':raise ValueError('Body fallback cannot replace an already accepted role: '+str(key))
                    if (b['kind']!='paragraph' or any(s['kind'] in ('drawing','math_inline','math_display') for s in b.get('segments',[]))):
                        raise ValueError('Body fallback is only for text paragraphs: '+str(key))
                    fallbacks.append({'source_id':key[0],'block_id':key[1],'source_sha256':source['source_sha256'],
                        'reason':d['reason'],'response':str(path),'response_sha256':actual,
                        'text':b.get('display_text','')[:200],'presentation':'body','recognition_correct':None})
                    accepted_items.append(dict(source_id=key[0],block_id=key[1],source_sha256=source['source_sha256'],
                        disposition='body_fallback',accepted_run=str(current),response_sha256=actual))
        parent=state.get('parent_revision')
        if not parent:break
        parent=Path(parent)
        if digest(parent/'state.json')!=state['parent_state_sha256']:raise ValueError('Parent revision changed')
        current=parent
    return resolved,fallbacks,accepted_items,lineage


def review_ledger(items,lineage):
    """Close semantic requests only against checked, source-bound drafts.

    Accepted-but-not-built items are execution work, never new model questions.
    No decision, input, event or prior output is deleted when an item closes.
    """
    ledger=[];cache={}
    for item in items:
        row=dict(item);key=(row['source_id'],row['block_id'])
        row['item_id']='REV_'+hashlib.sha256((row['source_sha256']+':'+':'.join(key)).encode()).hexdigest()[:24]
        row.update(status='awaiting_generation',documents=[])
        for home,state in lineage:
            docs=state.get('deliverables',[])
            if docs:
                selected=[d for d in docs if key[0] in d.get('selected_source_ids',[])]
                expected=1+int(bool(state['job'].get('make_master',True)) and any(s['id']==key[0] and s['include_in_master'] for s in state['sources']))
                okay=len(selected)==expected
                for doc in selected:
                    path=home/doc['file'];report=home/doc['assembly_report'];cache_key=str(report)
                    if cache_key not in cache:cache[cache_key]=read(report)
                    report_data=cache[cache_key];source=next((s for s in report_data['sources'] if s['source_id']==key[0]),{})
                    mapping=next((m for m in report_data['mapping'] if (m['source_id'],m['block_id'])==key),None)
                    good=(doc.get('checks',{}).get('assembly')=='PASS' and doc.get('checks',{}).get('content')=='PASS'
                          and key[0] in doc.get('source_ids',[]) and source.get('sha256')==row['source_sha256']
                          and path.is_file() and digest(path)==doc['sha256'] and mapping is not None)
                    if row['disposition']=='body_fallback':good=good and mapping is not None and mapping['role']=='body'
                    okay=okay and good
                    row['documents'].append(dict(file=str(path),applied=bool(good)))
                row['status']='closed_in_checked_draft' if okay else 'execution_issue'
                row['applied_run']=str(home);break
            if str(home)==row['accepted_run']:break
        ledger.append(row)
    return {'schema':'wps-review-ledger-v1','items':ledger,
            'closed':sum(x['status']=='closed_in_checked_draft' for x in ledger),
            'awaiting_generation':sum(x['status']=='awaiting_generation' for x in ledger),
            'execution_issues':sum(x['status']=='execution_issue' for x in ledger),
            'meaning':'Closed removes the active semantic request, not its evidence. Execution failures are reported separately and are never sent back for the same judgment.'}

def review_plan(run,batch_size=50):
    """Classify bounded source-bound packets; never decide a semantic role.

    Keep all structural questions of one source together. Caption families stay
    together. A complete dependency unit may exceed the requested batch size.
    """
    if batch_size<1:raise ValueError('batch-size must be positive')
    from collections import Counter,defaultdict
    from copy import deepcopy
    sys.path.insert(0,str(ROOT/'src'))
    from wps_pipeline.make_review_packets import build_packets
    run=Path(run);state=read(run/'state.json');scans={};candidates={}
    for source in state['sources']:
        home=run/'sources'/source['id']
        if not (home/'candidate.json').is_file():continue
        scans[source['id']]=read(home/'scan/scan.json');candidates[source['id']]=read(home/'candidate.json')
    resolved,fallbacks,accepted,lineage=review_history(run,scans,candidates)
    held={(f['source_id'],f['block_id']) for f in accepted};ledger=review_ledger(accepted,lineage)
    units=[];extra_count=0;runtime_pending=0
    for source in state['sources']:
        sid=source['id']
        if sid not in candidates:continue
        candidate=candidates[sid];prediction=deepcopy(candidate)
        runtime_pending+=len(candidate['review']);queued={r['block_id'] for r in candidate['review']}
        pairs=defaultdict(list)
        for row in candidate['blocks']:
            if row['role'] in ('figure_caption','table_caption') and row.get('relates_to') and not row.get('subcaptions'):
                pairs[(row['role'],row['relates_to'])].append(row)
        # A repeated pairing is a request to inspect, not a blanket rejection.
        for rows in pairs.values():
            if len(rows)<2:continue
            for row in rows:
                if row.get('decision') or row['block_id'] in queued:continue
                prediction['review'].append({'block_id':row['block_id'],'reason':['same_object_multiple_whole_captions']})
                queued.add(row['block_id']);extra_count+=1
        groups=defaultdict(list)
        for packet in build_packets(scans[sid],prediction):
            q=packet['request'];b=q['candidate'];key=(sid,b['id'])
            if key in held:continue
            row=next(r for r in candidate['blocks'] if r['block_id']==b['id'])
            category='captions' if row['role'] in ('figure_caption','table_caption','equation_description') else 'structure'
            groups[category].append(packet)
        for category,packets in groups.items():
            if category=='structure' and any(len(p['request']['candidate']['group_options'])>1 or
                    any(e in ('competing_number_groups','unresolved_group_membership') for e in p['request']['candidate']['existing_evidence']) for p in packets):
                category='number_groups'
            units.append({'category':category,'source_id':sid,'source_sha256':scans[sid]['source_sha256'],
                'candidate_sha256':digest(run/'sources'/sid/'candidate.json'),'requests':packets})
    batches=[]
    for category in ('number_groups','structure','captions'):
        batch=[];count=0
        for unit in [u for u in units if u['category']==category]:
            size=len(unit['requests'])
            if batch and count+size>batch_size:
                batches.append({'category':category,'sources':batch,'count':count});batch=[];count=0
            batch.append(unit);count+=size
        if batch:batches.append({'category':category,'sources':batch,'count':count})
    folder=run/'review_batches';folder.mkdir(exist_ok=True);round_no=state.get('review_rounds',0);files=[]
    for i,batch in enumerate(batches,1):
        path=folder/f'r{round_no:03d}_batch_{i:03d}.json'
        payload={'schema':'wps-review-batch-v1','batch_size':batch_size,**batch,
                 'instruction':'Only judge these source-bound local packets. No truth files. Use defer with fallback=body only when text heading level remains uncertain; keep captions/objects unresolved explicitly.'}
        path.write_text(json.dumps(payload,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
        files.append({'category':batch['category'],'count':batch['count'],'file':str(path),
                      'source_ids':[s['source_id'] for s in batch['sources']]})
    result={'schema':'wps-review-plan-v1','batch_size':batch_size,'active_requests':sum(b['count'] for b in batches),
        'runtime_unresolved':runtime_pending,'additional_pairing_requests':extra_count,'resolved_decisions':resolved,
        'body_fallback_count':len(fallbacks),'body_fallbacks':fallbacks,'batches':files,
        'closed_review_items':ledger['closed'],'awaiting_generation':ledger['awaiting_generation'],'execution_issues':ledger['execution_issues'],
        'meaning':'Body fallback is handled presentation, not confirmed recognition; raw unresolved items remain in core receipts.'}
    ledger_path=run/f'review_ledger_{round_no:03d}.json';ledger_path.write_text(json.dumps(ledger,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    result['review_ledger']=str(ledger_path)
    path=run/f'review_plan_{round_no:03d}.json';path.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    return dict(result,path=str(path))

def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    commands=parser.add_subparsers(dest='command',required=True)
    draft=commands.add_parser('draft',help='扫描、保留疑点的首稿、检查与status连续执行')
    selection=draft.add_mutually_exclusive_group(required=True)
    selection.add_argument('--job',type=Path);selection.add_argument('--input-dir',type=Path)
    draft.add_argument('--run',type=Path,required=True);draft.add_argument('--template',type=Path);draft.add_argument('--title')
    prepare=commands.add_parser('prepare-review',help='另建评审任务，不生成第二稿')
    prepare.add_argument('parent',type=Path);prepare.add_argument('--run',type=Path,required=True)
    complete=commands.add_parser('complete-review',help='集中提交一次响应后统一生成第二稿')
    complete.add_argument('run',type=Path);complete.add_argument('response',type=Path)
    checkpoint=commands.add_parser('checkpoint',help='接受本批判断并交阶段稿；可自动准备下一阶段')
    checkpoint.add_argument('run',type=Path);checkpoint.add_argument('response',type=Path)
    checkpoint.add_argument('--next-run',type=Path)
    for cmd in (draft,prepare,complete,checkpoint):cmd.add_argument('--batch-size',type=int,default=50)
    status=commands.add_parser('status',help='只读检查已有任务');status.add_argument('run',type=Path)
    args=parser.parse_args(argv);run=args.run.resolve();events=[];started=time.perf_counter();planning=None;next_run=None
    existed_on_entry=run.exists()

    def step(name,*parameters):
        cmd=[sys.executable,'-X','utf8','-B',str(ROOT/'report_tool.py'),*map(str,parameters)]
        env=dict(os.environ,PYTHONUTF8='1',PYTHONDONTWRITEBYTECODE='1')
        t=time.perf_counter();cp=subprocess.run(cmd,capture_output=True,env=env)
        try:value=json.loads(cp.stdout.decode('utf-8-sig'))
        except (UnicodeError,ValueError):value={'status':'ERROR','message':'Child returned invalid JSON'}
        event={'stage':name,'command':cmd,'returncode':cp.returncode,
               'wall_seconds':round(time.perf_counter()-t,6),'status':value.get('status')};events.append(event)
        if args.command!='status' and (run/'state.json').is_file():
            logs=run/'batch_steps';logs.mkdir(exist_ok=True)
            prefix=f'{len(list(logs.glob("*.json")))+1:03d}_{name}'
            (logs/(prefix+'.json')).write_bytes(cp.stdout)
            if cp.stderr:(logs/(prefix+'.stderr.log')).write_bytes(cp.stderr)
            event['log']=str((logs/(prefix+'.json')).relative_to(run))
        if value.get('status')=='ERROR' or value.get('receipt_integrity')=='FAIL':
            raise ValueError(value.get('message') or value.get('integrity_errors') or 'Task receipt integrity failed')
        if cp.returncode and value.get('status')!='NEEDS_REPAIR':
            raise ValueError(value.get('message') or 'Stage failed: '+name)
        return value

    result={}
    try:
        if args.command in ('draft','prepare-review') and run.exists():raise FileExistsError('Choose a new run directory: '+str(run))
        if getattr(args,'batch_size',50)<1:raise ValueError('batch-size must be positive')
        if args.command=='checkpoint' and args.next_run:
            candidate=args.next_run.resolve()
            if candidate.exists() or candidate==run or candidate.is_relative_to(run):raise ValueError('Next stage requires a new sibling/outside directory')
        if args.command=='draft':
            params=['start']
            if args.job:params.append(args.job.resolve())
            else:params.extend(['--input-dir',args.input_dir.resolve()])
            params.extend(['--run',run,'--master'])
            if args.template:params.extend(['--template',args.template.resolve()])
            if args.title:params.extend(['--title',args.title])
            step('start',*params);step('finish','finish',run,'--retain-unresolved');result=step('status','status',run)
        elif args.command=='prepare-review':
            result=step('revise','revise',args.parent.resolve(),'--run',run)
            if result.get('deliverables'):raise ValueError('Preparation unexpectedly generated documents')
        elif args.command in ('complete-review','checkpoint'):
            # A rejected response never falls through to a rebuild.
            step('review','review',run,args.response.resolve())
            step('finish','finish',run,'--retain-unresolved');result=step('status','status',run)
        else:result=step('status','status',run)
        if args.command!='status':
            planning=review_plan(run,args.batch_size)
            if args.command=='checkpoint' and args.next_run and planning['active_requests'] and result.get('status')!='NEEDS_REPAIR':
                step('prepare_next','revise',run,'--run',args.next_run.resolve())
                next_run={'run':str(args.next_run.resolve()),'review_plan':review_plan(args.next_run.resolve(),args.batch_size)}
    except Exception as exc:
        # Keep existing draft paths visible even when a later command fails.
        saved={}
        if (run/'state.json').is_file():
            try:saved=json.loads((run/'state.json').read_text('utf-8-sig'))
            except (OSError,ValueError):pass
        result={'status':'ERROR','message':str(exc),'deliverables':saved.get('deliverables',[]),
                'preserved_task_status':saved.get('status')}
    receipt={'schema':'wps-batch-flow-v1','command':args.command,'run':str(run),
             'finished':datetime.now(timezone.utc).isoformat(),'wall_seconds':round(time.perf_counter()-started,6),
             'events':events,'result':result,'model_api_calls':0,'actual_agent_tokens':None}
    if args.command!='status' and (run/'state.json').is_file() and not (args.command in ('draft','prepare-review') and existed_on_entry):
        dest=run/'batch_steps';dest.mkdir(exist_ok=True);n=len(list(dest.glob('flow_*.json')))+1
        (dest/f'flow_{n:03d}.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    summary={k:result.get(k) for k in ('status','receipt_integrity','required_steps_attempted','remaining','review_requests','preparation_failures') if k in result}
    summary.update(run=str(run),command=args.command,wall_seconds=receipt['wall_seconds'],model_api_calls=0,
                   deliverables=result.get('deliverables',[]),events=events)
    if result.get('message'):summary['message']=result['message']
    if planning is not None:summary['review_plan']=planning
    if next_run is not None:summary['next_run']=next_run
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    return 2 if result.get('status') in ('ERROR','NEEDS_REPAIR','FAIL') else 0

if __name__=='__main__':raise SystemExit(main())

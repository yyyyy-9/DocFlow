"""Subprocess adapter for the original three-form workflow, without new parsers.

Only current 02/03 workbooks are editable. Submissions are frozen before the
original engine validates them. Every successful batch yields another full draft.
"""
from pathlib import Path
import json
import math
import os
import re
import subprocess
import sys
from .runtime import BridgeError,sha

KIND='wps-forms'


def root(rt):
    path=getattr(rt,'forms',None)
    if path is None or not (path/'forms_tool.py').is_file():
        raise BridgeError('Three-form workflow missing; set --forms-root or WPS_FORMS_ROOT')
    info=json.loads((path/'BUILD_INFO.json').read_text('utf-8-sig'))
    if info.get('protocol')!='0.5.7' or info.get('forms_version')!='0.1.6':
        raise BridgeError('This adapter requires Forms 0.1.6 / kernel protocol 0.5.7')
    return path


def fingerprint(rt):
    path=root(rt);manifest=json.loads((path/'runtime_manifest.json').read_text('utf-8-sig'))
    values={}
    for name in sorted(set(manifest['files'])|{'forms_tool.py','BUILD_INFO.json'}):
        target=(path/name).resolve()
        if not target.is_relative_to(path) or not target.is_file():raise BridgeError('Invalid Forms runtime entry: '+name)
        values[name]=sha(target.read_bytes())
    return sha(json.dumps(values,sort_keys=True).encode())


def validate(task):
    if task['outputs']!=['docx']:raise BridgeError('wps-forms produces Word drafts only')
    options=task.get('forms_options',{})
    if not isinstance(options,dict) or set(options)-{'pages','doc_conversion','policy'}:raise BridgeError('Unknown forms_options')
    for key in ('pages','doc_conversion'):
        if key in options and options[key] not in ('off','wps'):raise BridgeError(key+' must be off or wps')
    if options.get('policy','人工核查') not in ('人工核查','Agent辅助核查'):
        raise BridgeError('Use explicit classify after the first draft; policy must be 人工核查 or Agent辅助核查')
    if 'review_policy' in task:raise BridgeError('wps-forms always delivers first; legacy review_policy does not apply')
    if 'title' in task and (not isinstance(task['title'],str) or not task['title'].strip() or len(task['title'])>200):
        raise BridgeError('Forms title must contain 1-200 characters')


def command(rt,home,operation,*args):
    from .batch import write
    core=root(rt);logs=home/'logs';logs.mkdir(parents=True,exist_ok=True)
    argv=[sys.executable,'-X','utf8','-B',str(core/'forms_tool.py'),operation,*map(str,args)]
    index=len(list(logs.glob('*.json')))+1
    env={**os.environ,'PYTHONUTF8':'1','PYTHONDONTWRITEBYTECODE':'1','PYTHONUNBUFFERED':'1'}
    # The native engine bounds WPS operations. Large local batches may exceed
    # five minutes, so there is no fixed whole-job timeout or whole-chain retry.
    process=subprocess.run(argv,cwd=core,env=env,capture_output=True,text=True,encoding='utf8')
    try:value=json.loads(process.stdout)
    except ValueError:
        write(logs/f'{index:03d}-forms-{operation}.json',dict(command=argv,returncode=process.returncode,stdout=process.stdout,stderr=process.stderr))
        raise BridgeError('Forms returned invalid JSON; see task logs')
    write(logs/f'{index:03d}-forms-{operation}.json',dict(command=argv,returncode=process.returncode,result=value,stderr=process.stderr))
    if value.get('status')=='ERROR':raise BridgeError('Forms '+operation+': '+value.get('error','unknown error'))
    if process.returncode and not value.get('latest',{}).get('deliverables'):
        raise BridgeError('Forms '+operation+' failed: '+str(value))
    return value


def latest(folder,task):
    from .batch import read
    path=(folder/task['forms_current']).resolve();home=(folder/'tasks'/task['id']).resolve()
    if not path.is_relative_to(home):raise BridgeError('Forms current path escaped task')
    state=read(path/'forms_state.json');core=(path/state['drafts'][-1]['run']).resolve()
    if not core.is_relative_to(path):raise BridgeError('Core draft path escaped Forms run')
    return path,state,core


def editable_workbooks(folder,task):
    """Derive the two allowed paths; never trust an arbitrary exception list."""
    relative=Path(task.get('forms_current','')).as_posix()
    prefix='tasks/'+task['id']+'/rounds/'
    if not relative.startswith(prefix) or not re.fullmatch(r'\d{3,}-(first|classify|feedback)',relative[len(prefix):]):
        raise BridgeError('Invalid current Forms round path')
    current=Path(relative).relative_to(Path('tasks')/task['id'])
    return [(current/name).as_posix() for name in ('02_扫描核查.xlsx','03_反馈提交.xlsx')]


def update(folder,task,path,value,operation,origin):
    """Validate/build the entire next round before changing committed state.

    A late read error must leave the old current round, editable declarations,
    deliverables and history together. The caller still records the failed
    attempt and generated files; explicit resubmission uses another directory.
    """
    from .batch import read,fingerprint
    path=path.resolve();home=(folder/'tasks'/task['id']).resolve()
    if not path.is_relative_to(home):raise BridgeError('Forms output escaped task')
    # Stage only round fields. Keep the caller's attempt/stage records intact,
    # since submit() finishes those records after this function returns.
    pending=dict(id=task['id'],forms_current=path.relative_to(folder).as_posix())
    native=value.get('latest') or {}
    current,state,core=latest(folder,pending);deliverables=[]
    for item in native.get('deliverables',[]):
        output=(core/item['file']).resolve()
        if not output.is_relative_to(core) or fingerprint(output)!=item['sha256']:raise BridgeError('Invalid Forms deliverable')
        deliverables.append(dict(path=output.relative_to(folder).as_posix(),sha256=item['sha256'],kind='docx',document=item['id']))
    pending.update(route=KIND,native_status=native,deliverables=deliverables,
        editable_forms=editable_workbooks(folder,pending),
        next_action='Read latest drafts; fill current 02/03 and forms-submit. Use forms-review-plan for article-ordered batches.')
    healthy=native.get('receipt_integrity')=='PASS' and native.get('required_steps_attempted') and bool(deliverables)
    pending['status']='COMPLETE' if healthy and not any(native.get(key) for key in ('remaining','failed_stages','preparation_failures')) and value.get('status')!='NEEDS_REPAIR' else 'PARTIAL'
    inv=read(current/'inventory.json');export=read(current/'feedback_export.json')
    pending['pending_count']=len(export['exports']);pending['source_gaps']=native.get('preparation_failures',[])
    pending['forms_rounds']=[*task.get('forms_rounds',[]),dict(run=pending['forms_current'],operation=operation,origin=origin,
        status=pending['status'],pending=pending['pending_count'],deliverables=deliverables)]
    pending['source_order']=[dict(id=row['id'],name=row['name'],relative=row['relative']) for row in inv['sources']]
    # All fallible reads and state construction above have completed. No task
    # field or nested history list was mutated before this single commit.
    task.update(pending)


def start(rt,folder,receipt,task):
    from .batch import stage,read,fingerprint
    home=folder/'tasks'/task['id'];options=task['spec'].get('forms_options',{})
    task['forms_identity']=read(root(rt)/'BUILD_INFO.json')
    form=home/'inputs/01_输入.xlsx';dest=home/'rounds/001-first'
    stage(folder,receipt,task,'forms.init',lambda:command(rt,home,'init','--out',form,'--root',task['source_root'],
        '--template',task['template'],'--run',dest,'--title',task['spec'].get('title','工程报告汇总'),
        '--policy',options.get('policy','人工核查'),'--pages',options.get('pages','off'),
        '--doc-conversion',options.get('doc_conversion','wps')))
    value=stage(folder,receipt,task,'forms.first-draft',lambda:command(rt,home,'scan',form))
    update(folder,task,dest,value,'scan','program_draft')
    # The original engine freezes the directory hierarchy. Bind its exact
    # binary originals, including failed DOC conversions, without rescanning.
    for item in read(dest/'inventory.json')['sources']:
        original=item['preprocessing_original'];frozen=(dest/original['file']).resolve()
        if not frozen.is_relative_to(dest) or fingerprint(frozen)!=original['sha256']:raise BridgeError('Invalid frozen Forms original')
        task['inputs'].append(dict(original=str(Path(task['source_root'])/original['relative']),frozen=str(frozen),sha256=original['sha256']))


def submit(rt,folder,task_id,operation,form,origin='user'):
    from .batch import load,integrity,inventory,save,status,stage,check_input_snapshots,stamp
    from project_commit import report_lock
    if operation not in ('classify','feedback') or origin not in ('user','model'):raise BridgeError('Invalid Forms submission')
    with report_lock(Path(folder)):
        folder,receipt=load(rt,folder);task=next((row for row in receipt['tasks'] if row['id']==task_id),None)
        if not task or task.get('route')!=KIND:raise BridgeError('Choose a wps-forms task with a first draft')
        if integrity(folder,task):raise BridgeError('Task changed outside its two editable workbooks')
        check_input_snapshots(task,originals=True)
        current,state,core=latest(folder,task);home=folder/'tasks'/task_id
        editable_before=editable_workbooks(folder,task);source=Path(form).absolute()
        from project_workspace import is_link
        if source.suffix.lower()!='.xlsx' or is_link(source):raise BridgeError('Submit a regular XLSX workbook')
        source=source.resolve()
        raw=source.read_bytes();before=sha(raw)
        attempt=len(task.get('forms_attempts',[]))+1;frozen=home/'submissions'/f'{attempt:03d}-{operation}.xlsx'
        frozen.parent.mkdir(exist_ok=True)
        dest=home/'rounds'/f'{attempt+1:03d}-{operation}'
        record=dict(operation=operation,origin=origin,input=frozen.relative_to(folder).as_posix(),sha256=before,
                    output=dest.relative_to(folder).as_posix(),started=stamp(),status='RUNNING')
        task.setdefault('forms_attempts',[]).append(record)
        try:
            with frozen.open('xb') as stream:stream.write(raw)
            if sha(source.read_bytes())!=before:raise BridgeError('Workbook changed while freezing submission')
            value=stage(folder,receipt,task,'forms.'+operation,lambda:command(rt,home,operation,current,'--form',frozen,'--out',dest,'--origin',origin))
            update(folder,task,dest,value,operation,origin);record['status']='APPLIED'
        except Exception as exc:
            record.update(status='REJECTED',error=str(exc));raise
        finally:
            changed=integrity(folder,task,allow_new=True,extra_editable=editable_before)
            if changed:task['status']='FAIL';task['error']='Protected task files changed: '+str(changed)
            else:task['files']=inventory(home)
            record['finished']=stamp();save(folder,receipt)
        return status(rt,folder)


def review_plan(rt,folder,task_id,fraction=.2):
    from .batch import load,integrity,read,fingerprint
    if not math.isfinite(fraction) or not 0<fraction<=1:raise BridgeError('fraction must be greater than 0 and at most 1')
    folder,receipt=load(rt,folder);task=next((row for row in receipt['tasks'] if row['id']==task_id),None)
    if not task or task.get('route')!=KIND:raise BridgeError('Choose a wps-forms task')
    if integrity(folder,task):raise BridgeError('Task protected files changed')
    current,state,core=latest(folder,task);meta=read(current/'feedback_export.json');inv=read(current/'inventory.json')
    approved=read(current/'approved_plan.json')
    owners={entry['id']:entry['source_id'] for entry in approved['entries']}
    target=max(1,math.ceil(len(meta['exports'])*fraction));batches=[];batch=[];count=0;used=set()
    for source in inv['sources']:
        items=[row for row in meta['exports'] if owners.get(row.get('source_id'),row.get('source_id'))==source['id']]
        if not items:continue
        used.update(row['id'] for row in items)
        batch.append(dict(source_id=source['id'],name=source['name'],relative=source['relative'],items=items));count+=len(items)
        if count>=target:batches.append(dict(articles=batch,pending=count));batch=[];count=0
    if batch:batches.append(dict(articles=batch,pending=count))
    other=[row for row in meta['exports'] if row['id'] not in used]
    if other:batches.append(dict(articles=[],pending=len(other),unassigned=other))
    return dict(ok=True,schema='hrs-bridge-forms-review-plan-v1',run=str(current),task=task_id,fraction=fraction,
        target_pending=target,pending=len(meta['exports']),batches=batches,
        bindings=dict(forms_state_sha256=fingerprint(current/'forms_state.json'),feedback_export_sha256=fingerprint(current/'feedback_export.json')),
        scope='Program-exported pending items; unsubmitted user-added workbook rows are not counted.',
        next_action='Read necessary article evidence, fill current 03, and forms-submit once per batch. Replan on the next draft; no model API is called.')

"""Source-bound task dispatch over the existing WPS and HRS entry points.

No document parser or assembler lives here. Tasks are independent; native Word
keeps the original WPS review/finish/status protocol and never passes through MD.
"""
from collections import Counter
from copy import deepcopy
from datetime import datetime,timezone
from pathlib import Path
import base64,json,os,re,shutil,subprocess,sys,time
from .runtime import BridgeError,sha,new_path

SCHEMA='hrs-bridge-job-v2'
KINDS={'docx','wps-job','wps-forms','markdown','hrs','html'}
OUTPUTS={'docx','hrs','html'}

def read(path):return json.loads(Path(path).read_text('utf-8-sig'))
def fingerprint(path):return sha(Path(path).read_bytes())
def stamp():return datetime.now(timezone.utc).isoformat()
def write(path,value):
    from studio import atomic_write
    atomic_write(Path(path),(json.dumps(value,ensure_ascii=False,indent=2)+'\n').encode('utf8'))

def runtime_fingerprint(rt,uses_forms=False):
    from wps_pipeline.job_support import fingerprints
    bridge=Path(__file__).resolve().parent
    hrs={p.relative_to(rt.hrs).as_posix():fingerprint(p) for p in sorted(rt.hrs.rglob('*'))
         if p.is_file() and p.suffix in ('.py','.js','.css') and '__pycache__' not in p.parts}
    result={'hrs':sha(json.dumps(hrs,sort_keys=True).encode()),
            'wps':sha(json.dumps(fingerprints(),sort_keys=True).encode()),
            'bridge':sha(json.dumps({p.name:fingerprint(p) for p in sorted(bridge.iterdir()) if p.is_file() and p.suffix in ('.py','.js','.cjs')},sort_keys=True).encode())}
    if uses_forms:
        from .forms import fingerprint as forms_fingerprint
        result['forms']=forms_fingerprint(rt)
    return result

def inventory(folder):
    return {p.relative_to(folder).as_posix():fingerprint(p) for p in sorted(Path(folder).rglob('*')) if p.is_file()}

def validate_manifest(job):
    if not isinstance(job,dict) or job.get('schema')!=SCHEMA:raise BridgeError('Unsupported task job schema')
    if not isinstance(job.get('output'),str) or not job['output'].strip():raise BridgeError('Job output must be a new directory path')
    tasks=job.get('tasks')
    if not isinstance(tasks,list) or not tasks:raise BridgeError('tasks must be a nonempty list')
    seen=set()
    for task in tasks:
        if not isinstance(task,dict):raise BridgeError('Each task must be an object')
        ident=task.get('id')
        if not isinstance(ident,str) or not re.fullmatch(r'[A-Za-z][A-Za-z0-9_-]{0,59}',ident) or ident.casefold() in seen:
            raise BridgeError('Task IDs must be unique simple names (case insensitive)')
        if ident.split('.')[0].casefold() in {'con','prn','aux','nul',*[f'com{i}' for i in range(1,10)],*[f'lpt{i}' for i in range(1,10)]}:
            raise BridgeError('Reserved task ID')
        seen.add(ident.casefold())
        source=task.get('source')
        if not isinstance(source,dict) or not isinstance(source.get('kind'),str) or source['kind'] not in KINDS or not isinstance(source.get('path'),str) or not source['path'].strip():
            raise BridgeError('Task source needs kind and path')
        outputs=task.get('outputs')
        if not isinstance(outputs,list) or not outputs or any(not isinstance(v,str) or v not in OUTPUTS for v in outputs) or len(set(outputs))!=len(outputs):
            raise BridgeError('Task outputs must select docx, hrs and/or html without duplicates')
        if source['kind']=='wps-job' and outputs!=['docx']:raise BridgeError('Native WPS jobs produce docx only')
        if source['kind']=='wps-forms':
            from .forms import validate as validate_forms
            validate_forms(task)
        elif 'forms_options' in task:raise BridgeError('forms_options applies only to wps-forms')
        if source['kind']=='docx' and 'docx' in outputs and len(outputs)>1:
            raise BridgeError('Choose a native docx task and a separate Word-to-HRS task; their structural scopes differ')
        if task.get('review_policy','pause') not in ('pause','retain-unresolved'):raise BridgeError('Unknown review_policy')
        if 'accept_numbered_headings' in task and type(task['accept_numbered_headings']) is not bool:raise BridgeError('accept_numbered_headings must be boolean')
        if 'word_options' in task:
            if source['kind'] in ('docx','wps-job','wps-forms') and outputs==['docx']:raise BridgeError('word_options applies to HRS/Markdown output; native Word uses its own contract')
            from .docx_export import output_options
            output_options(task['word_options'])
        for container,key in ((job,'template'),(task,'template'),(source,'provenance')):
            if key in container and (not isinstance(container[key],str) or not container[key].strip()):raise BridgeError(key+' must be a nonempty path string')
    return tasks

def aggregate(receipt):
    counts=Counter(t['status'] for t in receipt['tasks'])
    if counts['FAIL']:state='FAIL' if counts['FAIL']==len(receipt['tasks']) else 'PARTIAL'
    elif counts['PARTIAL']:state='PARTIAL'
    elif counts['NEEDS_REVIEW']:state='NEEDS_REVIEW'
    elif counts['COMPLETE']==len(receipt['tasks']):state='COMPLETE'
    else:state='RUNNING'
    receipt.update(status=state,ok=not counts['FAIL'],accepted=state=='COMPLETE',counts=dict(counts),updated=stamp())
    return receipt

def save(folder,receipt):
    aggregate(receipt);write(folder/'WORKFLOW.json',receipt)
    lines=['# 工作流交付','',f'批次状态：{receipt["status"]}','', '| 任务 | 路线 | 状态 | 交付 / 下一步 |','| --- | --- | --- | --- |']
    remaining=[]
    for task in receipt['tasks']:
        items=['['+Path(x['path']).name+']('+x['path'].replace(' ','%20')+')' for x in task.get('deliverables',[])]
        if task['status']!='COMPLETE':
            item={'task':task['id'],'status':task['status'],'reason':task.get('error') or task.get('next_action','需要检查本任务回执'),'review':task.get('review_requests')}
            remaining.append(item);items.append(item['reason'])
        lines.append('| '+task['id']+' | '+task.get('route','待执行')+' | '+task['status']+' | '+'；'.join(items).replace('|','\\|')+' |')
        if task.get('route')=='wps-forms':
            current=task.get('forms_current','')
            lines.extend(['',f'三表任务 {task["id"]}：已交 {len(task.get("forms_rounds",[]))} 轮，当前待审 {task.get("pending_count",0)} 项。',
                          f'当前02/03位于 `{current}`；按文章顺序填写一批，再显式 forms-submit 生成整份新稿。',''])
    lines+=['','原件保留。COMPLETE仅表示本次适用本地检查完成；原生WPS/Word页面验收单列。',
            '待判断任务：通过 job-review 提交来源绑定响应，再用 resume-job 接续；未完成项不因其他任务成功而消失。']
    (folder/'DELIVERY.md').write_text('\n'.join(lines)+'\n',encoding='utf8')
    write(folder/'finish_list.json',remaining)

def stage(folder,receipt,task,name,fn):
    event={'name':name,'started':stamp(),'status':'RUNNING'};task['stages'].append(event);save(folder,receipt);start=time.perf_counter()
    try:
        value=fn();event.update(status='PASS',result=value);return value
    except Exception as exc:
        event.update(status='FAIL',error=str(exc));raise
    finally:event['wall_seconds']=round(time.perf_counter()-start,6)

def copy_input(source,target,records):
    source=Path(source).resolve()
    from project_workspace import is_link
    if is_link(source) or not source.is_file():raise BridgeError('Missing or linked source file: '+str(source))
    data=source.read_bytes();digest=sha(data);target.parent.mkdir(parents=True,exist_ok=True)
    with target.open('xb') as stream:stream.write(data)
    if fingerprint(source)!=digest:raise BridgeError('Input changed while taking its snapshot: '+str(source))
    records.append({'original':str(source),'frozen':str(target),'sha256':digest});return target

def default_template(rt):
    config=read(rt.wps/'config.json');path=Path(config['template'])
    return path if path.is_absolute() else rt.wps/path

def prepare_inputs(rt,folder,task,base,job_template):
    spec=task['spec'];source=spec['source'];path=(base/source['path']).resolve();inputs=folder/'tasks'/task['id']/'inputs'
    inputs.mkdir(parents=True);records=[];kind=source['kind'];task['source_kind']=kind
    if path==folder or path.is_relative_to(folder):raise BridgeError('A task source cannot be inside its new output directory')
    if kind=='wps-forms':
        from .forms import root as forms_root
        from project_workspace import is_link
        if not path.is_dir() or is_link(path):raise BridgeError('wps-forms source must be an existing unlinked directory')
        core=forms_root(rt);config=read(core/'config.json');default=Path(config['template'])
        if not default.is_absolute():default=core/default
        template=(base/spec['template']).resolve() if spec.get('template') else (base/job_template).resolve() if job_template else default
        task['source_root']=str(path);task['template']=str(copy_input(template,inputs/'template.docx',records))
    elif kind=='wps-job':
        original=copy_input(path,inputs/'source-job.json',records);native=read(original)
        if not isinstance(native.get('sources'),list) or not native['sources']:raise BridgeError('WPS job needs sources')
        paths=set()
        for index,item in enumerate(native['sources']):
            p=(path.parent/item['path']).resolve()
            if str(p).casefold() in paths:raise BridgeError('Duplicate source path in WPS job')
            paths.add(str(p).casefold());item.setdefault('name',p.name)
            item['path']=str(copy_input(p,inputs/'documents'/f'{index:03d}.docx',records))
        template=(base/spec['template']).resolve() if spec.get('template') else (base/job_template).resolve() if job_template else (path.parent/native['template']).resolve() if native.get('template') else default_template(rt)
        native['template']=str(copy_input(template,inputs/'template.docx',records));task['native_job']=native
    elif kind=='hrs':
        if path.is_file():path=path.parent
        if not path.is_dir():raise BridgeError('HRS project directory missing')
        project=inputs/'project';copy_input(path/'report.md',project/'report.md',records)
        for name in ('meta.json','bridge-origin.json'):
            if (path/name).is_file():copy_input(path/name,project/name,records)
        assets=path/'assets'
        from project_workspace import is_link
        if is_link(assets):raise BridgeError('Linked asset directory is not supported')
        if assets.is_dir():
            for item in sorted(assets.rglob('*')):
                if is_link(item):raise BridgeError('Linked asset path is not supported: '+str(item))
                if item.is_file():copy_input(item,project/'assets'/item.relative_to(assets),records)
        (project/'assets').mkdir(exist_ok=True);task['project']=str(project)
    elif kind=='markdown':
        from markdown_import import prepare_markdown_import
        from studio import check_image
        frozen=copy_input(path,inputs/'markdown-source'/path.name,records)
        from object_clipboard import _fragment,local_urls,local_path
        ir=_fragment(frozen.read_text('utf-8-sig'));copied=set()
        for url in sorted(set(ir['assets'])|local_urls(ir)):
            image=local_path(path.parent.resolve(),url)
            if image is None:raise BridgeError('Markdown image source disappeared: '+url)
            if image not in copied:
                copy_input(image,frozen.parent/image.relative_to(path.parent),records);copied.add(image)
        # The existing adapter reads the frozen MD/dependencies, not a second
        # live snapshot that might change between resource enumeration and use.
        prepared=prepare_markdown_import(frozen,check_image)
        project=inputs/'project';(project/'assets').mkdir(parents=True)
        (project/'report.md').write_text(prepared['markdown'],encoding='utf8',newline='')
        for name,data in prepared['assets'].items():(project/'assets'/name).write_bytes(base64.b64decode(data,validate=True))
        write(project/'meta.json',{'render_profile':'free','render_options':prepared['render_options']})
        task['project']=str(project)
    else:
        task['source']=str(copy_input(path,inputs/('source.docx' if kind=='docx' else 'source.html'),records))
        if kind=='docx' and spec['outputs']!=['docx']:
            provenance=(base/source['provenance']).resolve() if source.get('provenance') else Path(str(path)+'.hrs-bridge.json')
            if source.get('provenance') or provenance.is_file():
                copy_input(provenance,Path(task['source']+'.hrs-bridge.json'),records)
    if kind not in ('wps-job','wps-forms') and 'docx' in spec['outputs']:
        template=(base/spec['template']).resolve() if spec.get('template') else (base/job_template).resolve() if job_template else default_template(rt)
        task['template']=str(copy_input(template,inputs/'template.docx',records))
    task['inputs']=records
    return {'files':len(records),'kind':kind}

def check_input_snapshots(task,originals=False):
    for entry in task['inputs']:
        if fingerprint(entry['frozen'])!=entry['sha256']:raise BridgeError('Frozen input changed: '+entry['frozen'])
        if originals and fingerprint(entry['original'])!=entry['sha256']:raise BridgeError('Source changed since its snapshot: '+entry['original'])

def wps_command(rt,task_folder,task,command,*arguments):
    args=[sys.executable,'-X','utf8','-B',str(rt.wps/'report_tool.py'),command,*map(str,arguments)]
    result=subprocess.run(args,cwd=task_folder,capture_output=True,text=True,encoding='utf8',timeout=300,
                          env={**os.environ,'PYTHONUTF8':'1','PYTHONDONTWRITEBYTECODE':'1'})
    log=task_folder/'logs';log.mkdir(exist_ok=True);index=len(list(log.glob('*.json')))+1
    try:value=json.loads(result.stdout)
    except ValueError:raise BridgeError('WPS returned invalid JSON: '+result.stderr[-2000:])
    write(log/f'{index:03d}-{command}.json',{'command':args,'returncode':result.returncode,'result':value,'stderr':result.stderr})
    retained_draft=(command in ('finish','status') and result.returncode==2
                    and value.get('status')=='NEEDS_REPAIR' and value.get('receipt_integrity')=='PASS'
                    and bool(value.get('deliverables')) and bool(value.get('required_steps_attempted')))
    if (result.returncode and not retained_draft) or value.get('receipt_integrity')=='FAIL':raise BridgeError('WPS '+command+': '+str(value))
    return value

def native_word(rt,folder,receipt,task):
    task['route']='wps-native';home=folder/'tasks'/task['id'];run=home/'wps'
    if not run.exists():
        native=task.get('native_job') or {'sources':[{'id':'source_001','path':task['source'],'name':Path(task['inputs'][0]['original']).name}],'make_master':False,'template':task['template']}
        job=home/'native-job.json';write(job,native)
        value=stage(folder,receipt,task,'word.start',lambda:wps_command(rt,home,task,'start',job,'--run',run))
    else:value=stage(folder,receipt,task,'word.status',lambda:wps_command(rt,home,task,'status',run))
    if value.get('review_requests') and task['spec'].get('review_policy','pause')=='pause' and not value['deliverables']:
        task.update(status='NEEDS_REVIEW',review_requests=str(run/'review_requests.json'),native_status=value,
                    next_action='Read source-bound review_requests.json; submit job-review, then resume-job.');return
    if not value['deliverables']:
        args=['--retain-unresolved'] if task['spec'].get('review_policy')=='retain-unresolved' else []
        stage(folder,receipt,task,'word.finish',lambda:wps_command(rt,home,task,'finish',run,*args))
    value=stage(folder,receipt,task,'word.status',lambda:wps_command(rt,home,task,'status',run))
    task['native_status']=value;task['deliverables']=[{'path':str((run/item['file']).relative_to(folder)),'sha256':item['sha256'],'kind':'docx'} for item in value['deliverables']]
    if not value.get('required_steps_attempted') or not value['deliverables']:raise BridgeError('WPS did not complete its required steps and deliverables')
    task['status']='PARTIAL' if value.get('remaining') or value.get('failed_stages') or value.get('preparation_failures') else 'COMPLETE'
    task['next_action']='Read native WPS DELIVERY.md and finish_list.json' if task['status']=='PARTIAL' else 'Read generated documents and native checks'

def report_route(rt,folder,receipt,task):
    from .browser import recover,export
    from .docx_import import convert as from_docx
    from .docx_export import convert as to_docx
    from .workflow import validate
    home=folder/'tasks'/task['id'];kind=task['source_kind'];spec=task['spec'];task['route']='hrs-bridge'
    if kind in ('hrs','markdown'):project=Path(task['project'])
    else:
        project=home/('project-reviewed-'+fingerprint(task['structure_map'])[:12] if task.get('structure_map') else 'project')
        if kind=='html':stage(folder,receipt,task,'html.recover',lambda:recover(rt,Path(task['source']),project))
        else:
            result=stage(folder,receipt,task,'word.recognize_and_import',lambda:from_docx(rt,Path(task['source']),project,structure_map=Path(task['structure_map']) if task.get('structure_map') else None))
            if result.get('recognition',{}).get('unresolved'):
                task.update(status='NEEDS_REVIEW',review_requests=str(project/'word-recognition.json'),
                            next_action='Review the source-bound Word structure. Submit a structure map through job-review, then resume-job.')
                task['project']=str(project);return
    stage(folder,receipt,task,'report.validate',lambda:validate(rt,project))
    deliverables=[]
    if 'docx' in spec['outputs']:
        path=home/'report.docx'
        value=stage(folder,receipt,task,'report.to_docx',lambda:to_docx(rt,project,path,template=Path(task['template']),accept_numbered_headings=spec.get('accept_numbered_headings',False),options=spec.get('word_options')))
        deliverables.extend([{'path':str(path.relative_to(folder)),'sha256':value['sha256'],'kind':'docx'},
                             {'path':str(Path(value['provenance']).relative_to(folder)),'sha256':value['provenance_sha256'],'kind':'provenance'}])
    if 'html' in spec['outputs']:
        path=home/'report.html';stage(folder,receipt,task,'report.export_html',lambda:export(rt,project,path,accept_numbered_headings=spec.get('accept_numbered_headings',False)))
        deliverables.append({'path':str(path.relative_to(folder)),'sha256':fingerprint(path),'kind':'html'})
    if 'hrs' in spec['outputs']:deliverables.append({'path':str((project/'report.md').relative_to(folder)),'sha256':fingerprint(project/'report.md'),'kind':'hrs'})
    task.update(status='COMPLETE',deliverables=deliverables,next_action='Read requested outputs; source project and assets remain with this task')

def execute(rt,folder,receipt,task,resuming=False):
    home=folder/'tasks'/task['id'];home.mkdir(parents=True,exist_ok=True)
    try:
        if resuming:
            check_input_snapshots(task,originals=True)
        else:
            stage(folder,receipt,task,'inputs.freeze',lambda:prepare_inputs(rt,folder,task,Path(receipt['base']),receipt.get('template')))
        if task['source_kind']=='wps-forms':
            from .forms import start
            start(rt,folder,receipt,task)
        elif task['source_kind'] in ('docx','wps-job') and task['spec']['outputs']==['docx']:native_word(rt,folder,receipt,task)
        else:report_route(rt,folder,receipt,task)
        check_input_snapshots(task,originals=True)
    except Exception as exc:task.update(status='FAIL',error=str(exc),accepted=False)
    finally:
        task['files']=inventory(home);task['finished']=stamp();save(folder,receipt)

def run(rt,job_file):
    job_file=Path(job_file).resolve();raw=job_file.read_bytes();job=json.loads(raw.decode('utf-8-sig'));tasks=validate_manifest(job)
    folder=new_path(job_file.parent/job['output']);folder.mkdir()
    receipt={'schema':SCHEMA,'job':str(job_file),'job_sha256':sha(raw),'base':str(job_file.parent),'output':str(folder),
             'template':job.get('template'),'started':stamp(),'runtime':runtime_fingerprint(rt,any(t['source']['kind']=='wps-forms' for t in tasks)),'conversion_model_calls':0,
             'tasks':[{'id':t['id'],'spec':deepcopy(t),'status':'PENDING','stages':[],'deliverables':[]} for t in tasks]}
    from project_commit import report_lock
    with report_lock(folder):
        (folder/'job.original.json').write_bytes(raw);receipt['frozen_job_sha256']=fingerprint(folder/'job.original.json');save(folder,receipt)
        for task in receipt['tasks']:execute(rt,folder,receipt,task)
    return status(rt,folder)

def load(rt,folder):
    folder=Path(folder).resolve();receipt=read(folder/'WORKFLOW.json')
    if receipt.get('schema')!=SCHEMA:raise BridgeError('Not a v2 task workflow')
    if fingerprint(folder/'job.original.json')!=receipt['frozen_job_sha256']:raise BridgeError('Frozen task manifest changed')
    if receipt['runtime']!=runtime_fingerprint(rt,any(t['spec']['source']['kind']=='wps-forms' for t in receipt['tasks'])):raise BridgeError('Workflow code changed; start a fresh revision')
    original=read(folder/'job.original.json');tasks=validate_manifest(original)
    if [(t.get('id'),t.get('spec')) for t in receipt['tasks']]!=[(t['id'],t) for t in tasks] or receipt.get('template')!=original.get('template'):
        raise BridgeError('Task specification changed outside the frozen manifest')
    return folder,receipt

def integrity(folder,task,allow_new=False,extra_editable=()):
    actual=inventory(folder/'tasks'/task['id']);expected=task.get('files',{})
    editable=set()
    if task.get('route')=='wps-forms':
        from .forms import editable_workbooks
        editable=set(editable_workbooks(folder,task))
        if task.get('editable_forms')!=sorted(editable):return ['Invalid editable workbook declaration']
        # Extra paths are the previous current round during a successful
        # submission, not caller-controlled files or filename globs.
        editable.update(extra_editable)
    names=expected.keys() if allow_new else actual.keys()|expected.keys()
    return [name for name in sorted(names) if actual.get(name)!=expected.get(name)
            and not (name in editable and name in actual and name in expected)]

def status(rt,folder):
    folder,receipt=load(rt,folder);errors=[]
    for task in receipt['tasks']:
        changed=integrity(folder,task)
        if changed:errors.append({'task':task['id'],'changed':changed})
    receipt=deepcopy(receipt);aggregate(receipt)
    receipt['integrity']='FAIL' if errors else 'PASS';receipt['integrity_errors']=errors
    if errors:receipt.update(status='FAIL',ok=False,accepted=False)
    return receipt

def resume(rt,folder):
    from project_commit import report_lock
    with report_lock(Path(folder)):
        return _resume(rt,folder)

def _resume(rt,folder):
    folder,receipt=load(rt,folder)
    for task in receipt['tasks']:
        if integrity(folder,task):raise BridgeError('Task files changed outside the workflow: '+task['id'])
    for task in receipt['tasks']:
        if task['status']=='NEEDS_REVIEW':
            if task.get('route')!='wps-native' and not task.get('structure_map'):continue
            execute(rt,folder,receipt,task,resuming=True)
    return status(rt,folder)

def review(rt,folder,task_id,response):
    from project_commit import report_lock
    with report_lock(Path(folder)):
        return _review(rt,folder,task_id,response)

def _review(rt,folder,task_id,response):
    folder,receipt=load(rt,folder);task=next((t for t in receipt['tasks'] if t['id']==task_id),None)
    if not task or task['status']!='NEEDS_REVIEW':raise BridgeError('Choose a task that needs review')
    if integrity(folder,task):raise BridgeError('Task files changed outside the workflow')
    check_input_snapshots(task,originals=True)
    home=folder/'tasks'/task['id'];response=Path(response).resolve()
    try:
        if task['route']=='wps-native':
            stage(folder,receipt,task,'word.review',lambda:wps_command(rt,home,task,'review',home/'wps',response))
        else:
            # External Word structural decisions are checked by the existing adapter
            # against its original source hash; no user-editable candidate is trusted.
            from .docx_import import NativeReader
            from .recognition import load_decisions
            import tempfile
            with tempfile.TemporaryDirectory(prefix='bridge-review-') as temp:
                reader=NativeReader(rt,Path(task['source']),Path(temp)/'scan',Path(temp)/'assets')
                load_decisions(reader,response)
            target=home/('structure-map-'+str(len(list(home.glob('structure-map-*.json')))+1)+'.json')
            copy_input(response,target,task['inputs']);task['structure_map']=str(target)
    finally:
        task['files']=inventory(home);save(folder,receipt)
    return status(rt,folder)

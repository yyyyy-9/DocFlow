"""Fixed single-report chains; paths are declared once in a source-bound job."""
from datetime import datetime, timezone
import json
from pathlib import Path
import time
from .browser import recover, export as export_html
from .docx_export import convert as to_docx
from .docx_import import convert as from_docx
from .runtime import BridgeError, new_path, sha


def project_hashes(project):
    project=Path(project)
    return {p.relative_to(project).as_posix():sha(p.read_bytes()) for p in [project/'report.md',*sorted((project/'assets').rglob('*'))] if p.is_file()}


def validate(rt, project):
    from .runtime import source_contract_valid
    project=Path(project);ir=rt.analyze((project/'report.md').read_bytes(),project/'assets')
    if not ir['ok'] or not source_contract_valid(ir):raise BridgeError('Recovered project fails HRS compilation checks')
    return {'ok':True,'blocks':len(ir['blocks']),'diagnostics':ir['diagnostics']}


def run_legacy(rt, job_file):
    job_file=Path(job_file).resolve();job=json.loads(job_file.read_text('utf-8-sig'))
    if job.get('schema')!='hrs-bridge-job-v1':raise BridgeError('Unsupported bridge job schema')
    base=job_file.parent
    def path(value):return (base/value).resolve()
    if not isinstance(job.get('inputs'),dict) or not job['inputs'] or set(job['inputs'])-{'html','docx'}:
        raise BridgeError('Job inputs must contain html and/or docx')
    sources={key:path(value) for key,value in job['inputs'].items()}
    plan=path(job['structure_map']) if job.get('structure_map') else None
    template=path(job['template']) if job.get('template') else None
    accept_numbers=job.get('accept_numbered_headings') is True
    tracked={'job':job_file,**sources,**({'structure_map':plan} if plan else {}),**({'template':template} if template else {})}
    before={key:sha(value.read_bytes()) for key,value in tracked.items()}
    output=new_path(path(job['output']));output.mkdir()
    receipt={'schema':'hrs-bridge-workflow-v1','started':datetime.now(timezone.utc).isoformat(),
             'job':str(job_file),'inputs':{key:{'path':str(value),'sha256':before[key]} for key,value in tracked.items()},
             'dependencies':{'hrs':str(rt.hrs),'wps':str(rt.wps)},'stages':[],
             'conversion_model_calls':0,'status':'RUNNING','accepted':False,'ok':False}
    started=time.perf_counter()
    def stage(name, function, *args, **kwargs):
        event={'name':name,'status':'RUNNING'};receipt['stages'].append(event);begin=time.perf_counter()
        try:
            result=function(*args,**kwargs)
            if result.get('ok') is False:raise BridgeError(str(result.get('error',result)))
            event.update(status='PASS',result=result);return result
        except Exception as exc:
            event.update(status='FAIL',error=str(exc));raise
        finally:event['wall_seconds']=round(time.perf_counter()-begin,6)
    def same_source(a,b):
        a,b=project_hashes(a),project_hashes(b)
        if a!=b:raise BridgeError('Word round trip changed Markdown or assets')
        return {'ok':True,'files':len(a),'markdown_sha256':a['report.md']}
    unresolved=0
    try:
        if 'html' in sources:
            source=output/'html_source';docx=output/'HTML归档转Word.docx'
            stage('html.recover',recover,rt,sources['html'],source)
            stage('html.to_docx',to_docx,rt,source,docx,template=template,accept_numbered_headings=accept_numbers)
            back=output/'html_roundtrip'
            stage('html.read_back',from_docx,rt,docx,back)
            stage('html.compare_source',same_source,source,back)
        if 'docx' in sources:
            project=output/'word_to_md'
            imported=stage('word.recognize_and_import',from_docx,rt,sources['docx'],project,structure_map=plan)
            unresolved=imported.get('recognition',{}).get('unresolved',0)
            stage('word.validate',validate,rt,project)
            stage('word.export_html',export_html,rt,project,output/'Word转Markdown_预览.html',accept_numbered_headings=accept_numbers)
        receipt.update(ok=True,status='NEEDS_REVIEW' if unresolved else 'COMPLETE',accepted=not unresolved)
    except Exception as exc:receipt.update(status='FAIL',error=str(exc))
    finally:
        try:
            receipt['source_unchanged']=all(sha(p.read_bytes())==before[key] for key,p in tracked.items())
            if not receipt['source_unchanged']:receipt.update(ok=False,status='FAIL',accepted=False,error='Input changed during workflow')
        except OSError as exc:receipt.update(ok=False,status='FAIL',accepted=False,error=str(exc))
        receipt['files']={p.relative_to(output).as_posix():sha(p.read_bytes()) for p in sorted(output.rglob('*')) if p.is_file()}
        receipt['wall_seconds']=round(time.perf_counter()-started,6)
        receipt['finished']=datetime.now(timezone.utc).isoformat()
        (output/'WORKFLOW.json').write_text(json.dumps(receipt,ensure_ascii=False,indent=2),encoding='utf8')
    return receipt


def run(rt, job_file):
    job=json.loads(Path(job_file).read_text('utf-8-sig'))
    if not isinstance(job,dict):raise BridgeError('Job manifest must be an object')
    if job.get('schema')=='hrs-bridge-job-v2':
        from .batch import run as batch_run
        return batch_run(rt,job_file)
    return run_legacy(rt,job_file)

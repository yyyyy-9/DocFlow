"""Single entry point for the portable document workflow (Python 3.10+)."""
from __future__ import annotations
import argparse
import ast
import hashlib
import importlib
import json
import sys
from pathlib import Path

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'src'))


def emit(value):
    print(json.dumps(value, ensure_ascii=False, indent=2))


def doctor():
    """Read-only installation, dependency, source and template diagnostics."""
    checks=[]
    product_identity={}
    def check(name, callback):
        try:
            detail=callback()
            checks.append({'check':name,'status':'PASS','detail':detail})
        except Exception as exc:
            checks.append({'check':name,'status':'FAIL','error':str(exc)})
    def version():
        if sys.version_info < (3,10):raise ValueError('Python 3.10 or later is required')
        return {'version':sys.version.split()[0],'executable':sys.executable}
    def dependency():
        module=importlib.import_module('lxml.etree')
        return '.'.join(map(str,module.LXML_VERSION))
    def code_and_manifest():
        from wps_pipeline.job_support import fingerprints
        actual=fingerprints()
        listed=set(actual)
        available={p.relative_to(ROOT).as_posix() for p in (ROOT/'src/wps_pipeline').glob('*.py')}
        if not available <= listed:raise ValueError('Unlisted runtime modules: '+', '.join(sorted(available-listed)))
        for name in sorted(available):
            path=ROOT/name
            ast.parse(path.read_text(encoding='utf-8-sig'),filename=name)
            importlib.import_module('wps_pipeline.'+path.stem) if path.stem!='__init__' else importlib.import_module('wps_pipeline')
        return {'fingerprinted_files':len(actual),'imported_modules':len(available)}
    def template():
        from wps_pipeline.resolve_styles import inspect_template
        config=json.loads((ROOT/'config.json').read_text(encoding='utf-8-sig'))
        value=Path(config['template'])
        if not value.is_absolute():value=ROOT/value
        report=inspect_template(value)
        return {'template_version':report['template_version'],'style_count':len(report['styles']),
                'sha256':hashlib.sha256(value.read_bytes()).hexdigest()}
    def release():
        data=json.loads((ROOT/'release_manifest.json').read_text(encoding='utf-8-sig'))
        changed=[]
        for item in data['files']:
            path=(ROOT/item['path']).resolve()
            if not path.is_relative_to(ROOT) or not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest()!=item['sha256']:
                changed.append(item['path'])
        if changed:raise ValueError('Missing or changed release files: '+', '.join(changed))
        from wps_pipeline.maintenance import identity
        actual=identity()
        if data.get('forms_version')!=actual['forms_version'] or data.get('build_id')!=actual['build_id'] or data.get('version')!=actual['kernel_protocol']:
            raise ValueError('发行清单与产品身份不同步')
        return {'forms_version':actual['forms_version'],'kernel_protocol':data['version'],
                'build_id':actual['build_id'],'verified_files':len(data['files'])}
    def product():
        from wps_pipeline.maintenance import identity
        product_identity.update(identity())
        if product_identity['identity_status']!='CONSISTENT':raise ValueError('; '.join(product_identity['identity_errors']))
        return dict(product_identity)
    check('product_identity',product)
    check('python',version);check('lxml',dependency)
    check('runtime_imports_and_manifest',code_and_manifest)
    check('template_contract',template);check('release_files',release)
    return {**product_identity,'status':'READY' if all(c['status']=='PASS' for c in checks) else 'ENVIRONMENT_OR_PACKAGE_ERROR',
            'checks':checks,'document_processing':'NOT_RUN','wps':'NOT_RUN','visual':'NOT_RUN',
            'note':'Installation diagnostics do not establish document recognition accuracy or WPS layout.'}


def make_parser():
    p=argparse.ArgumentParser(description='将用户选择的DOCX整理到模板；来源保持不变。')
    p.add_argument('--version',action='version',version='0.5.7')
    sub=p.add_subparsers(dest='command',required=True)
    sub.add_parser('doctor',help='检查解释器、依赖、运行文件和模板；不启动WPS')
    template=sub.add_parser('template',help='检查外部模板或在新副本中准备正文替换槽位')
    commands=template.add_subparsers(dest='template_command',required=True)
    t=commands.add_parser('inspect');t.add_argument('source',type=Path)
    t=commands.add_parser('styles');t.add_argument('source',type=Path)
    t=commands.add_parser('prepare');t.add_argument('source',type=Path)
    t.add_argument('--output',type=Path,required=True)
    t.add_argument('--body-start',type=int,required=True)
    t.add_argument('--expected-sha256',required=True)
    t.add_argument('--style-source',type=Path,help='Compatible DOCX supplying the named report styles; default is the current mother template')
    t.add_argument('--expected-style-sha256',help='Bind the selected style template to its inspected bytes')
    t.add_argument('--trim-cover-gap',action='store_true',help='仅移除封面首个人工分页符前紧邻的空段')
    s=sub.add_parser('start',help='从任务清单或明确选择的来源开始新任务')
    s.add_argument('job',type=Path,nargs='?')
    s.add_argument('--input',type=Path,action='append',default=[])
    s.add_argument('--input-dir',type=Path)
    s.add_argument('--run',type=Path,required=True)
    s.add_argument('--master',action='store_true',help='另外生成总稿')
    s.add_argument('--title',help='总稿标题')
    s.add_argument('--template',type=Path)
    s=sub.add_parser('review',help='提交模型或用户的局部结构决定')
    s.add_argument('run',type=Path);s.add_argument('response',type=Path)
    s=sub.add_parser('finish',help='执行装配、公式处理和全部适用本地检查')
    s.add_argument('run',type=Path);s.add_argument('--retain-unresolved',action='store_true')
    s=sub.add_parser('status',help='核对当前阶段、回执与成稿指纹')
    s.add_argument('run',type=Path)
    s=sub.add_parser('revise',help='来源核对后另建修订；保留旧版')
    s.add_argument('parent',type=Path);s.add_argument('--run',type=Path,required=True)
    s=sub.add_parser('feedback',help='定位问题或在新修订中应用明确反馈')
    fsub=s.add_subparsers(dest='feedback_command',required=True)
    f=fsub.add_parser('locate');f.add_argument('run',type=Path);f.add_argument('--document',required=True)
    f.add_argument('--quote');f.add_argument('--locator');f.add_argument('--before')
    f.add_argument('--problem',required=True);f.add_argument('--wanted',required=True)
    f.add_argument('--origin',required=True,choices=['user','model']);f.add_argument('--output',required=True,type=Path)
    f=fsub.add_parser('apply');f.add_argument('feedback',type=Path);f.add_argument('--run',required=True,type=Path)
    return p


def selected_job(args):
    if args.job:
        if args.input or args.input_dir:raise ValueError('Choose a job file OR direct inputs')
        path=args.job.resolve()
        job=json.loads(path.read_text(encoding='utf-8-sig'));base=path.parent
        if args.master:job['make_master']=True
    else:
        paths=[p.resolve() for p in args.input]
        if args.input_dir:
            directory=args.input_dir.resolve()
            if not directory.is_dir():raise ValueError('Input directory does not exist')
            if directory==ROOT or directory.is_relative_to(ROOT/'templates') or directory.is_relative_to(ROOT/'src'):
                raise ValueError('Choose a separate source directory, not the implementation or template directory')
            for parent in [directory,*directory.parents]:
                if (parent/'state.json').is_file() and (parent/'inputs').is_dir():
                    raise ValueError('Directory is inside an existing task; use feedback or explicitly select a file')
            paths.extend(sorted((p.resolve() for p in directory.iterdir()
                                 if p.is_file() and p.suffix.lower()=='.docx' and not p.name.startswith('~$')),
                                key=lambda p:p.name.casefold()))
        if not paths:raise ValueError('Select --input, --input-dir, or a job.json with sources')
        if len(set(paths))!=len(paths):raise ValueError('Duplicate source path selected')
        job={'name':'文档整理','make_master':args.master,
             'sources':[{'id':f'source_{i:03d}','path':str(p)} for i,p in enumerate(paths,1)]}
        base=Path.cwd()
    if args.title is not None:job['title']=args.title
    if args.template is not None:job['template']=str(args.template.resolve())
    return job,base


def execute(args):
    if args.command=='doctor':return doctor()
    if args.command=='template':
        from wps_pipeline.template_prepare import inspect_document, inspect_styles, prepare_template
        if args.template_command=='inspect':return inspect_document(args.source)
        if args.template_command=='styles':return inspect_styles(args.source)
        return prepare_template(args.source,args.output,args.body_start,args.expected_sha256,args.trim_cover_gap,
                                style_source=args.style_source,expected_style_sha256=args.expected_style_sha256)
    from wps_pipeline import workflow
    from wps_pipeline.job_support import read, write, check_inputs, sha, stamp, event
    if args.command=='start':
        job,base=selected_job(args)
        state=workflow.start_job(job,base,args.run)
        return workflow.status(args.run.resolve(),state)
    if args.command=='revise':
        state=workflow.revise(args.parent,args.run)
        return workflow.status(args.run.resolve(),state)
    if args.command=='feedback':
        from wps_pipeline import feedback
        if args.feedback_command=='apply':
            state=feedback.apply(args.feedback.resolve(),args.run.resolve())
            return workflow.status(args.run.resolve(),state)
        if args.output.exists():raise FileExistsError('Feedback output already exists; choose a new path')
        result=feedback.locate(args.run.resolve(),args.document,quote=args.quote,locator=args.locator,
                               problem=args.problem,wanted=args.wanted,before=args.before,origin=args.origin)
        args.output.parent.mkdir(parents=True,exist_ok=True)
        with args.output.open('x',encoding='utf-8') as out:json.dump(result,out,ensure_ascii=False,indent=2)
        return dict(result,feedback_file=str(args.output.resolve()))
    run=args.run.resolve();state=read(run/'state.json')
    if args.command=='status':return workflow.status(run,state)
    check_inputs(run,state)
    if args.command=='finish':
        return workflow.status(run,workflow.finish(run,state,retain=args.retain_unresolved))
    if args.command=='review':
        from wps_pipeline.model_review import accept
        if workflow.status(run,state).get('receipt_integrity')!='PASS':
            raise ValueError('Existing task receipts have changed; create or repair a revision before review')
        if state['deliverables']:raise ValueError('Create a fresh revision for completed output changes')
        if state['review_rounds']>=state['max_review_rounds']:
            raise ValueError('Configured review rounds reached; retain unresolved items or create a revision')
        response_dest=run/f'response_{state["review_rounds"]+1:03d}.json'
        if response_dest.exists():raise FileExistsError('Review response path already exists')
        result=accept(run,state,args.response.resolve())
        state['review_rounds']+=1
        response=result['accepted_response']
        write(response_dest,response)
        submitted={item['source_id'] for item in response['sources']}
        for src in state['sources']:
            if src['id'] not in submitted:continue
            path=run/'sources'/src['id']/'candidate.json'
            if 'recognize:'+src['id'] in state['stages'] and path.exists():
                candidate=read(path)
                if 'blocks' in candidate:
                    state['stages']['recognize:'+src['id']]['evidence_sha256']=sha(path)
                    state['stages']['recognize:'+src['id']]['status']='REVIEW_APPLIED'
        event(run,'local_review','APPLIED',response_sha256=result['response_sha256'],
              stored_response_sha256=sha(response_dest),round=state['review_rounds'])
        workflow.refresh_preparation_status(run,state)
        workflow.save(run,state)
        return workflow.status(run,state)
    raise ValueError('Unknown command')


def main():
    if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
    args=make_parser().parse_args()
    try:
        result=execute(args);emit(result)
        bad={'NEEDS_REPAIR','ENVIRONMENT_OR_PACKAGE_ERROR','FAIL','NOT_FOUND','AMBIGUOUS'}
        return 2 if result.get('status') in bad or result.get('receipt_integrity')=='FAIL' else 0
    except Exception as exc:
        emit({'status':'ERROR','exception':type(exc).__name__,'message':str(exc)})
        return 2


if __name__=='__main__':
    raise SystemExit(main())

"""Local JSON CLI. Outputs are new; only explicit block-apply commits source edits."""
import argparse
import json
import os
from pathlib import Path
import sys
from .runtime import Runtime,BridgeError,new_path,ROOT,default_roots
from . import __version__


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--version',action='version',version=__version__)
    parser.add_argument('--hrs-root',type=Path,default=os.environ.get('HRS_ROOT'))
    parser.add_argument('--wps-root',type=Path,default=os.environ.get('WPS_ROOT'))
    parser.add_argument('--forms-root',type=Path,default=os.environ.get('WPS_FORMS_ROOT'))
    parser.add_argument('--browser');parser.add_argument('--node');parser.add_argument('--pandoc')
    sub=parser.add_subparsers(dest='command',required=True)
    job=sub.add_parser('run-job');job.add_argument('job',type=Path)
    for command in ('job-status','resume-job'):
        p=sub.add_parser(command);p.add_argument('run',type=Path)
    p=sub.add_parser('job-review');p.add_argument('run',type=Path);p.add_argument('--task',required=True);p.add_argument('--response',type=Path,required=True)
    p=sub.add_parser('forms-submit');p.add_argument('run',type=Path);p.add_argument('--task',required=True)
    p.add_argument('--operation',choices=['classify','feedback'],required=True);p.add_argument('--form',type=Path,required=True)
    p.add_argument('--origin',choices=['user','model'],default='user')
    p=sub.add_parser('forms-review-plan');p.add_argument('run',type=Path);p.add_argument('--task',required=True);p.add_argument('--fraction',type=float,default=.2)
    analyze=sub.add_parser('analyze-word');analyze.add_argument('source',type=Path)
    analyze.add_argument('--structure-map',type=Path)
    recover=sub.add_parser('recover-html');recover.add_argument('source',type=Path);recover.add_argument('--out',required=True,type=Path)
    export=sub.add_parser('export-html');export.add_argument('source',type=Path);export.add_argument('--out',required=True,type=Path)
    export.add_argument('--quality',choices=['original','standard','compact','print'],default='original')
    export.add_argument('--mode',choices=['recompilable','reading'],default='recompilable')
    export.add_argument('--accept-numbered-headings',action='store_true')
    read=sub.add_parser('read');read.add_argument('source',type=Path)
    read.add_argument('--action',choices=['summary','objects','get','outline','get-block','references','diagnostics'],default='objects')
    read.add_argument('--id');read.add_argument('--index',type=int);read.add_argument('--expect-sha256')
    edit=sub.add_parser('block-plan');edit.add_argument('source',type=Path)
    target=edit.add_mutually_exclusive_group(required=True);target.add_argument('--id');target.add_argument('--index',type=int)
    edit.add_argument('--expect-sha256',required=True);edit.add_argument('--text-file',type=Path,required=True)
    edit.add_argument('--operation',choices=['replace','insert-before','insert-after'],default='replace')
    edit.add_argument('--out',type=Path,required=True)
    edit.add_argument('--plan-out',type=Path,help='Write the reviewed HRS plan next to the candidate')
    apply=sub.add_parser('block-apply');apply.add_argument('plan',type=Path);apply.add_argument('--expect-plan-sha256',required=True)
    for name in ['to-docx','from-docx']:
        p=sub.add_parser(name);p.add_argument('source',type=Path);p.add_argument('--out',type=Path,required=True)
        if name=='to-docx':
            p.add_argument('--template',type=Path,help='Prepared WPS template; default uses the WPS config')
            p.add_argument('--accept-numbered-headings',action='store_true')
            p.add_argument('--notes-mode',choices=['footnotes','references'],default='footnotes')
            p.add_argument('--strip-body-emphasis',action='store_true',help='Omit source bold/italic; preserve template and mathematical styles')
        if name=='from-docx':
            p.add_argument('--provenance',type=Path)
            p.add_argument('--structure-map',type=Path,help='Source-bound, explicit external Word group decisions')
    args=parser.parse_args(argv)
    try:
        if args.hrs_root is None or args.wps_root is None:
            hrs_default,wps_default=default_roots()
            args.hrs_root=args.hrs_root or hrs_default
            args.wps_root=args.wps_root or wps_default
        rt=Runtime(args.hrs_root,args.wps_root,args.browser,args.node,args.pandoc,args.forms_root)
        if args.command=='run-job':
            from .workflow import run
            result=run(rt,args.job)
        elif args.command=='forms-submit':
            from .forms import submit
            result=submit(rt,args.run,args.task,args.operation,args.form,args.origin)
        elif args.command=='forms-review-plan':
            from .forms import review_plan
            result=review_plan(rt,args.run,args.task,args.fraction)
        elif args.command in ('job-status','resume-job','job-review'):
            from .batch import status,resume,review
            result=status(rt,args.run) if args.command=='job-status' else resume(rt,args.run) if args.command=='resume-job' else review(rt,args.run,args.task,args.response)
        elif args.command=='analyze-word':
            import tempfile
            from .docx_import import NativeReader
            from .recognition import analyze,load_decisions
            with tempfile.TemporaryDirectory(prefix='bridge-analysis-') as folder:
                reader=NativeReader(rt,args.source,Path(folder),Path(folder)/'assets')
                result=analyze(reader,load_decisions(reader,args.structure_map))
                result['ok']=True
        elif args.command=='recover-html':
            from .browser import recover
            result=recover(rt,args.source,args.out)
        elif args.command=='export-html':
            from .browser import export
            result=export(rt,args.source,args.out,args.quality,args.mode,accept_numbered_headings=args.accept_numbered_headings)
        elif args.command=='read':
            from .blocks import read
            result=read(rt,args.source,args.action,args.id,args.index,args.expect_sha256)
        elif args.command=='block-plan':
            from .blocks import plan
            result,candidate=plan(rt,args.source,args.expect_sha256,args.text_file.read_bytes().decode('utf8'),
                                  object_id=args.id,index=args.index,operation=args.operation)
            if args.plan_out:
                import block_cli
                result=block_cli.write_plan(args.plan_out,result,candidate,candidate_path=args.out)
            else:
                output=new_path(args.out);output.write_bytes(candidate);result['candidate']=str(output)
        elif args.command=='block-apply':
            from .blocks import apply
            result=apply(rt,args.plan,args.expect_plan_sha256)
        elif args.command=='to-docx':
            from .docx_export import convert
            result=convert(rt,args.source,args.out,template=args.template,accept_numbered_headings=args.accept_numbered_headings,options={'notes_mode':args.notes_mode,'preserve_emphasis':not args.strip_body_emphasis})
        else:
            from .docx_import import convert
            result=convert(rt,args.source,args.out,provenance=args.provenance,structure_map=args.structure_map)
        print(json.dumps(result,ensure_ascii=False,indent=2))
        if result.get('status') in ('NEEDS_REVIEW','PARTIAL') and result.get('ok'):return 2
        return 0 if result.get('ok') else 1
    except (BridgeError,ValueError,OSError) as exc:
        print(json.dumps({'ok':False,'error':str(exc)},ensure_ascii=False));return 1


if __name__=='__main__':raise SystemExit(main())

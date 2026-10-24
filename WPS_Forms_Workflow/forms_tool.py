"""Offline forms entry; an external Agent may invoke these same commands."""
import argparse,json,sys
from contextlib import nullcontext
from pathlib import Path
sys.dont_write_bytecode=True
sys.path.insert(0,str(Path(__file__).resolve().parent/'src'))
from wps_pipeline import forms_workflow as f


def main():
    p=argparse.ArgumentParser(description='文件树 → 核查表 → 总稿 → 批量反馈 → 新稿（无模型 API）')
    sub=p.add_subparsers(dest='command',required=True)
    a=sub.add_parser('init');a.add_argument('--out',required=True);a.add_argument('--root',default='');a.add_argument('--template',default='');a.add_argument('--run',default='');a.add_argument('--title',default='工程报告汇总');a.add_argument('--policy',choices=f.POLICIES,default='人工核查');a.add_argument('--pages',choices=['off','wps'],default='off');a.add_argument('--doc-conversion',choices=['off','wps'],default='wps')
    a=sub.add_parser('scan');a.add_argument('input');a.add_argument('--run');a.add_argument('--pages',choices=['off','wps']);a.add_argument('--doc-conversion',choices=['off','wps'])
    a=sub.add_parser('assemble');a.add_argument('run');a.add_argument('--form');a.add_argument('--origin',choices=['user','model'],default='user')
    a=sub.add_parser('classify');a.add_argument('run');a.add_argument('--form');a.add_argument('--out',required=True);a.add_argument('--origin',choices=['user','model'],default='user')
    a=sub.add_parser('feedback');a.add_argument('run');a.add_argument('--form');a.add_argument('--out',required=True);a.add_argument('--origin',choices=['user','model'],default='user')
    a=sub.add_parser('status');a.add_argument('run')
    a=sub.add_parser('doctor',help='同时显示三表版本、内核协议、构建号及安装诊断')
    a=sub.add_parser('cleanup',help='保留阶段总稿并清理中间数据；默认预览')
    a.add_argument('run',nargs='?',default=str(Path(__file__).resolve().parent/'outputs'));a.add_argument('--apply',action='store_true');a.add_argument('--keep-latest',action='store_true')
    args=p.parse_args()
    try:
        if args.command=='doctor':
            from report_tool import doctor
            r=doctor();print(json.dumps(r,ensure_ascii=False,indent=2));return 0 if r['status']=='READY' else 2
        if args.command=='cleanup':
            from wps_pipeline.maintenance import cleanup_tasks
            default=Path(__file__).resolve().parent/'outputs'
            r=(dict(status='NOTHING_TO_CLEAN',root=str(default),cleaned_tasks=0,
                    note='包内尚无outputs目录；可传入实际任务的父目录。')
               if Path(args.run)==default and not default.exists()
               else cleanup_tasks(args.run,args.apply,args.keep_latest))
            print(json.dumps(r,ensure_ascii=False,indent=2));return 2 if r['status']=='PARTIAL' else 0
        from wps_pipeline.maintenance import task_lease
        lease=task_lease(args.run) if args.command in ('assemble','classify','feedback') else nullcontext()
        with lease:return execute(args)
    except Exception as exc:
        print(json.dumps(dict(status='ERROR',error=str(exc),exception=type(exc).__name__),ensure_ascii=False,indent=2));return 2


def execute(args):
    try:
        if args.command=='init':r={'input_form':f.new_input(args.out,args.root,args.template,args.run,args.title,args.policy,args.pages,args.doc_conversion)}
        elif args.command=='scan':r=f.scan_input(args.input,args.run,args.pages,args.doc_conversion)
        elif args.command=='assemble':r=f.assemble(args.run,args.form,args.origin)
        elif args.command=='classify':r=f.classify(args.run,args.form,args.out,args.origin)
        elif args.command=='feedback':r=f.submit_feedback(args.run,args.form,args.out,args.origin)
        else:r=f.forms_status(args.run)
        print(json.dumps(r,ensure_ascii=False,indent=2))
        # A delivered, explicitly flagged draft is still a successful handover.
        # Its check/gap status remains in the JSON and DELIVERY, never hidden.
        delivered=(r.get('latest') or {}).get('master',{}).get('status','').startswith('PRODUCED')
        return 0 if delivered or r.get('status')!='NEEDS_REPAIR' else 2
    except Exception as exc:
        print(json.dumps(dict(status='ERROR',error=str(exc),exception=type(exc).__name__),ensure_ascii=False,indent=2));return 2


if __name__=='__main__':raise SystemExit(main())

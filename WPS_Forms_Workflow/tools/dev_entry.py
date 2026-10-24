"""Thin local dispatcher. All conversion semantics remain in report_tool.py."""
from pathlib import Path
import hashlib,json,os,runpy,sys
ROOT=Path(__file__).resolve().parents[1]
sys.dont_write_bytecode=True
if hasattr(os,'add_dll_directory') and (ROOT/'environment/deps/pywin32_system32').is_dir():
 DLL_HANDLE=os.add_dll_directory(str(ROOT/'environment/deps/pywin32_system32'))

def main():
 args=sys.argv[1:]
 if not args:
  print('run.cmd verify | doctor | forms ... | cleanup [TASK_PARENT] [--apply] [--keep-latest] | wps ... | template inspect/styles/prepare ...');return 2
 mode,*args=args
 if mode=='verify':
  path=ROOT/'MANIFEST.json'
  if not path.is_file():raise SystemExit('Source worktree has no package MANIFEST; use report_tool.py doctor.')
  manifest=json.loads(path.read_text('utf8'));bad=[]
  for item in manifest['files']:
   p=ROOT/item['path']
   if not p.is_file() or hashlib.sha256(p.read_bytes()).hexdigest()!=item['sha256']:bad.append(item['path'])
  print(json.dumps({'status':'FAIL' if bad else 'PASS','files':len(manifest['files']),'mismatches':bad},ensure_ascii=False));return int(bool(bad))
 if mode=='forms':
  script=ROOT/'forms_tool.py'
 elif mode=='cleanup':
  script=ROOT/'cleanup_tool.py'
 elif mode=='batch':
  script=ROOT/'batch_tool.py'
 elif mode in ('doctor','wps','template'):
  script=ROOT/'report_tool.py';args=(['doctor'] if mode=='doctor' else ['template',*args] if mode=='template' else args)
 elif mode in ('trial','wps-observe','trial-native','wps-edit'):
  script=ROOT/'tests'/dict(trial='run_word_to_word.py',**{'wps-observe':'check_wps_labels.py','trial-native':'run_native_numbering.py','wps-edit':'check_wps_editing.py'})[mode]
  if not script.is_file():raise SystemExit('Tests are in the Developer package.')
 else:raise SystemExit('Unknown command '+mode)
 sys.path.insert(0,str(script.parent));sys.argv=[str(script),*args];runpy.run_path(str(script),run_name='__main__');return 0

if __name__=='__main__':raise SystemExit(main())

"""Run declared public suites and retain per-module results in a chosen output."""
from pathlib import Path
import argparse,hashlib,json,os,subprocess,sys,tempfile,time,traceback,unittest
ROOT=Path(__file__).resolve().parents[1]

def main():
    p=argparse.ArgumentParser();p.add_argument('--group',choices=['core','bridge','all'],default='core');p.add_argument('--out',type=Path,required=True);a=p.parse_args()
    out=a.out.resolve();out.mkdir(parents=True,exist_ok=False);temporary=out/'tmp';temporary.mkdir();tempfile.tempdir=str(temporary)
    os.environ['PYTHONDONTWRITEBYTECODE']='1';sys.dont_write_bytecode=True
    sys.path[:0]=[str(ROOT/'tools'),str(ROOT/'HRS_Bridge'),str(ROOT/'HRS_Bridge/tests'),str(ROOT/'tests')]
    os.environ['HRS_ROOT']=str(ROOT/'HTML_Report_Studio');os.environ['WPS_ROOT']=str(ROOT/'WPS_Forms_Workflow');os.environ['WPS_FORMS_ROOT']=str(ROOT/'WPS_Forms_Workflow')
    modules=[]
    if a.group in ('core','all'):modules+=['test_public_contracts','test_ui_workflow']
    if a.group=='core':modules+=['test_forms_workflow']
    else:modules += [p.stem for p in sorted((ROOT/'HRS_Bridge/tests').glob('test_*.py'))]
    if a.group in ('bridge','all') and not (ROOT/'HRS_Bridge/fixtures/hrs_sample.html').is_file():
        args=[sys.executable,'-X','utf8','-B',str(ROOT/'tools/create_public_fixture.py'),'--archive']
        if os.environ.get('CHROMIUM_PATH'):args+=['--browser',os.environ['CHROMIUM_PATH']]
        subprocess.run(args,check=True)
    results=[]
    for name in modules:
        start=time.perf_counter()
        with (out/(name+'.log')).open('w',encoding='utf8') as log:
            try:result=unittest.TextTestRunner(stream=log,verbosity=2).run(unittest.defaultTestLoader.loadTestsFromName(name))
            except Exception:
                traceback.print_exc(file=log);result=None
        item=dict(module=name,status='PASS' if result and result.wasSuccessful() else 'FAIL',run=result.testsRun if result else 0,failures=len(result.failures) if result else 0,errors=len(result.errors) if result else 1,skipped=len(result.skipped) if result else 0,seconds=round(time.perf_counter()-start,3))
        results.append(item);print(json.dumps(item),flush=True)
    value=dict(status='PASS' if all(x['status']=='PASS' for x in results) else 'FAIL',group=a.group,results=results,python=sys.version,program_model_api_calls=0,native_office='NOT_RUN')
    (out/'RESULT.json').write_text(json.dumps(value,indent=2)+'\n',encoding='utf8')
    return 0 if value['status']=='PASS' else 1
if __name__=='__main__':raise SystemExit(main())

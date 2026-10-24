"""One live two-sample acceptance run; WPS COM and browser are actual processes."""
import argparse
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import time
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from hrs_bridge.browser import recover,export,call
from hrs_bridge.docx_export import convert as to_docx
from hrs_bridge.docx_import import convert as from_docx
from hrs_bridge.runtime import Runtime,sha


def fingerprint(folder):
    return {p.relative_to(folder).as_posix():sha(p.read_bytes()) for p in folder.rglob('*') if p.is_file()}


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--out',type=Path,required=True)
    p.add_argument('--wps-python',type=Path,required=True)
    args=p.parse_args();out=args.out.resolve();out.mkdir(parents=True,exist_ok=False)
    rt=Runtime(Path(os.environ['HRS_ROOT']),Path(os.environ['WPS_ROOT']))
    report={'schema':'hrs-wps-live-samples-v1','steps':[],'overall':'FAIL','platform':platform.platform()}
    before=fingerprint(ROOT/'fixtures')
    def step(name,fn):
        start=time.perf_counter()
        try:
            result=fn()
            item={'name':name,'status':'PASS','elapsed_seconds':round(time.perf_counter()-start,3),'result':result}
            report['steps'].append(item)
            print(name,'PASS',flush=True)
            return result
        except Exception as e:
            report['steps'].append({'name':name,'status':'FAIL','elapsed_seconds':round(time.perf_counter()-start,3),'error':repr(e)})
            raise
    def wps(source,name,edit=False):
        target=out/(name+'.docx')
        command=[str(args.wps_python),'-X','utf8','-B',str(ROOT/'tools/render_wps.py'),str(source),
                 '--pdf',str(out/(name+'.pdf')),'--refresh-out',str(target)]
        if edit:command+=['--edit-cell','3','2','2','WPS_EDIT_20261017']
        r=subprocess.run(command,capture_output=True,text=True,encoding='utf8',timeout=120)
        (out/(name+'.stdout.log')).write_text(r.stdout+r.stderr,encoding='utf8')
        if r.returncode:raise RuntimeError('WPS process failed: '+r.stderr)
        result=json.loads(r.stdout)
        if result.get('status')!='PASS' or result.get('fields_update_result')!=0:raise RuntimeError('WPS fields did not refresh successfully')
        fields=result['field_results']
        if not edit and name=='hrs_wps_saved':
            refs=[f['result'].strip() for f in fields if f['code'].strip().startswith('REF ')]
            assert refs==['5'],refs
            figures=[f['result'].strip() for f in fields if 'SEQ 图 ' in f['code']]
            tables=[f['result'].strip() for f in fields if 'SEQ 表 ' in f['code']]
            assert figures==list(map(str,range(1,11))),figures
            assert tables==list(map(str,range(1,5))),tables
        return result
    try:
        source=out/'hrs_source'
        step('recover_fixed_html',lambda:recover(rt,ROOT/'fixtures/hrs_sample.html',source))
        source_hashes=fingerprint(source)
        hrs_docx=out/'hrs_template.docx'
        receipt=step('hrs_to_docx',lambda:to_docx(rt,source,hrs_docx))
        step('real_wps_refresh_render',lambda:wps(hrs_docx,'hrs_wps_saved'))
        back=out/'hrs_after_wps'
        step('real_wps_saved_to_hrs',lambda:from_docx(rt,out/'hrs_wps_saved.docx',back,provenance=Path(receipt['provenance'])))
        assert fingerprint(back)==source_hashes,'WPS round trip changed Markdown or assets'
        html=out/'hrs_after_wps.html'
        step('original_html_export_after_word',lambda:export(rt,back,html))
        html_back=out/'hrs_after_html'
        step('original_html_recovery_after_word',lambda:recover(rt,html,html_back))
        assert fingerprint(html_back)==source_hashes,'HTML round trip changed Markdown or assets'
        step('real_wps_native_cell_edit',lambda:wps(hrs_docx,'hrs_wps_edited',True))
        edited=out/'hrs_after_native_edit'
        change=step('real_wps_edit_to_hrs',lambda:from_docx(rt,out/'hrs_wps_edited.docx',edited,provenance=Path(receipt['provenance'])))
        ir=rt.analyze((edited/'report.md').read_bytes(),edited/'assets')
        table=next(b for b in ir['blocks'] if b.get('id')=='tbl-deebf0cf6ded')
        tokens=call(rt,'tokens',markdown=table['data']['rows'][1][1]['text'])
        assert ''.join(t['text'] for t in tokens[0]['tokens'])=='WPS_EDIT_20261017'
        assert len(change['changed_blocks'])==1,change['changed_blocks']
        ext=out/'wps_source'
        step('external_wps_sample_to_hrs',lambda:from_docx(rt,ROOT/'fixtures/wps_sample.docx',ext))
        ext_docx=out/'wps_template.docx'
        ext_receipt=step('external_sample_back_to_docx',lambda:to_docx(rt,ext,ext_docx))
        step('external_sample_real_wps_render',lambda:wps(ext_docx,'wps_sample_saved'))
        ext_back=out/'wps_after_save'
        step('external_sample_after_save_to_hrs',lambda:from_docx(rt,out/'wps_sample_saved.docx',ext_back,provenance=Path(ext_receipt['provenance'])))
        assert fingerprint(ext_back)==fingerprint(ext),'External sample derived source changed after WPS save'
        assert fingerprint(source)==source_hashes
        assert fingerprint(ROOT/'fixtures')==before
        report.update(overall='PASS',source_unchanged=True,fixture_fingerprints=before,
                      hrs_source_fingerprints=source_hashes,token_usage=None,token_usage_status='UNAVAILABLE')
    except Exception as error:
        report['error']=repr(error)
        raise
    finally:
        (out/'RESULTS.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    return 0


if __name__=='__main__':raise SystemExit(main())

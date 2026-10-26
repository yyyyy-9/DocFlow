"""Generate/check suite metadata from actual component bytes and constants."""
from pathlib import Path
import argparse,ast,hashlib,json
ROOT=Path(__file__).resolve().parents[1]

def encoded(x):return (json.dumps(x,ensure_ascii=False,indent=2)+'\n').encode('utf8')
def digest(b):return hashlib.sha256(b).hexdigest()
def constants(path,*names):
    values={}
    for n in ast.parse(path.read_text('utf-8-sig')).body:
        if isinstance(n,ast.Assign):
            for t in n.targets:
                if isinstance(t,ast.Name) and t.id in names:values[t.id]=ast.literal_eval(n.value)
    assert set(values)==set(names),(path,values)
    return values

def maintain(write=False):
    def output(path,value):
        data=encoded(value)
        if write:path.write_bytes(data)
        elif path.read_bytes()!=data:raise ValueError('Stale generated metadata: '+str(path.relative_to(ROOT)))
    core=ROOT/'WPS_Forms_Workflow'
    forms=json.loads((core/'BUILD_INFO.json').read_text('utf8'))
    for name in ('release_manifest.json','MANIFEST.json'):
        p=core/name;obj=json.loads(p.read_text('utf8'))
        for key,value in (('version',forms['protocol']),('forms_version',forms['forms_version']),('build_id',forms['build_id'])):
            if key in obj:obj[key]=value
        for row in obj['files']:
            data=(core/row['path']).read_bytes();row.update(bytes=len(data),sha256=digest(data))
        output(p,obj)
    bridge=constants(ROOT/'HRS_Bridge/hrs_bridge/__init__.py','__version__')['__version__']
    p=ROOT/'HRS_Bridge/package.json';meta=json.loads(p.read_text('utf8'));meta['version']=bridge;output(p,meta)
    ui=constants(ROOT/'WPS_Forms_UI/wps_forms_ui.py','VERSION','BUILD_ID')
    assert constants(core/'src/wps_pipeline/forms_workflow.py','FORMS_VERSION')['FORMS_VERSION']==forms['forms_version']
    assert constants(core/'src/wps_pipeline/job_support.py','VERSION')['VERSION']==forms['protocol']
    tool=json.loads((ROOT/'WPS_Forms_UI/tools/template_style/COMPONENT_MANIFEST.json').read_text('utf8'))
    for row in tool['files']:
        data=(ROOT/'WPS_Forms_UI/tools/template_style'/row['path']).read_bytes()
        assert digest(data)==row['sha256'],row['path']
    hrs=json.loads((ROOT/'HTML_Report_Studio/app/release.json').read_text('utf8'))
    value=dict(schema='docflow-components-v1',suite=json.loads((ROOT/'release.json').read_text('utf8'))['version'],
               bridge=dict(version=bridge),ui=dict(version=ui['VERSION'],build=ui['BUILD_ID']),
               forms=dict(version=forms['forms_version'],protocol=forms['protocol'],build=forms['build_id']),
               template_tool=dict(version=tool['tool_version']),hrs=hrs)
    output(ROOT/'COMPONENTS.json',value)
    return value

def main():
    p=argparse.ArgumentParser();g=p.add_mutually_exclusive_group(required=True);g.add_argument('--write',action='store_true');g.add_argument('--check',action='store_true');a=p.parse_args()
    print(json.dumps(dict(status='PASS',components=maintain(a.write)),ensure_ascii=False))
if __name__=='__main__':main()

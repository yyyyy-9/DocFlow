"""Build the HRS-compatible CPython 3.13 component from current Bridge/Core.

The v1 component manifest and open-page protocol stay unchanged. Wheel hashes
are generated once from exact declared dependencies; no application is rewritten.
"""
from pathlib import Path,PurePosixPath
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from email.parser import BytesParser
import argparse
import hashlib
import json
from export_hrs_component import ROOT,snapshot,version

def digest(raw):return hashlib.sha256(raw).hexdigest()
def encoded(value):return (json.dumps(value,ensure_ascii=False,indent=2)+'\n').encode('utf8')

def wheels_lock(folder):
    expected={}
    for line in (ROOT/'requirements.txt').read_text('utf8').splitlines():
        if line and not line.startswith('#'):
            name,value=line.split('==');expected[name.lower().replace('-','_')]=value
    expected.pop('playwright')
    records=[]
    for path in sorted(folder.glob('*.whl')):
        with ZipFile(path) as archive:
            metadata,=[n for n in archive.namelist() if n.endswith('.dist-info/METADATA')]
            info=BytesParser().parsebytes(archive.read(metadata));name=info['Name'].lower().replace('-','_')
        if name not in expected:continue
        if info['Version']!=expected[name]:raise ValueError('Wrong wheel version: '+path.name)
        if '-cp313-cp313-win_amd64.whl' not in path.name and '-py3-none-any.whl' not in path.name:
            raise ValueError('Wrong CPython 3.13 wheel: '+path.name)
        records.append(dict(name=name,version=info['Version'],file=path.name,bytes=path.stat().st_size,sha256=digest(path.read_bytes())))
    if sorted(r['name'] for r in records)!=sorted(expected):raise ValueError('Need exactly four matching CP313 wheels')
    return dict(schema='hrs-bridge-word-wheels-v1',python='3.13',platform='win_amd64',wheels=records)

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--wheels',type=Path,required=True)
    p.add_argument('--lock',type=Path,required=True);p.add_argument('--write-lock',action='store_true')
    p.add_argument('--forms-clean',type=Path);p.add_argument('--hrs-app',type=Path);p.add_argument('--output',type=Path)
    a=p.parse_args()
    if a.write_lock:
        a.lock.write_bytes(encoded(wheels_lock(a.wheels)));print('Locked four CP313 wheels');return
    locked=json.loads(a.lock.read_text('utf8'));files={}
    for wheel in locked['wheels']:
        path=a.wheels/wheel['file']
        if digest(path.read_bytes())!=wheel['sha256']:raise ValueError('Wheel SHA mismatch: '+path.name)
        with ZipFile(path) as archive:
            if archive.testzip() is not None:raise ValueError('Wheel CRC failure')
            for item in archive.infolist():
                if item.is_dir():continue
                name=PurePosixPath(item.filename)
                if name.is_absolute() or '..' in name.parts or '\\' in item.filename:raise ValueError('Unsafe wheel path')
                if name.suffix in {'.h','.c','.pyx','.pxd','.pxi','.pyi'} or any(p in {'tests','__pycache__'} for p in name.parts):continue
                files['libs/'+name.as_posix()]=archive.read(item)
    engine,component=snapshot()
    if component['source_dirty']:raise ValueError('Commit canonical Bridge before building the component')
    for name,raw in engine.items():files['bridge/'+name]=raw
    files['bridge/COMPONENT_MANIFEST.json']=encoded(component)
    with ZipFile(a.forms_clean) as archive:
        if archive.testzip() is not None:raise ValueError('Core ZIP CRC failure')
        for item in archive.infolist():
            if item.is_dir():continue
            name=PurePosixPath(item.filename)
            if name.is_absolute() or '..' in name.parts:raise ValueError('Unsafe Core path')
            files['wps/'+('/'.join(name.parts[1:]))]=archive.read(item)
    if files['bridge/hrs_bridge/browser_model.js']!=(a.hrs_app/'web/bridge_model.js').read_bytes():
        raise ValueError('HRS page model mismatch; this component cannot update that application')
    for name in ('LICENSE','COPYRIGHT'):files[name]=(a.hrs_app/name).read_bytes()
    files['README.txt']=('HRS Word support '+version()+'; CPython3.13/Windows x64.\n'
        'Uses the HRS 2.0 page and model, with the current Bridge and Forms0.1.6 kernel build recorded in MANIFEST.json.\n'
        'Four pinned libraries; no Node/browser/Office runtime is embedded. Pandoc is selected separately.\n'
        'Use a new directory, select it in HRS Word support settings, and retain existing user data.\n').encode()
    files['SOURCE.txt']=('Matching source: HRS_Bridge_v'+version()+'_Full.zip / HRS_Bridge and WPS_Document_Tool.\n'
        'The generated component manifest binds the canonical Bridge commit and file hashes.\n'
        'Existing project and third-party copyright/licence notices are retained.\n').encode()
    engines=dict(bridge=dict(version=version(),commit=component['source_commit'],component_manifest_sha256=digest(files['bridge/COMPONENT_MANIFEST.json'])),
                 wps=json.loads(files['wps/BUILD_INFO.json']))
    manifest=dict(schema='hrs-word-support-v1',python=locked['python'],platform=locked['platform'],engines=engines,
        wheels=locked['wheels'],browser_model_sha256=digest(files['bridge/hrs_bridge/browser_model.js']),
        files=[dict(path=n,bytes=len(b),sha256=digest(b)) for n,b in sorted(files.items())])
    files['MANIFEST.json']=encoded(manifest)
    a.output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(a.output,'x',compression=ZIP_DEFLATED,compresslevel=9) as archive:
        for name,raw in sorted(files.items()):
            info=ZipInfo('word_support/'+name,(2026,10,24,0,0,0));info.compress_type=ZIP_DEFLATED
            info.create_system=3;info.external_attr=0o100644<<16;archive.writestr(info,raw,compresslevel=9)
    with ZipFile(a.output) as archive:
        if archive.testzip() is not None:raise ValueError('Output CRC failure')
        for item in manifest['files']:
            if digest(archive.read('word_support/'+item['path']))!=item['sha256']:raise ValueError('Output hash mismatch')
    print(json.dumps(dict(file=str(a.output),bytes=a.output.stat().st_size,sha256=digest(a.output.read_bytes()),files=len(files)),indent=2))

if __name__=='__main__':main()

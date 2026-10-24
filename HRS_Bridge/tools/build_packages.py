"""Build Clean/Full snapshots from declared inputs; never recurse over runs."""
from pathlib import Path,PurePosixPath
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
import argparse
import hashlib
import json
import subprocess
from export_hrs_component import ROOT,version,snapshot

def digest(raw):return hashlib.sha256(raw).hexdigest()
def encoded(value):return (json.dumps(value,ensure_ascii=False,indent=2)+'\n').encode('utf8')

def archive_files(path):
    with ZipFile(path) as archive:
        if archive.testzip() is not None:raise ValueError('Archive CRC failure: '+str(path))
        result={}
        for name in archive.namelist():
            if name.endswith('/'):continue
            part=PurePosixPath(name)
            if part.is_absolute() or '..' in part.parts or '\\' in name:raise ValueError('Unsafe archive path')
            relative='/'.join(part.parts[1:])
            if relative in result:raise ValueError('Duplicate archive path')
            result[relative]=archive.read(name)
        return result

def build_zip(path,files,identity):
    manifest=dict(schema='hrs-bridge-package-v1',identity=identity,
        files=[dict(path=name,bytes=len(raw),sha256=digest(raw)) for name,raw in sorted(files.items())])
    payload={**files,'MANIFEST.json':encoded(manifest)};prefix='HRS_Bridge_v'+version()+'/'
    with ZipFile(path,'x',compression=ZIP_DEFLATED,compresslevel=9) as archive:
        for name,raw in sorted(payload.items()):
            info=ZipInfo(prefix+name,(2026,10,24,0,0,0));info.compress_type=ZIP_DEFLATED
            info.create_system=3;info.external_attr=0o100644<<16;archive.writestr(info,raw,compresslevel=9)
    with ZipFile(path) as archive:
        if archive.testzip() is not None:raise ValueError('Built ZIP failed CRC')
        for item in manifest['files']:
            if digest(archive.read(prefix+item['path']))!=item['sha256']:raise ValueError('Built ZIP digest mismatch')
    return dict(file=path.name,bytes=path.stat().st_size,sha256=digest(path.read_bytes()),files=len(payload))

def main():
    p=argparse.ArgumentParser(description=__doc__)
    for name in ('hrs-source','forms-clean','wheels','word-wheels','node-module','native-fixture','output'):p.add_argument('--'+name,type=Path,required=True)
    a=p.parse_args();output=a.output.resolve()
    if output.exists() and any(output.iterdir()):raise ValueError('Use a new empty output directory')
    output.mkdir(parents=True,exist_ok=True)
    hrs=archive_files(a.hrs_source);forms=archive_files(a.forms_clean);engine,component=snapshot()
    release=json.loads(hrs['app/release.json']);core=json.loads(forms['BUILD_INFO.json'])
    identity=dict(bridge=version(),hrs=release,forms=core['forms_version'],kernel=core['protocol'],core_build=core['build_id'],
        source_commit=component['source_commit'],source_dirty=component['source_dirty'])
    if component['source_dirty']:raise ValueError('Commit the frozen Bridge source before packaging')
    files={}
    for name,raw in engine.items():files['HRS_Bridge/'+name]=raw
    files['HRS_Bridge/COMPONENT_MANIFEST.json']=encoded(component)
    for name in ('README.md','AGENT_TASK.md','CHANGELOG.md','requirements.txt','package.json'):
        files['HRS_Bridge/'+name]=(ROOT/name).read_bytes()
    for path in sorted((ROOT/'tools').iterdir()):
        if path.is_file() and path.suffix in ('.py','.json'):files['HRS_Bridge/tools/'+path.name]=path.read_bytes()
    files.update({'HTML_Report_Studio/'+name:raw for name,raw in hrs.items() if name.startswith('app/')})
    files.update({'WPS_Document_Tool/'+name:raw for name,raw in forms.items()})
    files['SOURCES.json']=encoded(dict(identity=identity,inputs=[dict(file=path.name,sha256=digest(path.read_bytes())) for path in (a.hrs_source,a.forms_clean)],
        note='One Forms 0.1.6 kernel serves both report_tool and forms_tool. Existing frozen tasks are not rewritten.'))
    files['bridge.py']=b"from pathlib import Path\nimport sys\nsys.path.insert(0,str(Path(__file__).resolve().parent/'HRS_Bridge'))\nfrom hrs_bridge.__main__ import main\nraise SystemExit(main())\n"
    files['run.cmd']=b'@echo off\r\npython -X utf8 -B "%~dp0bridge.py" %*\r\n'
    files['verify.py']=b"from pathlib import Path\nimport hashlib,json\nr=Path(__file__).resolve().parent\nm=json.loads((r/'MANIFEST.json').read_text('utf8'))\nfor x in m['files']:\n p=(r/x['path']).resolve()\n assert p.is_relative_to(r) and p.is_file() and hashlib.sha256(p.read_bytes()).hexdigest()==x['sha256'],x['path']\nprint('PASS',m['identity'],len(m['files']))\n"
    files['README.md']=('''# HRS Bridge '''+version()+'''

在本目录运行 `python bridge.py --version`、`python verify.py`。Python 3.12+。
安装 `python -m pip install -r HRS_Bridge/requirements.txt`；完整命令和三表用法见 HRS_Bridge/README.md。

运行 `python bridge.py run-job job.json`。新三表任务类型是 wps-forms；首稿后从回执取得当前02/03，使用forms-submit生成分类/反馈新稿。
HRS/Markdown/HTML和原生docx/wps-job入口保持。旧0.2.9/0.4.0任务保留原包和原指纹；不能改冻结哈希升级任务。

Clean包含实现、原样HRS2.0核心、原样三表0.1.6共享内核及必要说明，没有测试/运行/个人配置。
Full增加当前测试、固定输入、CPython3.12/Win64离线wheel、CP313 Word组件wheel和playwright-core，不内置Python、Node、Chrome、Pandoc或WPS。
Full可离线安装 `python -m pip install --no-index --find-links offline/wheels -r HRS_Bridge/requirements.txt`。
Word→Word三表不需要Node/Chrome/Pandoc；HRS浏览器CLI需已安装Node20+及Chrome，公式转换需Pandoc。WPS原生接入能力还需本机WPS与pywin32。
仅保留或缺口分别看DELIVERY/finish_list/状态，不以有文件或退出0代替完整性。UI0.1.2独立使用原表单；Bridge管理的批次由forms-submit接续。
''').encode('utf8')
    clean=build_zip(output/('HRS_Bridge_v'+version()+'_Clean.zip'),files,identity)
    full=dict(files)
    for folder in ('tests','fixtures'):
        for path in sorted((ROOT/folder).rglob('*')):
            if path.is_file() and '__pycache__' not in path.parts and path.suffix not in ('.pyc','.pyo'):
                full['HRS_Bridge/'+path.relative_to(ROOT).as_posix()]=path.read_bytes()
    full['WPS_Document_Tool/tests/run_native_numbering.py']=a.native_fixture.read_bytes()
    for path in sorted(a.wheels.glob('*.whl')):full['offline/wheels/'+path.name]=path.read_bytes()
    if not list(a.wheels.glob('*.whl')):raise ValueError('Offline wheels are required for Full')
    from build_word_component import wheels_lock
    locked=json.loads((ROOT/'tools/word_support_wheels.json').read_text('utf8'))
    if wheels_lock(a.word_wheels)!=locked:raise ValueError('Word wheels differ from the committed lock')
    for item in locked['wheels']:
        path=a.word_wheels/item['file'];full['offline/word-wheels/'+path.name]=path.read_bytes()
    package=json.loads((a.node_module/'package.json').read_text('utf8'))
    required=json.loads((ROOT/'package.json').read_text('utf8'))['dependencies']['playwright-core']
    if package['version']!=required:raise ValueError('Pinned playwright-core mismatch')
    for path in sorted(a.node_module.rglob('*')):
        if path.is_file() and '__pycache__' not in path.parts:
            full['HRS_Bridge/node_modules/playwright-core/'+path.relative_to(a.node_module).as_posix()]=path.read_bytes()
    for name in ('LICENSE','COPYRIGHT'):
        if name in hrs:full['HTML_Report_Studio/'+name]=hrs[name]
    complete=build_zip(output/('HRS_Bridge_v'+version()+'_Full.zip'),full,identity)
    (output/'PACKAGE_RECEIPT.json').write_bytes(encoded(dict(identity=identity,clean=clean,full=complete)))
    (output/'SHA256SUMS.txt').write_text('\n'.join(item['sha256']+'  '+item['file'] for item in (clean,complete))+'\n',encoding='utf8')
    print(json.dumps(dict(identity=identity,clean=clean,full=complete),ensure_ascii=False,indent=2))

if __name__=='__main__':main()

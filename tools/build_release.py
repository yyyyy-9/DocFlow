"""Build allowlisted DocFlow releases; exclude tasks, credentials and generated results."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import argparse, hashlib, json, subprocess
from maintain_manifest import maintain
from check_publication import check

ROOT = Path(__file__).resolve().parents[1]
ROOT_FILES = ['.gitattributes', '.gitignore', '.gitmodules', 'AGENTS.md', 'AUTHORS.md', 'CHANGELOG.md',
              'COMPONENTS.json', 'COPYRIGHT', 'LICENSE', 'PROJECT_STATE.md', 'README.md', 'THIRD_PARTY_NOTICES.md',
              'release.json', 'requirements.txt', 'requirements-browser.txt', 'requirements-wps.txt', 'run.py', 'verify.py', '启动界面.pyw']
EXCLUDE = {'.git', '.work', '.build', 'dist', '__pycache__', 'node_modules', 'data', 'outputs'}


def encoded(value): return (json.dumps(value, ensure_ascii=False, indent=2)+'\n').encode('utf8')
def sha(raw): return hashlib.sha256(raw).hexdigest()


def source_files(tests):
    result = {name: (ROOT/name).read_bytes() for name in ROOT_FILES}
    for folder in ['HRS_Bridge', 'WPS_Forms_UI', 'WPS_Forms_Workflow', 'integration'] + (['tools', 'tests', '.github'] if tests else []):
        for path in sorted((ROOT/folder).rglob('*')):
            rel = path.relative_to(ROOT)
            if any(part in EXCLUDE for part in rel.parts) or not path.is_file() or path.suffix in ('.pyc', '.pyo', '.log'): continue
            if path.is_symlink(): raise ValueError('Symlink in release input: '+str(rel))
            if not tests and any(part in ('tests', 'fixtures') for part in rel.parts): continue
            if not tests and rel.parts[:2] == ('HRS_Bridge', 'tools'): continue
            if rel.as_posix() == 'HRS_Bridge/fixtures/hrs_sample.html': continue
            result[rel.as_posix()] = path.read_bytes()
    # The git checkout pins a submodule; downloadable ZIPs include the runnable
    # app only. The app contains its own licenses, notices and dependency data.
    hrs = ROOT/'HTML_Report_Studio'
    for path in sorted((hrs/'app').rglob('*')):
        rel = path.relative_to(hrs)
        if any(part in EXCLUDE for part in rel.parts) or not path.is_file() or path.suffix in ('.pyc', '.pyo'): continue
        result['HTML_Report_Studio/'+rel.as_posix()] = path.read_bytes()
    for name in ('LICENSE', 'COPYRIGHT'): result['HTML_Report_Studio/'+name] = (hrs/name).read_bytes()
    # UI template tool is a verified isolated dependency, needed in every ZIP.
    for path in (ROOT/'WPS_Forms_UI/tools/template_style').rglob('*'):
        if path.is_file() and '__pycache__' not in path.parts:
            result[path.relative_to(ROOT).as_posix()] = path.read_bytes()
    if not tests:
        # UI module doc: the isolated source tool is included above; build and
        # test tools remain in Source_and_Tests, never in the Clean runtime.
        result.pop('AGENTS.md', None)
    return result


def archive(path, files, identity):
    manifest = dict(schema='docflow-package-v1', identity=identity,
                    files=[dict(path=name, bytes=len(raw), sha256=sha(raw)) for name, raw in sorted(files.items())])
    prefix = 'DocFlow_v'+identity['suite']+'/'
    with ZipFile(path, 'x', ZIP_DEFLATED, compresslevel=9) as z:
        for name, raw in sorted({**files, 'MANIFEST.json': encoded(manifest)}.items()):
            info = ZipInfo(prefix+name, (2026, 10, 24, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            info.create_system = 3; info.external_attr = 0o100644 << 16
            z.writestr(info, raw, compresslevel=9)
    with ZipFile(path) as z:
        if z.testzip(): raise ValueError('ZIP CRC failure')
        for row in manifest['files']:
            if sha(z.read(prefix+row['path'])) != row['sha256']: raise ValueError('ZIP hash mismatch')
    return dict(file=path.name, bytes=path.stat().st_size, sha256=sha(path.read_bytes()), files=len(manifest['files']))


def main():
    p = argparse.ArgumentParser(description=__doc__); p.add_argument('--out', type=Path, required=True)
    p.add_argument('--wheels', type=Path); p.add_argument('--node-module', type=Path)
    a = p.parse_args(); out = a.out.resolve()
    if out.exists() and any(out.iterdir()): raise ValueError('Use a new empty release directory')
    out.mkdir(parents=True, exist_ok=True); maintain(); check()
    identity = json.loads((ROOT/'COMPONENTS.json').read_text('utf8'))
    identity['source_commit'] = subprocess.check_output(['git', '-C', str(ROOT), 'rev-parse', 'HEAD'], text=True).strip()
    if subprocess.check_output(['git', '-C', str(ROOT), 'status', '--porcelain']).strip(): raise ValueError('Commit reviewed source before building a release')
    records = []
    for label, tests in [('Clean', False), ('Source_and_Tests', True)]:
        records.append(archive(out/f"DocFlow_v{identity['suite']}_{label}.zip", source_files(tests), identity))
    if bool(a.wheels) != bool(a.node_module): raise ValueError('Full requires both wheels and playwright-core')
    if a.wheels:
        files = source_files(True); wheels = sorted(a.wheels.glob('*.whl'))
        if not wheels: raise ValueError('No offline wheels')
        files.update({'offline/wheels/'+path.name: path.read_bytes() for path in wheels})
        package = json.loads((a.node_module/'package.json').read_text('utf8'))
        expected = json.loads((ROOT/'HRS_Bridge/package.json').read_text('utf8'))['dependencies']['playwright-core']
        if package['name'] != 'playwright-core' or package['version'] != expected: raise ValueError('Node package pin mismatch')
        for path in a.node_module.rglob('*'):
            if path.is_file(): files['HRS_Bridge/node_modules/playwright-core/'+path.relative_to(a.node_module).as_posix()] = path.read_bytes()
        files['offline/README.md'] = b'Offline wheels target CPython 3.12 / Windows x64.\nInstall: python -m pip install --no-index --find-links offline/wheels -r requirements-browser.txt\nplaywright-core is included for an already installed Node 20+.\nPython, Node, Chromium/Chrome, Pandoc and WPS are separate installations.\nNative WPS features additionally require pywin32; install it with requirements-wps.txt.\n'
        records.append(archive(out/f"DocFlow_v{identity['suite']}_Full_Win64_CP312.zip", files, identity))
    (out/'PACKAGE_RECEIPT.json').write_bytes(encoded(dict(identity=identity, packages=records)))
    (out/'SHA256SUMS.txt').write_text(''.join(row['sha256']+'  '+row['file']+'\n' for row in records), 'utf8')
    print(json.dumps(records, ensure_ascii=False, indent=2))


if __name__ == '__main__': main()

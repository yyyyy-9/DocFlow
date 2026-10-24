"""Verify an extracted DocFlow release using only Python's standard library."""
from pathlib import Path, PurePosixPath
import hashlib, json

ROOT = Path(__file__).resolve().parent
IGNORED = {'.git', '.work', '.build', '.venv', '__pycache__', 'outputs'}


def verify():
    path = ROOT/'MANIFEST.json'
    if not path.is_file():
        raise ValueError('This is a source checkout. Use tools/maintain_manifest.py --check instead.')
    manifest = json.loads(path.read_text('utf8')); expected = set()
    for row in manifest['files']:
        relative = PurePosixPath(row['path'])
        if relative.is_absolute() or '..' in relative.parts or '\\' in row['path'] or row['path'] in expected:
            raise ValueError('Invalid manifest path')
        expected.add(row['path']); target = ROOT.joinpath(*relative.parts)
        if target.is_symlink() or not target.is_file(): raise ValueError('Missing/linked file: '+row['path'])
        raw = target.read_bytes()
        if len(raw) != row['bytes'] or hashlib.sha256(raw).hexdigest() != row['sha256']:
            raise ValueError('Changed release file: '+row['path'])
    actual = set()
    for target in ROOT.rglob('*'):
        rel = target.relative_to(ROOT)
        if any(part in IGNORED for part in rel.parts) or rel.parts[:2] == ('WPS_Forms_UI', 'data'): continue
        if target.is_file() and target.suffix not in ('.pyc', '.pyo') and rel.as_posix() != 'MANIFEST.json': actual.add(rel.as_posix())
    if actual != expected: raise ValueError('Release file set differs: '+str(sorted(actual ^ expected)[:20]))
    return dict(status='PASS', files=len(expected), identity=manifest['identity'])


if __name__ == '__main__': print(json.dumps(verify(), ensure_ascii=False))

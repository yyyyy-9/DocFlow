"""Small file-backed job protocol. No model service or global application state."""
from pathlib import Path, PurePosixPath, PureWindowsPath
from datetime import datetime, timezone
import hashlib
import json
import os
import tempfile
import time
import traceback

VERSION = '0.5.7'
ROOT = Path(__file__).resolve().parents[2]


def _read_bytes(path):
    # Atomic replacement may briefly deny opening a file on Windows too.
    # Retry only sharing/permission errors; malformed JSON and missing input
    # stay explicit errors, and a held lock still fails within a bounded time.
    for attempt in range(8):
        try:return Path(path).read_bytes()
        except PermissionError:
            if attempt==7:raise
            time.sleep(.01*(attempt+1))


def read(path):
    return json.loads(_read_bytes(path).decode('utf-8-sig'))


def write(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    # Unique siblings prevent two writers from replacing each other's staged
    # data. A short Windows sharing lock may outlive a reader's close.
    payload=json.dumps(value, ensure_ascii=False, indent=2)
    with tempfile.NamedTemporaryFile(mode='w',encoding='utf-8',dir=path.parent,prefix='.'+path.name+'.',suffix='.tmp',delete=False) as stream:
        stream.write(payload);stream.flush();temp=Path(stream.name)
    try:
        for attempt in range(8):
            try:os.replace(temp,path);break
            except PermissionError:
                if attempt==7:raise
                time.sleep(.01*(attempt+1))
    finally:
        if temp.exists():temp.unlink()


def sha(path):
    return hashlib.sha256(_read_bytes(path)).hexdigest()


def stamp():
    return datetime.now(timezone.utc).isoformat()


def event(run, stage, status, **values):
    record = dict(time=stamp(), stage=stage, status=status, **values)
    with (Path(run) / 'events.jsonl').open('a', encoding='utf-8') as stream:
        stream.write(json.dumps(record, ensure_ascii=False) + '\n')
    return record


def stage(run, key, callback, evidence):
    """Every attempted stage leaves a receipt, including its actual exception."""
    event(run, key, 'RUNNING')
    try:
        result = callback()
        write(Path(run) / evidence, result)
        status = result.get('status', 'DONE') if isinstance(result, dict) else 'DONE'
        return event(run, key, status, evidence=evidence,
                     evidence_sha256=sha(Path(run) / evidence))
    except Exception as exc:
        failure = dict(status='FAILED', error=str(exc), exception=type(exc).__name__,
                       traceback=traceback.format_exc())
        write(Path(run) / evidence, failure)
        return event(run, key, 'FAILED', evidence=evidence,
                     evidence_sha256=sha(Path(run) / evidence), error=str(exc))


def contained(run, relative):
    if not isinstance(relative, (str, Path)) or not str(relative):
        raise ValueError('A non-empty relative path is required')
    value = str(relative)
    # Frozen v0.4.0 evidence may contain Windows separators. Accept those while
    # rejecting absolute, drive-relative, UNC and parent traversal paths.
    windows = PureWindowsPath(value)
    portable = PurePosixPath(value.replace('\\', '/'))
    if windows.drive or windows.root or portable.is_absolute() or '..' in portable.parts:
        raise ValueError('Job-local path must be relative without parent traversal')
    path = (Path(run) / Path(*portable.parts)).resolve()
    if not path.is_relative_to(Path(run).resolve()):
        raise ValueError('Job-local path escapes its run directory')
    return path


def fingerprints():
    manifest_path = ROOT / 'runtime_manifest.json'
    manifest = read(manifest_path)
    if manifest.get('version') != VERSION:
        raise ValueError('Runtime manifest version does not match this release')
    files = manifest.get('files')
    if not isinstance(files, list) or not files:
        raise ValueError('runtime_manifest.json must contain a non-empty files list')
    result = {}
    seen = set()
    for relative in files:
        if not isinstance(relative, str) or '\\' in relative:
            raise ValueError('Manifest files must use normalized relative paths')
        path = contained(ROOT, relative)
        if PurePosixPath(relative).as_posix() != relative or relative in ('', '.'):
            raise ValueError('Manifest path is not normalized: ' + relative)
        if relative.casefold() in seen:
            raise ValueError('Duplicate manifest file: ' + relative)
        seen.add(relative.casefold())
        if not path.is_file():
            raise FileNotFoundError('Required runtime file missing: ' + relative)
        result[relative] = sha(path)
    # Bind the explicit inventory as well, so removing an entry changes the run.
    result['runtime_manifest.json'] = sha(manifest_path)
    return result


def check_inputs(run, state):
    if state.get('version') != VERSION:
        raise ValueError('Run version differs from this runtime. Use an explicit new revision; do not resume old evidence.')
    check_frozen_inputs(run, state)
    if state['code_sha256'] != fingerprints():
        raise ValueError('Runtime files changed. Create a new revision; do not resume old evidence.')
    for key, record in state.get('stages', {}).items():
        if sha(contained(run, record['evidence'])) != record['evidence_sha256']:
            raise ValueError('Recorded stage evidence changed: ' + key)
    for src in state['sources']:
        record=state.get('stages',{}).get('scan:'+src['id'])
        if record and record.get('status') not in ('FAILED','FAIL'):
            if sha(contained(run,'sources/'+src['id']+'/scan/scan.json'))!=record['evidence_sha256']:
                raise ValueError('Scan facts changed: ' + src['id'])
        decisions_path=contained(run,'sources/'+src['id']+'/decisions.json')
        candidate_path=contained(run,'sources/'+src['id']+'/candidate.json')
        if candidate_path.exists():
            candidate=read(candidate_path)
            bound={row['block_id']:row['decision'] for row in candidate.get('blocks',[]) if 'decision' in row}
            decisions=read(decisions_path) if decisions_path.exists() else []
            if (not isinstance(decisions,list) or len(decisions)!=len(bound)
                    or any(bound.get(d.get('block_id'))!=d for d in decisions)):
                raise ValueError('Decision provenance changed: ' + src['id'])


def check_frozen_inputs(run, state):
    """Check immutable inputs without imposing the current runtime version."""
    for item in state['sources'] + [state['template']]:
        if sha(contained(run, item['frozen'])) != item['sha256']:
            raise ValueError('Frozen input changed: ' + item['frozen'])


def issue(kind, message, source_id=None, block_id=None, **extra):
    return dict(kind=kind, message=message, source_id=source_id, block_id=block_id, **extra)

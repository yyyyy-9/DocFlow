"""Export/check a generated HRS component from the one Bridge source tree."""
from pathlib import Path
import argparse
import ast
import hashlib
import json
import subprocess

ROOT=Path(__file__).resolve().parents[1]


def version():
    tree=ast.parse((ROOT/'hrs_bridge/__init__.py').read_text('utf8'))
    return next(node.value.value for node in tree.body if isinstance(node,ast.Assign)
        and any(isinstance(target,ast.Name) and target.id=='__version__' for target in node.targets))


def files():
    return {'hrs_bridge/'+p.name:p.read_bytes() for p in sorted((ROOT/'hrs_bridge').iterdir())
            if p.is_file() and p.suffix in ('.py','.js','.cjs')}


def snapshot():
    payload=files();records=[dict(path=name,bytes=len(raw),sha256=hashlib.sha256(raw).hexdigest()) for name,raw in payload.items()]
    return payload,dict(schema='hrs-bridge-component-v1',version=version(),
        source_commit=subprocess.check_output(['git','-C',str(ROOT),'rev-parse','HEAD'],text=True).strip(),
        source_dirty=bool(subprocess.check_output(['git','-C',str(ROOT),'status','--porcelain']).strip()),files=records)


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--out',type=Path,required=True)
    mode=parser.add_mutually_exclusive_group(required=True);mode.add_argument('--write',action='store_true');mode.add_argument('--check',action='store_true')
    args=parser.parse_args();target=args.out.resolve();payload,metadata=snapshot()
    if args.write:
        if target.exists() and any(target.iterdir()):raise ValueError('Export to a new empty directory; review replacement separately')
        target.mkdir(parents=True,exist_ok=True)
        for name,raw in payload.items():
            path=target/name;path.parent.mkdir(exist_ok=True);path.write_bytes(raw)
        (target/'COMPONENT_MANIFEST.json').write_text(json.dumps(metadata,indent=2)+'\n',encoding='utf8')
    else:
        actual={p.relative_to(target).as_posix() for p in (target/'hrs_bridge').iterdir() if p.is_file() and p.suffix in ('.py','.js','.cjs')}
        if actual!=set(payload):raise ValueError('Component file set differs from canonical source')
        for name,raw in payload.items():
            if (target/name).read_bytes()!=raw:raise ValueError('Generated component differs: '+name)
        saved=json.loads((target/'COMPONENT_MANIFEST.json').read_text('utf8'))
        if saved['version']!=metadata['version'] or saved['files']!=metadata['files']:raise ValueError('Component manifest is stale')
    print(json.dumps(dict(status='PASS',version=metadata['version'],files=len(payload))))


if __name__=='__main__':main()

"""Keep npm's generated identity aligned with hrs_bridge.__version__."""
from pathlib import Path
import argparse
import json
from export_hrs_component import ROOT,version

p=argparse.ArgumentParser(description=__doc__);p.add_argument('--check',action='store_true');a=p.parse_args()
path=ROOT/'package.json';data=json.loads(path.read_text('utf8'))
if a.check:
    if data.get('version')!=version():raise SystemExit('Stale package.json version; run tools/maintain_metadata.py')
else:
    data['version']=version();path.write_text(json.dumps(data,indent=2)+'\n',encoding='utf8')
print('Metadata matches canonical Bridge '+version())

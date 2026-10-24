"""Unified standard-library dispatcher; processing remains in the original components."""
from pathlib import Path
import argparse,json,subprocess,sys
ROOT=Path(__file__).resolve().parent

def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--version',action='store_true');p.add_argument('component',nargs='?',choices=['ui','bridge','forms'])
    p.add_argument('arguments',nargs=argparse.REMAINDER);a=p.parse_args(argv)
    if a.version:
        print('DocFlow '+json.loads((ROOT/'release.json').read_text('utf8'))['version']);return 0
    if not a.component:p.print_help();return 0
    if a.component=='bridge':
        sys.path.insert(0,str(ROOT/'HRS_Bridge'))
        from hrs_bridge.__main__ import main as bridge_main
        return bridge_main(a.arguments)
    else:
        folder='WPS_Forms_UI' if a.component=='ui' else 'WPS_Forms_Workflow'
        entry='ui_tool.py' if a.component=='ui' else 'forms_tool.py'
        cmd=[sys.executable,'-X','utf8','-B',str(ROOT/folder/entry),*a.arguments]
    return subprocess.call(cmd)

if __name__=='__main__':raise SystemExit(main())

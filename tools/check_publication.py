"""Audit public project files and Office XML; never traverse user run data."""
from pathlib import Path
from zipfile import ZipFile
import argparse,ast,json,re
ROOT=Path(__file__).resolve().parents[1]
EXCLUDED={'.git','.work','.build','dist','node_modules','__pycache__','data','outputs'}
TEXT={'.py','.pyw','.md','.txt','.json','.js','.cjs','.css','.html','.yml','.yaml','.toml'}
ORG_FILES={'recognize_reports.py','verify_report_build.py'}

def files():
    for p in ROOT.rglob('*'):
        rel=p.relative_to(ROOT)
        if any(part in EXCLUDED for part in rel.parts) or not p.is_file():continue
        # Upstream HRS is a fixed, separately published repository.
        if rel.parts[0]=='HTML_Report_Studio':continue
        yield p

def check():
    errors=[];org=[];count=0
    for p in files():
        rel=p.relative_to(ROOT).as_posix();count+=1
        if p.suffix in TEXT:
            text=p.read_text('utf-8-sig')
            # The audit's own literal patterns are rules, not background prose.
            if p.resolve()==Path(__file__).resolve():
                ast.parse(text,filename=rel);continue
            if re.search(r'[A-Za-z]:[/\\]+Users[/\\]+[^/\\\s]+|wxid_|xwechat_files|codex-clipboard-|多Agent测试|测试项目',text,re.I):errors.append(rel+': private path/context')
            for line in text.splitlines():
                if re.search('公司|内网|内部机',line):
                    if p.name in ORG_FILES and '有限公司' in line and re.search(r'INSTITUTION=|institution=bool',line) and not re.search('内网|内部机',line):org.append(rel)
                    else:errors.append(rel+': private contextual wording')
            if p.suffix in {'.py','.pyw'}:ast.parse(text,filename=rel)
        if p.suffix in {'.docx','.xlsx'}:
            with ZipFile(p) as archive:
                if archive.testzip():errors.append(rel+': ZIP CRC')
                for name in archive.namelist():
                    if not name.endswith('.xml'):continue
                    text=archive.read(name).decode('utf8',errors='replace')
                    if re.search('公司|内网|内部机|溢 杨',text):errors.append(rel+': document context '+name)
                    if re.search(r'[A-Za-z]:[/\\]+Users[/\\]+[^/\\\s]+',text,re.I):errors.append(rel+': document path '+name)
    if len(org)!=2:errors.append('Expected exactly two approved organization-term recognition rules')
    if errors:raise ValueError('\n'.join(errors))
    return dict(status='PASS',project_files=count,contextual_wording_hits=0,authorized_organization_rules=org)
if __name__=='__main__':print(json.dumps(check(),ensure_ascii=False))

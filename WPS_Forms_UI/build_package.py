"""Build the UI plus byte-identical Forms kernel; no user data or test assets."""
import ast
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

HERE = Path(__file__).resolve().parent


def string_constants(content, *names):
    """Read release identity from the bytes being packaged, without importing it."""
    values = {}
    for node in ast.parse(content.decode('utf-8-sig')).body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id in names:
                    if target.id in values: raise ValueError('重复身份字段：' + target.id)
                    value = ast.literal_eval(node.value)
                    if not isinstance(value, str) or not value: raise ValueError('身份字段无效：' + target.id)
                    values[target.id] = value
    if set(values) != set(names): raise ValueError('发布身份字段缺失')
    return values


def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('--core', type=Path, default=HERE.parent / 'WPS_Forms_Workflow')
    p.add_argument('--out', type=Path, required=True)
    a = p.parse_args(); core = a.core.resolve(); output = a.out.resolve()
    if output.exists(): raise ValueError('发行包使用新路径，不覆盖已有包')
    manifest = json.loads((core / 'MANIFEST.json').read_text('utf8'))
    files = {}
    for item in manifest['files']:
        path = core / item['path']; content = path.read_bytes()
        if len(content) != item['bytes'] or hashlib.sha256(content).hexdigest() != item['sha256']:
            raise ValueError('内核快照不一致：' + item['path'])
        if 'core/' + item['path'] in files: raise ValueError('内核清单重复：' + item['path'])
        files['core/' + item['path']] = content
    files['core/MANIFEST.json'] = (core / 'MANIFEST.json').read_bytes()
    for name in ('ui_tool.py', 'wps_forms_ui.py', 'wps_forms_native.py', '启动界面.pyw', 'README.md', 'requirements.txt', 'web/index.html', 'web/style.css', 'web/app.js'):
        content = (HERE / name).read_bytes()
        if name.endswith(('.py', '.pyw')): ast.parse(content.decode('utf-8-sig'), filename=name)
        files[name] = content
    component_root = HERE / 'tools/template_style'
    component = json.loads((component_root / 'COMPONENT_MANIFEST.json').read_text('utf8'))
    component_names = set()
    for item in component['files']:
        path = (component_root / item['path']).resolve()
        if not path.is_relative_to(component_root.resolve()) or item['path'] in component_names:
            raise ValueError('模板工具组件清单路径无效')
        content = path.read_bytes()
        if len(content) != item['bytes'] or hashlib.sha256(content).hexdigest() != item['sha256']:
            raise ValueError('模板工具组件快照不一致：' + item['path'])
        if path.suffix == '.py': ast.parse(content.decode('utf-8-sig'), filename=item['path'])
        component_names.add(item['path']); files['tools/template_style/' + item['path']] = content
    if {p.relative_to(component_root).as_posix() for p in component_root.rglob('*') if p.is_file()} != component_names | {'COMPONENT_MANIFEST.json'}:
        raise ValueError('模板工具组件存在未登记文件')
    files['tools/template_style/COMPONENT_MANIFEST.json'] = (component_root / 'COMPONENT_MANIFEST.json').read_bytes()
    tool_version = string_constants(files['tools/template_style/src/wps_pipeline/template_runtime.py'], 'TOOL_VERSION')['TOOL_VERSION']
    if component['tool_version'] != tool_version: raise ValueError('模板工具身份不一致')
    identity = string_constants(files['wps_forms_ui.py'], 'VERSION', 'BUILD_ID')
    build = json.loads(files['core/BUILD_INFO.json'].decode('utf-8-sig'))
    forms = string_constants(files['core/src/wps_pipeline/forms_workflow.py'], 'FORMS_VERSION')['FORMS_VERSION']
    protocol = string_constants(files['core/src/wps_pipeline/job_support.py'], 'VERSION')['VERSION']
    if build['forms_version'] != forms or build['protocol'] != protocol or build['build_id'] != manifest['build_id']:
        raise ValueError('内核代码、构建身份和清单不一致')
    release = dict(ui_version=identity['VERSION'], ui_build=identity['BUILD_ID'],
                   forms_version=forms, kernel_protocol=protocol, core_build=build['build_id'],
                   template_tool=dict(version=tool_version, source_archive_sha256=component['upstream_sha256']),
                   files=[dict(path=name, bytes=len(data), sha256=hashlib.sha256(data).hexdigest()) for name, data in sorted(files.items())])
    files['UI_MANIFEST.json'] = (json.dumps(release, ensure_ascii=False, indent=2) + '\n').encode('utf8')
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, 'x', ZIP_DEFLATED) as z:
        for name, content in sorted(files.items()): z.writestr('WPS_Forms_UI/' + name, content)
    print(json.dumps(dict(file=str(output), ui_version=release['ui_version'], ui_build=release['ui_build'],
                          forms_version=forms, kernel_protocol=protocol, core_build=build['build_id'],
                          files=len(files), bytes=output.stat().st_size,
                          sha256=hashlib.sha256(output.read_bytes()).hexdigest(), core_unchanged=True), ensure_ascii=False))


if __name__ == '__main__': main()

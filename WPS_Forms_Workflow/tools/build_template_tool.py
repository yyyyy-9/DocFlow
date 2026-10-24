"""Build one Windows EXE from the same template backend, in explicit new paths.

Build dependency: PyInstaller 6.22.3. End users need neither Python nor Office.
"""
from pathlib import Path
import argparse,hashlib,importlib.metadata,json,shutil,subprocess,sys
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
from wps_pipeline.template_runtime import TOOL_VERSION

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--out',type=Path,required=True);p.add_argument('--work',type=Path,required=True);a=p.parse_args()
    if sys.platform!='win32':raise RuntimeError('Build the Windows EXE on Windows')
    if importlib.metadata.version('pyinstaller')!='6.22.3':raise RuntimeError('Use pinned PyInstaller 6.22.3')
    out=a.out.resolve();work=a.work.resolve();out.mkdir(parents=True,exist_ok=False);work.mkdir(parents=True,exist_ok=False)
    version=tuple(int(n) for n in TOOL_VERSION.split('.'))+(0,)
    version_file=work/'version_info.txt'
    version_file.write_text(f'''VSVersionInfo(ffi=FixedFileInfo(filevers={version!r},prodvers={version!r},mask=0x3f,flags=0,OS=0x40004,fileType=0x1,subtype=0,date=(0,0)),kids=[StringFileInfo([StringTable('080404b0',[StringStruct('FileDescription','WPS 模板样式工具'),StringStruct('FileVersion','{TOOL_VERSION}'),StringStruct('ProductVersion','{TOOL_VERSION}'),StringStruct('ProductName','WPS 模板样式工具 正式版')])]),VarFileInfo([VarStruct('Translation',[2052,1200])])])''',encoding='utf8')
    cmd=[sys.executable,'-m','PyInstaller','--onefile','--windowed','--noupx','--name','WPS_Template_Style_Tool',
         '--distpath',str(work/'dist'),'--workpath',str(work/'build'),'--specpath',str(work),
         '--version-file',str(version_file),
         '--paths',str(ROOT/'src'),'--hidden-import','lxml.etree',
         '--add-data',str(ROOT/'config.json')+':wps_tool',
         '--add-data',str(ROOT/'templates/target_template.docx')+':wps_tool/templates',
         '--add-data',str(ROOT/'template_tool.py')+':wps_tool']
    # Ship the exact same source for an explicitly selected Python. Do not add
    # another interpreter or depend on a separate source checkout at run time.
    source_files=[ROOT/'template_tool.py',ROOT/'config.json',ROOT/'templates/target_template.docx',*sorted((ROOT/'src/wps_pipeline').glob('*.py'))]
    for source in source_files[3:]:
        cmd+=['--add-data',str(source)+':wps_tool/src/wps_pipeline']
    cmd.append(str(ROOT/'template_tool.py'))
    proc=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True,encoding='utf8',errors='replace')
    (work/'build.log').write_text(proc.stdout+proc.stderr,encoding='utf8')
    if proc.returncode:raise RuntimeError('Build failed; see '+str(work/'build.log'))
    target=out/'WPS模板样式工具.exe';shutil.copyfile(work/'dist/WPS_Template_Style_Tool.exe',target)
    licenses=out/'licenses';licenses.mkdir()
    for package in ('pyinstaller','pyinstaller-hooks-contrib','lxml'):
        dist=importlib.metadata.distribution(package)
        for entry in dist.files or []:
            if Path(str(entry)).name.lower().startswith(('license','copying')):
                source=Path(dist.locate_file(entry))
                if source.is_file():shutil.copyfile(source,licenses/(package+'_'+source.name))
    for name in ('tcl8.6','tk8.6'):
        source=Path(sys.base_prefix)/'tcl'/name/'license.terms'
        if source.is_file():shutil.copyfile(source,licenses/(name+'_license.terms'))
    # The bundled Python build omits Tcl's notice even though it includes Tcl.
    # Keep the exact upstream notice with this builder instead of fetching at run time.
    shutil.copyfile(ROOT/'tools/licenses/Tcl_8.6.15_license.terms',licenses/'Tcl_8.6.15_license.terms')
    source=Path(sys.base_prefix)/'LICENSE.txt'
    if source.is_file():shutil.copyfile(source,licenses/'Python_LICENSE.txt')
    for source in source_files+[Path(__file__),ROOT/'tools/licenses/Tcl_8.6.15_license.terms']:
        destination=out/'source'/source.relative_to(ROOT)
        destination.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,destination)
    (out/'使用说明.txt').write_text('''WPS 模板样式工具 '''+TOOL_VERSION+''' 正式版

直接双击 EXE，无需安装 Python，不需要启动 Word 或 WPS。
1. 样式来源默认是内置母模板 v0.3.11；也可选择包含同名报告样式的 DOCX。
2. 选择带封面、目录、页眉的目标模板，再选择新模板的保存位置。
3. 点击“读取样式与模板内容”。可切换“样式信息”查看字体、字号及段落信息。
4. 在“选择正文起点”列表中选第一个要替换的块。此前的封面/目录正文全部保留。
5. 核对预览后点击“生成新模板”。生成 DOCX 和可追溯的 template.json，原件不改。

如果同一文件能由 template_tool.py 读取，但内置方式失败：
在“运行方式”选择“指定 Python”，选择你运行源码时所用环境的 python.exe，重新读取。
读取、摘要校验与生成均由这个 Python 执行，EXE 保留窗口；所选环境需有 lxml，不要求 Tk。
工具已内置对应源码，无需另外找源码目录；运行方式仅作为本机偏好保存。
“导出读取诊断”对照内置/指定 Python 的进程、读取长度、SHA256、文件头和 ZIP 结构。
诊断不保存文档正文、不上传内容。探测是分别读取，不是对失败瞬间的原子取证。
工具使用所选 Python 读取文件；文件访问权限保持不变，兼容性请在目标运行环境验证。
缺 Python/依赖、超时或旧快照均明确失败，不自动切换、不重试写入。
源码位于 source/template_tool.py，可直接由原有 Python 环境运行。

输出文件可作为 WPS 文档工具或 Bridge 的母模板；伴随 JSON 只是记录，不是使用依赖。
目录保留原生域，装配报告后请在 Word/WPS 中更新目录。
新模板正文是普通空段落，以 WPS_ReportBody 书签定位；没有正文大框、固定报告题名或占位章号。
装配报告时会移除该空段落并写入实际正文；封面、目录域、页眉页脚继续保留。
已有书签模板可直接使用；旧版未填写的正文大框可选中其起点后升级。旧框已被填写时拒绝覆盖。
复杂多节正文需明确兼容范围。
不会猜测正文起点、补写封面内容或覆盖已有输出。所选样式来源缺少必需样式时，请使用内置母模板。
''',encoding='utf8')
    metadata={'tool_version':TOOL_VERSION,'release_channel':'stable','python':sys.version,'pyinstaller':'6.22.3','lxml':importlib.metadata.version('lxml'),
              'source_files':{p.relative_to(ROOT).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in source_files},
              'exe':{'file':target.name,'bytes':target.stat().st_size,'sha256':hashlib.sha256(target.read_bytes()).hexdigest()},
              'build_command':cmd}
    (out/'BUILD_MANIFEST.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    print(json.dumps(metadata['exe'],ensure_ascii=False),flush=True)

if __name__=='__main__':main()

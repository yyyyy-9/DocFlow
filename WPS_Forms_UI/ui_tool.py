"""Local, offline UI for the unchanged WPS Forms kernel."""
import argparse
import json
import sys
from pathlib import Path

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent


def template_window():
    """Open the bundled, isolated source tool with this Python; report GUI readiness."""
    import os
    import runpy
    import tkinter as tk
    os.environ['WPS_TEMPLATE_PYTHON'] = sys.executable
    entry = HERE / 'tools' / 'template_style' / 'template_tool.py'
    scope = runpy.run_path(str(entry))
    if sys.platform == 'win32':
        import ctypes
        try: ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except OSError: pass
    root = tk.Tk()
    scope['TemplateWindow'](root)
    root.after_idle(lambda: print(json.dumps(dict(event='TEMPLATE_TOOL_READY', version=scope['TOOL_VERSION'],
                                                  python=sys.executable)), flush=True))
    root.mainloop()
    return 0


def main():
    p = argparse.ArgumentParser(description='WPS文档整理本地界面（无模型API）')
    p.add_argument('command', nargs='?', default='serve', choices=['serve', 'choose', 'template-tool'])
    p.add_argument('--core', type=Path)
    p.add_argument('--data', type=Path, default=HERE / 'data')
    p.add_argument('--port', type=int, default=0)
    p.add_argument('--no-browser', action='store_true')
    p.add_argument('--kind', choices=['directory', 'docx', 'xlsx'], default='directory')
    a = p.parse_args()
    if a.command == 'template-tool': return template_window()
    if a.command == 'choose':
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
        if a.kind == 'directory':
            path = filedialog.askdirectory(parent=root)
        else:
            path = filedialog.askopenfilename(parent=root, filetypes=[(a.kind.upper(), '*.' + a.kind)])
        root.destroy()
        print(json.dumps({'path': path}, ensure_ascii=True))
        return 0
    core = a.core or (HERE / 'core' if (HERE / 'core').is_dir() else HERE.parent / 'WPS_Forms_Workflow')
    from wps_forms_ui import serve
    return serve(core.resolve(), a.data.resolve(), a.port, not a.no_browser)


if __name__ == '__main__':
    raise SystemExit(main())

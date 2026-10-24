"""Source launcher: the selected Python runs the server and all document I/O."""
import json
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

HERE = Path(__file__).resolve().parent
DATA = HERE / 'data'
DATA.mkdir(exist_ok=True)
PREF = DATA / 'launcher.json'
try: previous = json.loads(PREF.read_text('utf8'))
except Exception: previous = {}

root = tk.Tk(); root.title('WPS 文档整理 · 启动'); root.geometry('690x265')
frame = ttk.Frame(root, padding=18); frame.pack(fill='both', expand=True)
ttk.Label(frame, text='选择本机可读取文档的 Python，界面与文档程序共用它。').pack(anchor='w')
value = tk.StringVar(value=previous.get('python') or sys.executable)
line = ttk.Frame(frame); line.pack(fill='x', pady=12)
ttk.Entry(line, textvariable=value).pack(side='left', fill='x', expand=True)
ttk.Button(line, text='选择 Python…', command=lambda: value.set(filedialog.askopenfilename(filetypes=[('Python', 'python*.exe')]) or value.get())).pack(side='left', padx=(8, 0))
status = tk.StringVar(value='界面在本机浏览器运行；关闭此窗口不会关闭已打开的Word。')
ttk.Label(frame, textvariable=status, wraplength=640).pack(anchor='w', pady=8)
child = None


def launch():
    global child
    try:
        executable = Path(value.get()).resolve()
        if not executable.is_file(): raise ValueError('请选择存在的Python程序')
        PREF.write_text(json.dumps({'python': str(executable)}, ensure_ascii=False, indent=2), encoding='utf8')
        if (DATA / 'server.json').exists():
            info = json.loads((DATA / 'server.json').read_text('utf8'))
            import urllib.request
            try:
                base = info['url'].split('/#')[0]
                request = urllib.request.Request(base + '/api/state', headers={'X-WPS-UI-Token': info['url'].split('token=')[1]})
                with urllib.request.urlopen(request, timeout=2) as reply: current = json.loads(reply.read())
            except Exception:
                current = None
            if current:
                if Path(current['python']).resolve() != executable:
                    raise ValueError('原服务仍在使用另一Python，请先在网页退出服务，再切换运行环境。')
                import webbrowser
                webbrowser.open(info['url']); status.set('已打开现有服务，不创建第二个并发服务。')
                return
        log = (DATA / 'launcher.log').open('a', encoding='utf8')
        child = subprocess.Popen([str(executable), '-B', str(HERE / 'ui_tool.py'), 'serve', '--data', str(DATA)],
                                 cwd=HERE, stdout=log, stderr=log,
                                 creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        log.close()
        status.set('已启动。进度与错误留在网页日志；启动失败记录在 data/launcher.log。')
        root.after(1600, check)
    except Exception as exc:
        messagebox.showerror('启动失败', str(exc))


def check():
    if child is not None and child.poll() is not None:
        status.set('服务已退出，请查看启动日志；窗口保持打开。')


ttk.Button(frame, text='启动 / 打开整理界面', command=launch).pack(anchor='w', pady=10)
ttk.Button(frame, text='查看启动日志', command=lambda: os.startfile(str(DATA / 'launcher.log')) if (DATA / 'launcher.log').exists() else None).pack(anchor='w')
root.mainloop()

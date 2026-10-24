"""Small desktop front end for the existing, source-bound template preparer."""
from pathlib import Path
import argparse,json,sys
ROOT=Path(__file__).resolve().parent
if not getattr(sys,'frozen',False):sys.path.insert(0,str(ROOT/'src'))
from wps_pipeline import resource_root
from wps_pipeline.template_runtime import (TOOL_VERSION, TemplateBackend, execute_request,
    load_python_preference, save_python_preference)


def paragraph_summary(row):
    p=row['paragraph'];spacing=p['spacing'];indent=p['indent'];items=[]
    if p['alignment']:items.append({'both':'两端对齐','left':'左对齐','center':'居中','right':'右对齐','distribute':'分散对齐'}.get(p['alignment'],p['alignment']))
    for key,label in (('before','段前'),('after','段后')):
        if key in spacing:items.append(label+f" {int(spacing[key])/20:g} pt")
    if 'line' in spacing:
        line=int(spacing['line']);rule=spacing.get('lineRule','auto')
        items.append(f'行距 {line/240:g} 倍' if rule=='auto' else ('固定行距' if rule=='exact' else '最小行距')+f' {line/20:g} pt')
    if 'firstLineChars' in indent:items.append(f"首行 {int(indent['firstLineChars'])/100:g} 字符")
    elif 'firstLine' in indent:items.append(f"首行 {int(indent['firstLine'])/20:g} pt")
    if row['numbering']['num_id']:items.append('原生编号')
    return '；'.join(items)


class TemplateWindow:
    def __init__(self,root):
        import tkinter as tk
        from tkinter import ttk,filedialog,messagebox
        self.root=root;self.tk=tk;self.ttk=ttk;self.filedialog=filedialog;self.messagebox=messagebox
        self.shell=None;self.donor=None;self.busy=False;self.last_result=None
        root.title('WPS 模板样式工具 '+TOOL_VERSION+' 正式版');root.geometry('1080x830');root.minsize(900,720)
        style=ttk.Style(root);style.theme_use('vista' if 'vista' in style.theme_names() else 'clam')
        style.configure('Treeview',rowheight=25,font=('Microsoft YaHei UI',9))
        style.configure('TLabel',font=('Microsoft YaHei UI',10));style.configure('TButton',font=('Microsoft YaHei UI',10))
        box=ttk.Frame(root,padding=18);box.pack(fill='both',expand=True);box.columnconfigure(1,weight=1)
        ttk.Label(box,text='把报告样式放入现成的封面模板',font=('Microsoft YaHei UI',17,'bold')).grid(row=0,column=0,columnspan=3,sticky='w',pady=(0,8))
        ttk.Label(box,text='保留正文起点之前的封面、目录和页眉页脚；生成新的 DOCX，原件保持不变。').grid(row=1,column=0,columnspan=3,sticky='w',pady=(0,14))
        config=json.loads((resource_root()/'config.json').read_text('utf-8-sig'))
        self.builtin_donor=resource_root()/config['template'];self.builtin_label='内置母模板 v'+config['template_version']
        self.donor_path=tk.StringVar(value=self.builtin_label)
        self.shell_path=tk.StringVar();self.output_path=tk.StringVar()
        for row,label,var,command in [(2,'样式来源',self.donor_path,self.choose_donor),(3,'带封面的模板',self.shell_path,self.choose_shell),(4,'新模板保存到',self.output_path,self.choose_output)]:
            ttk.Label(box,text=label).grid(row=row,column=0,sticky='w',pady=5,padx=(0,12))
            ttk.Entry(box,textvariable=var).grid(row=row,column=1,sticky='ew',pady=5)
            ttk.Button(box,text='选择…',command=command).grid(row=row,column=2,padx=(10,0),pady=5)
        preferred=load_python_preference()
        self.runtime_mode=tk.StringVar(value='python' if preferred else 'embedded')
        self.python_path=tk.StringVar(value=preferred)
        ttk.Label(box,text='运行方式').grid(row=5,column=0,sticky='w',pady=5)
        runtime=ttk.Frame(box);runtime.grid(row=5,column=1,columnspan=2,sticky='ew',pady=5);runtime.columnconfigure(2,weight=1)
        ttk.Radiobutton(runtime,text='内置',variable=self.runtime_mode,value='embedded').grid(row=0,column=0,padx=(0,12))
        ttk.Radiobutton(runtime,text='指定 Python',variable=self.runtime_mode,value='python').grid(row=0,column=1,padx=(0,8))
        ttk.Entry(runtime,textvariable=self.python_path).grid(row=0,column=2,sticky='ew')
        ttk.Button(runtime,text='选择 Python…',command=self.choose_python).grid(row=0,column=3,padx=(8,0))
        actions=ttk.Frame(box);actions.grid(row=6,column=0,columnspan=3,sticky='ew',pady=(8,12))
        self.scan_button=ttk.Button(actions,text='读取样式与模板内容',command=self.inspect);self.scan_button.pack(side='left')
        ttk.Button(actions,text='导出读取诊断…',command=self.export_diagnostics).pack(side='left',padx=10)
        tabs=ttk.Notebook(box);tabs.grid(row=7,column=0,columnspan=3,sticky='nsew');box.rowconfigure(7,weight=1)
        body_tab=ttk.Frame(tabs,padding=10);styles_tab=ttk.Frame(tabs,padding=10)
        tabs.add(body_tab,text='选择正文起点');tabs.add(styles_tab,text='样式信息')
        ttk.Label(body_tab,text='选中第一个要替换的正文块。此前全部保留，从此处到文档结尾作为新的报告正文区域。',wraplength=950).pack(anchor='w',pady=(0,9))
        self.blocks=self.tree(body_tab,('index','type','text'),('位置','类型','内容预览'),(65,80,770))
        self.blocks.bind('<<TreeviewSelect>>',self.selection_changed)
        self.styles=self.tree(styles_tab,('name','fonts','size','paragraph'),('样式名称','字体','字号 pt','段落 / 编号'),(190,250,85,390))
        self.note=tk.StringVar(value='默认使用当前母模板 v0.3.11 的报告样式。选择带封面的文件后读取内容。')
        ttk.Label(box,textvariable=self.note,wraplength=1010,foreground='#425675').grid(row=8,column=0,columnspan=3,sticky='w',pady=(12,8))
        self.generate_button=ttk.Button(box,text='生成新模板',command=self.generate,state='disabled')
        self.generate_button.grid(row=9,column=2,sticky='e',pady=(3,0))
        ttk.Label(box,text='目录沿用原生目录域；使用生成的模板装配报告后，在 Word / WPS 中更新目录。',foreground='#667085').grid(row=9,column=0,columnspan=2,sticky='w')
        for variable in (self.shell_path,self.donor_path,self.runtime_mode,self.python_path):variable.trace_add('write',self.invalidate)

    def tree(self,parent,columns,labels,widths):
        frame=self.ttk.Frame(parent);frame.pack(fill='both',expand=True)
        tree=self.ttk.Treeview(frame,columns=columns,show='headings',selectmode='browse')
        bar=self.ttk.Scrollbar(frame,orient='vertical',command=tree.yview)
        horizontal=self.ttk.Scrollbar(frame,orient='horizontal',command=tree.xview)
        tree.configure(yscrollcommand=bar.set,xscrollcommand=horizontal.set)
        frame.columnconfigure(0,weight=1);frame.rowconfigure(0,weight=1)
        tree.grid(row=0,column=0,sticky='nsew');bar.grid(row=0,column=1,sticky='ns');horizontal.grid(row=1,column=0,sticky='ew')
        for key,label,width in zip(columns,labels,widths):tree.heading(key,text=label);tree.column(key,width=width,minwidth=55,stretch=key in ('text','paragraph','fonts'))
        return tree

    def invalidate(self,*_):
        self.shell=None;self.donor=None;self.generate_button.configure(state='disabled')
        self.note.set('路径或运行方式已改变，请重新读取样式与模板内容。')

    def choose_python(self):
        name=self.filedialog.askopenfilename(title='选择已能运行 template_tool.py 的 Python',filetypes=[('Python 解释器','python.exe pythonw.exe'),('可执行文件','*.exe')])
        if name:self.python_path.set(name);self.runtime_mode.set('python')

    def backend(self):
        python=self.python_path.get().strip() if self.runtime_mode.get()=='python' else None
        if self.runtime_mode.get()=='python' and not python:raise ValueError('请选择已能运行 template_tool.py 的 Python 解释器')
        try:save_python_preference(python)
        except OSError:self.note.set('本机运行偏好未能保存；本次仍使用所选运行方式。')
        return TemplateBackend(python)

    def selected_style_source(self):
        return None if self.donor_path.get()==self.builtin_label else self.donor_path.get()

    def checked_result(self,result):
        self.last_result=result
        if result.get('status')=='FAIL':raise ValueError(result['error'])
        return result

    def export_diagnostics(self):
        try:
            result=self.backend().diagnose(source=self.shell_path.get(),style_source=self.selected_style_source())
            if self.last_result and self.last_result.get('status')=='FAIL':result['last_failure']=self.last_result
            name=self.filedialog.asksaveasfilename(title='保存读取诊断',defaultextension='.json',initialfile='模板工具_读取诊断.json',filetypes=[('JSON 诊断','*.json')])
            if name:
                Path(name).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
                self.messagebox.showinfo('读取诊断已保存',name+'\n\n记录进程、文件长度、摘要和文件头，不保存文档正文。')
        except Exception as exc:self.messagebox.showerror('诊断失败',str(exc))

    def choose_donor(self):
        name=self.filedialog.askopenfilename(title='选择提供报告样式的模板',filetypes=[('Word 文档','*.docx')])
        if name:self.donor_path.set(name)
    def choose_shell(self):
        name=self.filedialog.askopenfilename(title='选择带封面、目录、页眉的模板',filetypes=[('Word 文档','*.docx')])
        if name:
            self.shell_path.set(name)
            path=Path(name);self.output_path.set(str(path.with_name(path.stem+'_报告母模板.docx')))
    def choose_output(self):
        name=self.filedialog.asksaveasfilename(title='保存新的报告模板',defaultextension='.docx',filetypes=[('Word 文档','*.docx')],confirmoverwrite=False)
        if name:self.output_path.set(name)

    def inspect(self):
        try:
            result=self.checked_result(self.backend().run('inspect-pair',source=self.shell_path.get(),style_source=self.selected_style_source()))
            donor=result['styles'];shell=result['document']
            self.blocks.delete(*self.blocks.get_children());self.styles.delete(*self.styles.get_children())
            for item in shell['blocks']:
                if item['type']=='sectPr':continue
                preview=('正文起点书签（普通空段落）' if 'WPS_ReportBody' in item['bookmarks'] else item['text']) or ('目录域：'+' '.join(item['fields']) if item['fields'] else '（空段落或非文字对象）')
                self.blocks.insert('', 'end',iid=str(item['index']),values=(item['index'],item['type'],preview))
            for row in donor['styles']:
                summary=paragraph_summary(row)
                self.styles.insert('','end',values=(row['name'],' / '.join(dict.fromkeys(x for x in (row['font_east_asia'],row['font_ascii']) if x)),row['size_pt'] if row['size_pt'] is not None else '继承/主题',summary))
            self.donor=donor;self.shell=shell
            if donor['missing_required_styles']:
                self.note.set('样式来源缺少：'+'、'.join(donor['missing_required_styles'])+'。可改用随工具提供的母模板。');return
            self.note.set('读取完成。选择正文起点，或选旧版的空正文框升级；生成后以普通空段落和书签定位，不添加固定题名。')
        except Exception as exc:
            self.shell=None;self.donor=None;self.generate_button.configure(state='disabled')
            self.messagebox.showerror('读取失败',str(exc)+'\n\n可导出读取诊断；若同一文件能由源码读取，请指定该 Python 后重新读取。')

    def selection_changed(self,*_):
        if not self.shell or not self.donor or self.donor['missing_required_styles']:return
        selected=self.blocks.selection()
        if not selected:return
        index=int(selected[0]);self.note.set(f'保留位置 0—{index-1} 的内容；从位置 {index} 起建立报告正文区域。请核对内容预览后生成。' if index else '从文档开头建立正文区域；没有封面正文需要保留。请核对后生成。')
        self.generate_button.configure(state='normal')

    def generate(self):
        if self.busy or not self.shell or not self.donor:return
        selected=self.blocks.selection()
        if not selected:return
        self.busy=True;self.generate_button.configure(state='disabled');self.root.update_idletasks()
        try:
            if not self.output_path.get().strip():raise ValueError('请选择新模板的保存位置')
            result=self.checked_result(self.backend().run('prepare',source=self.shell['source'],output=self.output_path.get(),
                body_start=int(selected[0]),expected_sha256=self.shell['sha256'],
                style_source=self.donor['source'],expected_style_sha256=self.donor['sha256']))
            self.note.set('已生成：'+result['output']+'。原件保持不变；该文件可直接作为报告母模板使用。')
            self.messagebox.showinfo('新模板已生成',result['output']+'\n\n封面和页眉保留，正文起点已用书签标记，没有正文大框或固定题名。装配报告后请更新目录。')
        except Exception as exc:self.messagebox.showerror('没有生成模板',str(exc))
        finally:self.busy=False;self.generate_button.configure(state='normal')


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    mode=parser.add_mutually_exclusive_group();mode.add_argument('--inspect',type=Path);mode.add_argument('--prepare',type=Path)
    mode.add_argument('--diagnose',type=Path,help='Read-only process/file diagnostics for a cover template')
    mode.add_argument('--worker-request',type=Path,help=argparse.SUPPRESS)
    parser.add_argument('--python',type=Path,help='Run document operations in this existing Python environment')
    parser.add_argument('--style-source',type=Path);parser.add_argument('--body-start',type=int)
    parser.add_argument('--expected-sha256');parser.add_argument('--expected-style-sha256')
    parser.add_argument('--output',type=Path);parser.add_argument('--result-json',type=Path)
    args=parser.parse_args(argv)
    if args.inspect or args.prepare or args.diagnose or args.worker_request:
        try:
            backend=TemplateBackend(args.python)
            if args.worker_request:
                if args.python:raise ValueError('A worker cannot select another Python worker')
                result=execute_request(json.loads(args.worker_request.read_text('utf-8-sig')))
            elif args.diagnose:result=backend.diagnose(source=str(args.diagnose),style_source=str(args.style_source) if args.style_source else None)
            elif args.inspect:result=backend.run('inspect',document=str(args.inspect))
            else:
                if args.output is None or args.body_start is None or not args.expected_sha256:raise ValueError('--prepare requires --output, --body-start and --expected-sha256')
                result=backend.run('prepare',source=str(args.prepare),output=str(args.output),body_start=args.body_start,
                    expected_sha256=args.expected_sha256,style_source=str(args.style_source) if args.style_source else None,
                    expected_style_sha256=args.expected_style_sha256)
            code=1 if result.get('status')=='FAIL' else 0
        except Exception as exc:result={'status':'FAIL','error':str(exc)};code=1
        if args.result_json:
            args.result_json.parent.mkdir(parents=True,exist_ok=True)
            with args.result_json.open('x',encoding='utf8') as stream:json.dump(result,stream,ensure_ascii=False,indent=2)
        if sys.stdout is not None:print(json.dumps(result,ensure_ascii=False,indent=2))
        return code
    import tkinter as tk
    if sys.platform=='win32':
        import ctypes
        try:ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except OSError:pass
    root=tk.Tk();TemplateWindow(root);root.mainloop();return 0

if __name__=='__main__':raise SystemExit(main())

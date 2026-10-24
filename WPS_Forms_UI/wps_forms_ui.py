"""Loopback UI adapter. XLSX and the original Forms CLI remain the inputs."""
import copy
import atexit
import hashlib
import io
import json
import os
import queue
import re
import secrets
import shutil
import stat
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from uuid import uuid4
from zipfile import ZipFile

VERSION = '0.1.2'
BUILD_ID = '20261024.ui-public-r1'
HERE = Path(__file__).resolve().parent
FORMS = {'02': ('核查', '02_扫描核查.xlsx'), '03': ('反馈', '03_反馈提交.xlsx')}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def now(): return datetime.now(timezone.utc).isoformat()
def digest(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def read(path): return json.loads(Path(path).read_text('utf-8-sig'))


def write(path, value):
    path = Path(path); path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name('.' + path.name + '.' + uuid4().hex + '.tmp')
    try:
        tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n', encoding='utf8')
        for attempt in range(5):
            try: os.replace(tmp, path); break
            except PermissionError:
                if attempt == 4: raise
                time.sleep(.05 * (attempt + 1))
    finally:
        if tmp.exists(): tmp.unlink()


def absolute(value):
    p = Path(str(value)).expanduser()
    if not p.is_absolute(): raise ValueError('请输入完整的绝对路径')
    return p.resolve()


class Conflict(Exception):
    def __init__(self, message, **data):
        super().__init__(message); self.data = data


class Application:
    def __init__(self, core, data):
        self.core, self.data = core, data
        if not (core / 'forms_tool.py').is_file(): raise ValueError('三表内核目录不存在：' + str(core))
        sys.path.insert(0, str(core / 'src'))
        from wps_pipeline import forms_workflow, form_xlsx, job_support
        from lxml import etree
        self.f, self.xlsx, self.j, self.E = forms_workflow, form_xlsx, job_support, etree
        if self.j.ROOT.resolve() != core.resolve(): raise ValueError('内核加载路径不一致')
        self.lock = threading.RLock(); self.native_lock = threading.Lock()
        self.data.mkdir(parents=True, exist_ok=True)
        self.state_file = data / 'workspace.json'
        self.state = read(self.state_file) if self.state_file.exists() else dict(version=VERSION, current=None, history=[])
        self.jobs = {}; self.active = None; self.native = None; self.native_lines = queue.Queue()
        self.template_process = None
        try:
            component = read(HERE / 'tools/template_style/COMPONENT_MANIFEST.json')
            self.template_state = dict(status='IDLE', available=True, version=component['tool_version'], python=sys.executable)
        except (OSError, ValueError, KeyError, TypeError):
            self.template_state = dict(status='UNAVAILABLE', available=False, message='模板工具组件不完整，请重新解压完整包。')
        self.log_file = data / 'ui_events.jsonl'
        self.log_lines = self.log_file.read_text('utf8').splitlines()[-1500:] if self.log_file.exists() else []
        self.event_offsets = {}
        self.note('服务', '已启动；文档与表单读写使用 ' + sys.executable)

    def note(self, stage, message, **extra):
        item = dict(time=now(), stage=stage, message=str(message), **extra)
        line = json.dumps(item, ensure_ascii=False)
        with self.lock:
            with self.log_file.open('a', encoding='utf8') as stream: stream.write(line + '\n')
            self.log_lines.append(line)
        return item

    def task(self, run=None):
        path = absolute(run or self.state.get('current') or '')
        run, state, inventory = self.f.load(path)
        core_run = self.j.contained(run, state['drafts'][-1]['run']) if state.get('drafts') else None
        core = read(core_run / 'state.json') if core_run else None
        return run, state, inventory, core_run, core

    def task_dir(self, run):
        key = hashlib.sha256(str(run).casefold().encode('utf8')).hexdigest()[:24]
        return self.data / 'tasks' / key

    def assert_idle(self):
        if self.active and self.jobs[self.active]['status'] == 'RUNNING':
            raise Conflict('当前批次仍在执行，请等待完成后再修改或提交表单。', job=self.active)

    def attach(self, run):
        self.assert_idle()
        run, state, inv, cr, core = self.task(run)
        with self.lock:
            self.state['current'] = str(run)
            old = next((x for x in self.state['history'] if x['run'] == str(run)), None)
            if old is None:
                self.state['history'].append(dict(run=str(run), created=state['created'], title=state['title']))
            write(self.state_file, self.state)
        self.note('任务', '当前稿切换到 ' + str(run))
        return self.summary()

    def summary(self):
        info = dict(ui_version=VERSION, ui_build=BUILD_ID, core=str(self.core), python=sys.executable,
                    forms_version=self.f.FORMS_VERSION, kernel_protocol=self.j.VERSION,
                    core_build=read(self.core / 'BUILD_INFO.json').get('build_id'),
                    data=str(self.data), current=self.state.get('current'), history=copy.deepcopy(self.state['history']),
                    job=copy.deepcopy(self.jobs.get(self.active)), log_count=len(self.log_lines),
                    template_tool=copy.deepcopy(self.template_state))
        if not self.state.get('current'): return info
        try:
            run, s, inv, cr, core = self.task()
            master = next((d for d in (core or {}).get('deliverables', []) if d['id'] == 'master'), None)
            master_path = self.j.contained(cr, master['file']) if master else None
            info.update(title=s['title'], status=s['status'], sources=len(inv['sources']),
                        pending=len(read(run / 'feedback_export.json')['exports']), run_id=s['run_id'],
                        master=str(master_path) if master_path and master_path.is_file() else None,
                        master_sha256=master['sha256'] if master else None,
                        gaps=len((master or {}).get('missing_source_ids', [])),
                        local_media_gaps=len((master or {}).get('local_media_gaps', [])))
            confirmed = self.task_dir(run) / 'final_confirmation.json'
            info['final'] = read(confirmed) if confirmed.exists() else None
            if info['final'] and not self.final_current(run, info['final']): info['final'] = None
        except Exception as exc:
            info['task_error'] = str(exc)
        return info

    def open_template_tool(self):
        with self.lock:
            if self.template_process is not None and self.template_process.poll() is None:
                return dict(status='ALREADY_RUNNING', message='模板样式工具正在启动或已打开，请从任务栏切换到该窗口。')
            folder = HERE / 'tools/template_style'
            manifest = read(folder / 'COMPONENT_MANIFEST.json')
            names = set()
            for item in manifest['files']:
                path = (folder / item['path']).resolve()
                if not path.is_relative_to(folder.resolve()) or item['path'] in names:
                    raise ValueError('模板工具组件清单路径无效')
                names.add(item['path'])
                if not path.is_file() or path.stat().st_size != item['bytes'] or digest(path) != item['sha256']:
                    raise ValueError('模板工具组件不完整：' + item['path'] + '；请重新解压完整包。')
            if 'template_tool.py' not in names: raise ValueError('模板工具入口未登记')
            if {p.relative_to(folder).as_posix() for p in folder.rglob('*') if p.is_file()} != names | {'COMPONENT_MANIFEST.json'}:
                raise ValueError('模板工具目录与清单不一致，请使用完整原组件。')
            ident = uuid4().hex
            log = self.data / 'tools' / ident / 'template_tool.log'
            log.parent.mkdir(parents=True)
            environment = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONDONTWRITEBYTECODE='1',
                               WPS_TEMPLATE_PYTHON=sys.executable)
            command = [sys.executable, '-u', '-B', str(HERE / 'ui_tool.py'), 'template-tool']
            state = dict(id=ident, status='STARTING', available=True, version=manifest['tool_version'],
                         python=sys.executable, log=str(log))
            with log.open('w', encoding='utf8') as output:
                process = subprocess.Popen(command, cwd=HERE, env=environment, stdout=output,
                                           stderr=subprocess.STDOUT,
                                           creationflags=0x08000000 if os.name == 'nt' else 0)
            state['pid'] = process.pid
            self.template_process = process; self.template_state = state
            write(log.parent / 'launch.json', dict(command=command, component_sha256=digest(folder / 'COMPONENT_MANIFEST.json'),
                                                  created=now(), **state))
            self.note('模板工具', '正在用当前Python启动独立模板样式窗口；生成后在01选择新模板。', log=str(log))
            threading.Thread(target=self.watch_template_tool, args=(process, state, log), daemon=True).start()
            return copy.deepcopy(state)

    def watch_template_tool(self, process, state, log):
        ready = False
        while True:
            code = process.poll()
            if not ready:
                for line in log.read_text('utf8', errors='replace').splitlines():
                    try: event = json.loads(line)
                    except ValueError: continue
                    if isinstance(event, dict) and event.get('event') == 'TEMPLATE_TOOL_READY':
                        with self.lock: state['status'] = 'RUNNING'
                        ready = True
                        self.note('模板工具', '模板样式工具 ' + str(state['version']) + ' 窗口已就绪。', log=str(log))
                        break
            if code is not None: break
            time.sleep(.25)
        output = log.read_text('utf8', errors='replace')
        with self.lock:
            state.update(status='CLOSED' if code == 0 and ready else 'ERROR', returncode=code,
                         message='模板工具已关闭。' if code == 0 and ready else '模板工具未能正常运行；查看模板工具日志。若缺少tkinter，请在当前Conda环境安装tk。')
            write(log.parent / 'result.json', dict(**state, finished=now(), ready=ready, output_tail=output[-4000:]))
        self.note('模板工具' if state['status'] == 'CLOSED' else '错误', state['message'], log=str(log), returncode=code)

    def form(self, kind, run=None):
        if kind not in FORMS: raise ValueError('表单类型无效')
        run, state, inv, cr, core = self.task(run)
        name, filename = FORMS[kind]
        observed = digest(run / filename)
        cells, binding = self.xlsx.read_form(run / filename, name)
        if digest(run / filename) != observed: raise Conflict('表单正在被外部保存，请完成保存后重新同步。')
        cells = {k: str(v) for k, v in cells.items()}
        path = self.task_dir(run) / (kind + '_working.json')
        if path.exists():
            work = read(path)
            if work['binding'] != binding: raise ValueError('表单绑定已改变，请保留现场')
        else:
            work = dict(version=1, run=str(run), kind=kind, binding=binding,
                        base=cells, cells=cells, original=cells, disk_sha256=digest(run / filename))
            write(path, work)
        work['_observed_disk_sha256'] = observed
        return run, work, cells, path

    @staticmethod
    def editable(kind, address, work):
        if address == 'B6': return False  # Submission confirmation is set on the frozen batch only.
        match = re.fullmatch(r'([A-I])([1-9][0-9]*)', address)
        if not match or int(match[2]) < 9 or address not in work['base']: return False
        if kind == '02': return match[1] in 'EFGHI'
        row = match[2]
        return match[1] in 'DEFGH' or (not work['base'].get('A' + row) and match[1] in 'BC')

    def form_data(self, kind):
        run, work, disk, path = self.form(kind)
        cells = work['cells']; rows = []
        numbers = sorted({int(re.search(r'\d+', k)[0]) for k in work['base'] if int(re.search(r'\d+', k)[0]) >= 9})
        for n in numbers:
            values = {col: cells.get(col + str(n), '') for col in 'ABCDEFGHI'}
            rows.append(dict(row=n, values=values,
                             editable=[col for col in 'ABCDEFGHI' if self.editable(kind, col + str(n), work)]))
        dirty = any(cells.get(k, '') != work['original'].get(k, '') for k in cells if self.editable(kind, k, work))
        return dict(kind=kind, version=work['version'], rows=rows, dirty=dirty,
                    external_changed=digest(run / FORMS[kind][1]) != work['disk_sha256'],
                    path=str(run / FORMS[kind][1]), title=cells.get('A1', ''), run=str(run))

    def sync(self, kind, version, changes=None, resolutions=None, import_path=None, origin='user'):
        self.assert_idle()
        if origin not in ('user', 'model'): raise ValueError('修改来源须为user或model')
        run, work, disk, path = self.form(kind)
        if version != work['version']: raise Conflict('界面记录已被其他窗口更新，请重新载入后核对。', stale=True)
        local = dict(work['cells'])
        for key, value in (changes or {}).items():
            if not self.editable(kind, key, work): raise ValueError('不能修改程序事实单元格：' + key)
            if not isinstance(value, (str, int, float)) or len(str(value)) > 12000: raise ValueError('表单值无效')
            local[key] = str(value)
        # Save UI edits before reporting a conflict so a page reload cannot lose them.
        work['cells'] = local
        work['version'] += 1
        write(path, work)
        merged = dict(disk); conflicts = []; resolutions = resolutions or {}
        remote = disk; base = work['base']
        if import_path:
            imported, meta = self.xlsx.read_form(absolute(import_path), FORMS[kind][0])
            if meta != work['binding']: raise ValueError('导入表不属于当前稿件或绑定被修改')
            imported = {k: str(v) for k, v in imported.items()}
            with ZipFile(import_path) as archive:
                if 'wps-ui-export.json' in archive.namelist():
                    tag = json.loads(archive.read('wps-ui-export.json'))
                    record = self.data / 'exports' / (str(tag.get('id', '')) + '.json')
                    if not re.fullmatch(r'[a-f0-9]{32}', str(tag.get('id', ''))) or not record.is_file():
                        raise ValueError('导出副本基线缺失，请使用当前共享表单')
                    previous = read(record)
                    if previous['binding'] != work['binding']: raise ValueError('导出副本任务不一致')
                    base = previous['cells']
            # Include changes already made to the canonical file in the local side.
            for key in disk:
                if self.editable(kind, key, work) and local.get(key, '') == work['base'].get(key, ''):
                    local[key] = disk[key]
            remote = imported
        for key in work['base']:
            if not self.editable(kind, key, work):
                if remote.get(key, '') != work['base'].get(key, ''):
                    # B6 is an external submit switch, not an immutable source fact.
                    if key != 'B6': raise ValueError('程序事实或表头被修改：' + key)
                continue
            a, b, c = base.get(key, ''), local.get(key, ''), remote.get(key, '')
            if b != a and c != a and b != c:
                if key in resolutions and resolutions[key] in ('local', 'remote'):
                    merged[key] = b if resolutions[key] == 'local' else c
                else: conflicts.append(dict(cell=key, base=a, local=b, remote=c))
            else: merged[key] = c if b == a else b
        if conflicts:
            raise Conflict('界面与Excel修改了相同字段，请明确选择保留值。', conflicts=conflicts, version=work['version'])
        canonical = run / FORMS[kind][1]
        expected = work['_observed_disk_sha256']
        self.patch_xlsx(canonical, merged, expected)
        actual, binding = self.xlsx.read_form(canonical, FORMS[kind][0])
        actual = {k: str(v) for k, v in actual.items()}
        if binding != work['binding']: raise ValueError('同步后绑定不符')
        work.update(cells=actual, base=actual, disk_sha256=digest(canonical), version=work['version'] + 1,
                    origin=origin, synchronized=now())
        write(path, work)
        # A former final confirmation stays in history, but is not valid for an edited form.
        self.note('同步', FORMS[kind][1] + ' 与界面已同步', run=str(run), version=work['version'])
        return self.form_data(kind)

    def patch_xlsx(self, path, values, expected):
        with ZipFile(path) as z:
            members = [(i, z.read(i)) for i in z.infolist()]
        parser = self.E.XMLParser(resolve_entities=False, no_network=True)
        sheet = self.E.fromstring(dict((i.filename, b) for i, b in members)['xl/worksheets/sheet1.xml'], parser)
        for cell in sheet.findall('.//{' + S + '}c'):
            address = cell.get('r')
            if address not in values: continue
            for child in list(cell): cell.remove(child)
            cell.set('t', 'inlineStr')
            text = self.E.SubElement(self.E.SubElement(cell, '{' + S + '}is'), '{' + S + '}t')
            text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); text.text = str(values[address])
        payload = self.E.tostring(sheet, encoding='UTF-8', xml_declaration=True, standalone=True)
        tmp = path.with_name('.' + path.name + '.' + uuid4().hex + '.tmp')
        try:
            with ZipFile(tmp, 'x') as z:
                for info, content in members: z.writestr(info, payload if info.filename == 'xl/worksheets/sheet1.xml' else content)
            if digest(path) != expected: raise Conflict('Excel在同步期间再次改变，本次未替换，请重新同步。')
            os.replace(tmp, path)
        except PermissionError as exc:
            raise ValueError('表单被WPS/Excel占用；请保存并关闭该表再同步。') from exc
        finally:
            if tmp.exists(): tmp.unlink()

    def exported_form(self, kind):
        run, work, disk, path = self.form(kind)
        if work['cells'] != work['base'] or digest(run / FORMS[kind][1]) != work['disk_sha256']:
            raise ValueError('请先同步界面和共享表单再导出')
        ident = uuid4().hex
        write(self.data / 'exports' / (ident + '.json'), dict(binding=work['binding'], cells=disk))
        out = io.BytesIO()
        with ZipFile(run / FORMS[kind][1]) as src, ZipFile(out, 'w') as dest:
            for info in src.infolist():
                if info.filename != 'wps-ui-export.json': dest.writestr(info, src.read(info))
            dest.writestr('wps-ui-export.json', json.dumps(dict(id=ident)))
        return out.getvalue()

    def core_document(self, cr, core, document):
        record = next((d for d in core['deliverables'] if d['id'] == document), None)
        if not record: raise ValueError('当前稿没有该文档')
        path = self.j.contained(cr, record['file'])
        if digest(path) != record['sha256']: raise ValueError('稿件已在外部修改，保留手工稿；不能继续使用旧定位')
        index = self.j.contained(cr, record['locator_index'])
        if digest(index) != record['locator_index_sha256']: raise ValueError('定位索引改变')
        return path, record, read(index)

    def body(self, path):
        with ZipFile(path) as z:
            root = self.E.fromstring(z.read('word/document.xml'), self.E.XMLParser(resolve_entities=False, no_network=True))
        return root.find('{' + W + '}body')

    def node_text(self, node):
        return ''.join(node.xpath('.//w:t/text()|.//m:t/text()', namespaces={'w': W, 'm': M}))

    def snippet(self, path, index):
        body = self.body(path)
        if type(index) is not int or not 0 <= index < len(body): raise ValueError('块位置不在文档内')
        return dict(before=[self.node_text(n)[:350] for n in body[max(0, index-2):index]],
                    text=self.node_text(body[index]), after=[self.node_text(n)[:350] for n in body[index+1:index+3]],
                    body_index=index)

    def context(self, row):
        run, s, inv, cr, core = self.task()
        _, work, _, _ = self.form('03')
        values = work['cells']; number = str(int(row))
        doc, _, locator = values.get('B' + number, '').partition('|'); doc = doc.strip() or 'master'; locator = locator.strip()
        quote = values.get('C' + number, '')
        output, record, index = self.core_document(cr, core, doc)
        matches = [r for r in index if (not locator or r['locator'] == locator) and (not quote or locator or quote in r['text'])]
        if len(matches) != 1:
            sid = locator.partition(':')[0]
            plan = read(run / 'approved_plan.json')
            unit = next((d for d in plan.get('derivations', []) if d['id'] == sid), None)
            original = next((x for x in inv['sources'] if x['id'] == (unit['source_id'] if unit else sid)), None)
            source = None
            if original:
                frozen = self.j.contained(run, original['frozen'])
                source = dict(path=str(frozen), title=original.get('true_title') or original['name'],
                              relative=original['relative'], sha256=original['sha256'], chapter='块范围未唯一确定')
            return dict(status='UNLOCATED', message='该事项尚无唯一块定位；可打开当前总稿核对，不能猜选一个位置。',
                        source=source, target=dict(path=str(output), title=doc, document=doc), candidates=len(matches))
        hit = matches[0]; sid = hit['source_id']; bid = hit['block_id']
        unit = next(x for x in core['sources'] if x['id'] == sid)
        candidate = read(cr / 'sources' / sid / 'candidate.json')
        block = next(x for x in candidate['blocks'] if x['block_id'] == bid)
        unit_index = block['source_locator']['body_child_index']
        plan = read(run / 'approved_plan.json')
        derivation = next((d for d in plan.get('derivations', []) if d['id'] == sid), None)
        original_id, original_index = sid, unit_index
        if derivation:
            mapped = [m for m in derivation['mapping'] if m['unit_index'] == unit_index]
            if len(mapped) > 1 and unit_index == 0:
                mapped = [m for m in mapped if m['operation'] == 'root_to_title'] or mapped
            if len(mapped) != 1: raise ValueError('拆分单元不能唯一回到原稿块，保留章节定位')
            original_id, original_index = derivation['source_id'], mapped[0]['source_index']
        original = next(x for x in inv['sources'] if x['id'] == original_id)
        source_path = self.j.contained(run, original['frozen'])
        if digest(source_path) != original['sha256']: raise ValueError('冻结来源改变')
        source_candidate = read(run / 'source_scan/sources' / original_id / 'candidate.json')
        chapter = []
        for b in source_candidate['blocks']:
            if b['source_locator']['body_child_index'] > original_index: break
            if b['role'] == 'heading' and b.get('status') != 'review':
                level = b.get('level') or 1; chapter = [x for x in chapter if x[0] < level] + [(level, b.get('title') or b['block_id'])]
        target_chapter = []
        for r in index:
            if r['output_index'] > hit['output_index']: break
            if r['source_id'] != sid: continue
            if r['role'] in ('heading', 'document_title'):
                if r['role'] == 'document_title': target_chapter = [(0, r['text'])]
                else:
                    level = r.get('level') or 1
                    target_chapter = [x for x in target_chapter if x[0] < level] + [(level, r['text'])]
        src = dict(path=str(source_path), title=original.get('true_title') or original['name'],
                   relative=original['relative'], sha256=original['sha256'], source_id=original_id,
                   chapter=' → '.join(x[1] for x in chapter), **self.snippet(source_path, original_index))
        entry = next((e for e in plan['entries'] if e['id'] == sid), {})
        target_path = (entry.get('group', []) if doc == 'master' else []) + [x[1] for x in target_chapter]
        dst = dict(path=str(output), title=s['title'], document=doc, sha256=record['sha256'],
                   chapter=' → '.join(target_path), role=hit['role'], level=hit.get('level'),
                   **self.snippet(output, hit['output_index']))
        return dict(status='LOCATED', source=src, target=dst, locator=hit['locator'], run=str(run), row=int(row))

    def view_copy(self, path, index, title):
        key = hashlib.sha256(('readonly-v2:' + digest(path) + ':' + str(index)).encode('utf8')).hexdigest()[:20]
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', title)[:55]
        folder = self.data / 'views' / key; folder.mkdir(parents=True, exist_ok=True)
        dest = folder / (safe + '_只读对照.docx'); bookmark = 'WPS_UI_' + key
        if dest.exists():
            receipt = read(folder / 'view.json')
            if digest(dest) != receipt['view_sha256']: raise ValueError('只读查看副本被修改，请保留该副本并另开任务')
            os.chmod(dest, stat.S_IREAD)
            return dict(path=str(dest), bookmark=bookmark if index is not None else None)
        with ZipFile(path) as z: members = [(i, z.read(i)) for i in z.infolist()]
        payload = None
        if index is not None:
            root = self.E.fromstring(dict((i.filename, b) for i, b in members)['word/document.xml'])
            body = root.find('{' + W + '}body')
            if not 0 <= index < len(body): raise ValueError('查看位置超出当前文档')
            identifiers = [int(n.get('{' + W + '}id')) for n in root.iter('{' + W + '}bookmarkStart') if str(n.get('{' + W + '}id')).isdigit()]
            ident = str(max(identifiers + [0]) + 1)
            first = self.E.Element('{' + W + '}bookmarkStart', {'{' + W + '}id': ident, '{' + W + '}name': bookmark})
            last = self.E.Element('{' + W + '}bookmarkEnd', {'{' + W + '}id': ident})
            body.insert(index, first); body.insert(index + 2, last)
            payload = self.E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)
        with ZipFile(dest, 'x') as z:
            for info, content in members: z.writestr(info, payload if info.filename == 'word/document.xml' and payload is not None else content)
        write(folder / 'view.json', dict(source=str(path), source_sha256=digest(path), body_index=index,
                                        view_sha256=digest(dest), bookmark=bookmark, created=now()))
        os.chmod(dest, stat.S_IREAD)
        return dict(path=str(dest), bookmark=bookmark if index is not None else None)

    def native_request(self, documents):
        with self.native_lock:
            if self.native is None or self.native.poll() is not None:
                self.native_lines = queue.Queue()
                self.native = subprocess.Popen([sys.executable, '-B', str(HERE / 'wps_forms_native.py')],
                                               stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                                               text=True, encoding='utf8', env=dict(os.environ, PYTHONIOENCODING='utf-8'),
                                               creationflags=0x08000000 if os.name == 'nt' else 0)
                process = self.native
                def collect():
                    for line in process.stdout: self.native_lines.put(line)
                threading.Thread(target=collect, daemon=True).start()
            self.native.stdin.write(json.dumps(dict(documents=documents), ensure_ascii=True) + '\n'); self.native.stdin.flush()
            try: reply = json.loads(self.native_lines.get(timeout=40))
            except queue.Empty:
                self.native.terminate(); self.native = None
                raise ValueError('WPS打开定位超时；未自动重试，已生成的总稿保留。')
            if reply['status'] == 'ERROR': raise ValueError('WPS定位失败：' + reply['message'] + '；可用“仅打开”由系统关联程序查看。')
            return reply

    def open_documents(self, row=None, side='both', plain=False):
        if row is None:
            run, s, inv, cr, core = self.task()
            p, d, _ = self.core_document(cr, core, 'master')
            items = [self.view_copy(p, None, '总稿_' + run.name + '_' + s['title'])]
        else:
            context = self.context(row); items = []
            for key in ('source', 'target'):
                if side not in ('both', key): continue
                item = context.get(key)
                if not item: continue
                label = ('原稿_' + item.get('source_id', '') if key == 'source' else '总稿_' + Path(self.state['current']).name)
                label += '_块' + str(item.get('body_index', '未定位')) + '_' + item['title']
                items.append(self.view_copy(Path(item['path']), item.get('body_index'), label))
            if not items: raise ValueError('此事项没有可打开的文档')
        if plain:
            if os.name != 'nt': raise ValueError('系统关联打开仅支持Windows')
            for item in items: os.startfile(item['path'])
            result = dict(status='OPENED_WITHOUT_POSITION', documents=items, message='仅打开只读对照副本；未执行原生定位')
        else: result = self.native_request(items)
        self.note('Word', '打开原稿/总稿：' + result['status'], result=result)
        return result

    def start_job(self, operation, payload):
        with self.lock:
            self.assert_idle()
            if operation not in ('scan', 'classify', 'feedback'): raise ValueError('操作无效')
            parent = None
            if operation == 'scan':
                root = absolute(payload.get('root', ''))
                if not root.is_dir(): raise ValueError('接入根目录不存在')
                title = str(payload.get('title') or '工程报告汇总').strip()
                if len(title) > 200: raise ValueError('总稿名过长')
                dest = absolute(payload['out']) if payload.get('out') else self.data / 'runs' / ('任务_' + time.strftime('%Y%m%d_%H%M%S') + '_' + uuid4().hex[:4]) / 'R01_首稿'
            else:
                parent, s, inv, cr, core = self.task()
                if payload.get('origin', 'user') not in ('user', 'model'): raise ValueError('提交来源须为user或model')
                kind = '02' if operation == 'classify' else '03'
                _, work, disk, _ = self.form(kind)
                if payload.get('version') != work['version'] or digest(parent / FORMS[kind][1]) != work['disk_sha256']:
                    raise Conflict('提交依据已经改变，请先同步最新表单。')
                if work['cells'] != work['base']: raise ValueError('界面还有未同步修改')
                other = '03' if kind == '02' else '02'
                _, otherwork, otherdisk, _ = self.form(other)
                pending = any(otherwork['cells'].get(k, '') != otherwork['original'].get(k, '')
                              for k in otherwork['cells'] if self.editable(other, k, otherwork))
                external_pending = any(otherdisk.get(k, '') != otherwork['base'].get(k, '')
                                       for k in otherdisk if self.editable(other, k, otherwork))
                if pending or external_pending:
                    raise ValueError('另一张表也有尚未提交的修改；请先提交那一批或明确撤销，防止换稿后遗漏。')
                suffix = '分类' if operation == 'classify' else '反馈'
                dest = absolute(payload['out']) if payload.get('out') else parent.parent / ('R' + str(len(self.state['history']) + 1).zfill(2) + '_' + suffix + '_' + uuid4().hex[:4])
            if dest.exists(): raise ValueError('输出必须为全新目录，不能覆盖已有稿件')
            if parent and (dest == parent or dest.is_relative_to(parent)): raise ValueError('新稿目录不能在旧任务内部')
            ident = uuid4().hex
            folder = self.data / 'jobs' / ident; folder.mkdir(parents=True)
            job = dict(id=ident, operation=operation, status='RUNNING', started=now(), output=str(dest),
                       parent=str(parent) if parent else None, stop_requested=False, phase='准备输入', returncode=None)
            if parent:
                canonical = parent / FORMS[kind][1]
                snapshot = folder / FORMS[kind][1]
                shutil.copy2(canonical, snapshot)
                if digest(snapshot) != work['disk_sha256']: raise Conflict('冻结提交时表单改变，请重新同步。')
                self.patch_xlsx(snapshot, {'B6': '确认本表' if kind == '02' else '提交本表'}, digest(snapshot))
                job.update(input=str(snapshot), input_sha256=digest(snapshot), form_version=work['version'],
                           origin=payload.get('origin', 'user'))
            else:
                input_path = folder / '01_输入.xlsx'
                self.f.new_input(input_path, root=str(root), template=str(payload.get('template') or ''), out=str(dest),
                                 title=title, policy='人工核查', pages='wps' if payload.get('pages') else 'off',
                                 doc_conversion='wps' if payload.get('doc_conversion', True) else 'off')
                job.update(input=str(input_path), input_sha256=digest(input_path))
            self.jobs[ident] = job; self.active = ident
            write(folder / 'job.json', job)
            self.note('提交', operation + '：输入已冻结，一次提交生成一份新总稿', output=str(dest), job=ident)
            threading.Thread(target=self.run_job, args=(ident,), daemon=True).start()
            return copy.deepcopy(job)

    def poll_events(self, dest):
        for path in (dest / 'source_scan/events.jsonl', dest / 'draft_001/events.jsonl'):
            if not path.exists(): continue
            offset = self.event_offsets.get(str(path), 0)
            with path.open('rb') as stream:
                stream.seek(offset); content = stream.read()
            end = content.rfind(b'\n') + 1
            if not end: continue
            self.event_offsets[str(path)] = offset + end
            for line in content[:end].splitlines():
                try: item = json.loads(line)
                except Exception: continue
                stage = item.get('stage', '内核')
                self.jobs[self.active]['phase'] = stage + ' · ' + item.get('status', '')
                self.note('内核', stage + ' · ' + item.get('status', ''), detail=item)

    def run_job(self, ident):
        job = self.jobs[ident]; folder = self.data / 'jobs' / ident; dest = Path(job['output'])
        if job['operation'] == 'scan': args = ['scan', job['input'], '--run', str(dest)]
        else: args = [job['operation'], job['parent'], '--form', job['input'], '--out', str(dest), '--origin', job['origin']]
        command = [sys.executable, '-B', str(self.core / 'forms_tool.py')] + args
        write(folder / 'command.json', dict(argv=command, shell=False, cwd=str(self.core)))
        try:
            environment = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUNBUFFERED='1', PYTHONDONTWRITEBYTECODE='1')
            process = subprocess.Popen(command, cwd=self.core, env=environment, stdout=subprocess.PIPE,
                                       stderr=subprocess.STDOUT, text=True, encoding='utf8', errors='replace',
                                       creationflags=0x08000000 if os.name == 'nt' else 0)
            chunks = []
            def collect():
                with (folder / 'program_output.log').open('w', encoding='utf8') as output:
                    for line in process.stdout:
                        chunks.append(line); output.write(line); output.flush()
                        self.note('程序输出', line.rstrip())
            reader = threading.Thread(target=collect, daemon=True); reader.start()
            stopped = False
            while process.poll() is None:
                self.poll_events(dest)
                if job['stop_requested']:
                    process.terminate(); stopped = True
                    break
                time.sleep(.25)
            try: process.wait(timeout=10)
            except subprocess.TimeoutExpired: process.kill(); process.wait(timeout=5)
            reader.join(timeout=5); self.poll_events(dest)
            job['returncode'] = process.returncode
            try: result = json.loads(''.join(chunks))
            except Exception: result = dict(status='UNPARSED', message='请查看程序输出原文')
            write(folder / 'result.json', result)
            if stopped:
                job.update(status='STOPPED', message='本批执行已停止；旧稿保留，未将中间产物当成完整新稿。')
            elif (dest / 'forms_state.json').is_file():
                run, s, inv, cr, core = self.task(dest)
                masters = [d for d in (core or {}).get('deliverables', []) if d['id'] == 'master']
                if masters:
                    job.update(status='DELIVERED' if process.returncode == 0 else 'DELIVERED_WITH_ISSUES',
                               phase='已交付；请查看剩余事项', result_status=s['status'])
                    # Set non-running before attaching; a new task must not race an active job.
                    self.attach(dest)
                else: job.update(status='ERROR', message='本批没有生成新总稿，旧稿保留。')
            else: job.update(status='ERROR', message=result.get('message') or '执行未完成，旧稿保留。')
        except Exception as exc:
            job.update(status='ERROR', message=str(exc))
            write(folder / 'failure.json', dict(message=str(exc), traceback=traceback.format_exc()))
        job['finished'] = now(); write(folder / 'job.json', job)
        self.note('交付', job['status'] + ' · ' + job.get('message', job.get('result_status', '')), output=job['output'])

    def final_current(self, run, confirmation):
        try:
            _, s, inv, cr, core = self.task(run)
            master, d, _ = self.core_document(cr, core, 'master')
            if confirmation['master_sha256'] != digest(master): return False
            if confirmation['forms_state_sha256'] != digest(run / 'forms_state.json'): return False
            for kind in FORMS:
                _, work, disk, _ = self.form(kind, run)
                if confirmation['forms'][kind] != digest(run / FORMS[kind][1]): return False
                if work['cells'] != work['base']: return False
                if any(work['cells'].get(k, '') != work['original'].get(k, '') for k in work['cells'] if self.editable(kind, k, work)):
                    return False
            return True
        except Exception: return False

    def finalize(self, accept_pending, origin='user'):
        self.assert_idle()
        if origin not in ('user', 'model'): raise ValueError('终稿确认来源无效')
        run, s, inv, cr, core = self.task()
        checked = self.f.forms_status(run)
        latest = checked['latest']
        if latest.get('receipt_integrity') != 'PASS' or s['status'] == 'NEEDS_REPAIR':
            raise ValueError('本稿检查尚有执行/内容缺口，不能标记为最终完成；已有稿仍可查看。')
        exports = read(run / 'feedback_export.json')['exports']
        if exports and not accept_pending: raise ValueError('仍有待核项，请明确接受保留，或继续修订。')
        _, master, _ = self.core_document(cr, core, 'master')
        forms = {}
        for kind in FORMS:
            _, work, disk, _ = self.form(kind)
            if work['cells'] != work['base'] or digest(run / FORMS[kind][1]) != work['disk_sha256']:
                raise ValueError('请先同步两张表，再确认终稿')
            if any(work['cells'].get(k, '') != work['original'].get(k, '') for k in work['cells'] if self.editable(kind, k, work)):
                raise ValueError('还有已填写但未生成新稿的决定，请先提交，或明确撤销。')
            forms[kind] = digest(run / FORMS[kind][1])
        confirmation = dict(version=VERSION, run_id=s['run_id'], run=str(run), created=now(),
                            master_sha256=master['sha256'], forms_state_sha256=digest(run / 'forms_state.json'),
                            forms=forms, pending_accepted=len(exports), automatic_items_are_truth=False,
                            origin=origin, human_confirmed=origin == 'user', checks=latest['receipt_integrity'])
        write(self.task_dir(run) / 'final_confirmation.json', confirmation)
        self.note('终稿', ('人工' if origin == 'user' else 'Agent') + '确认当前文档已完成最终修正；自动项/保留项不因此转为训练真值。')
        return confirmation

    def reset_form(self, kind, version):
        self.assert_idle()
        run, work, disk, path = self.form(kind)
        if version != work['version']: raise Conflict('工作记录已更新，请重新载入。')
        self.patch_xlsx(run / FORMS[kind][1], work['original'], work['_observed_disk_sha256'])
        work.update(cells=work['original'], base=work['original'], disk_sha256=digest(run / FORMS[kind][1]), version=work['version'] + 1)
        write(path, work); self.note('撤销', kind + '未提交修改已明确撤销')
        return self.form_data(kind)

    def training_export(self):
        self.assert_idle()
        run, s, inv, cr, core = self.task()
        confirmation_path = self.task_dir(run) / 'final_confirmation.json'
        if not confirmation_path.exists() or not self.final_current(run, read(confirmation_path)):
            raise ValueError('请先确认当前稿最终完成；新修改或新稿会使确认失效。')
        export = self.data / 'training' / (time.strftime('%Y%m%d_%H%M%S') + '_' + s['run_id'][:8])
        export.mkdir(parents=True, exist_ok=False)
        # Collapse cumulative ledger entries by stable ID. Only explicit, checked changes are labels.
        history = {x['id']: x for x in s.get('feedback_history', [])}
        labels, model_labels, retained = [], [], []
        for item in history.values():
            target = (labels if item.get('origin') == 'user' else model_labels if item.get('origin') == 'model' else retained) if item['status'] == 'closed_in_checked_draft' else retained
            row = dict(schema='wps-ui-training/v1', origin=item.get('origin', 'unknown'), evidence=item,
                       final_output_sha256=read(confirmation_path)['master_sha256'],
                       label_status='human_checked_correction' if target is labels else 'model_checked_candidate' if target is model_labels else 'not_a_label')
            target.append(row)
        for name, rows in [('explicit_corrections.jsonl', labels), ('model_review_candidates.jsonl', model_labels), ('retained_and_history.jsonl', retained)]:
            (export / name).write_text(''.join(json.dumps(x, ensure_ascii=False) + '\n' for x in rows), encoding='utf8')
        raw = export / 'sources'; raw.mkdir()
        source_records = []
        for source in inv['sources']:
            src = self.j.contained(run, source['frozen']); sha = digest(src)
            dest = raw / (sha + src.suffix)
            if not dest.exists(): shutil.copy2(src, dest)
            if digest(dest) != sha: raise ValueError('训练原件复制校验失败')
            source_records.append(dict(source_id=source['id'], original_relative=source['relative'], sha256=sha,
                                       file=dest.relative_to(export).as_posix(), split_group=sha))
        for name in ('inventory.json', 'approved_plan.json', 'feedback_ledger.json'):
            if (run / name).exists(): shutil.copy2(run / name, export / name)
        for source in core['sources']:
            target = export / 'evidence' / source['id']; target.mkdir(parents=True)
            for name in ('scan/scan.json', 'candidate.json', 'decisions.json'):
                original = cr / 'sources' / source['id'] / name
                if original.is_file(): shutil.copy2(original, target / Path(name).name)
        shutil.copy2(confirmation_path, export / 'final_confirmation.json')
        write(export / 'sources.json', source_records)
        (export / 'README.md').write_text('# 本轮修正数据\n\n人工明确且已检查的修正保存在explicit_corrections，模型判断单列为候选。保留项、自动规则和整稿终审不自动成为结构真值。原件按SHA去重保存；同母本及格式变体需要人工归入同一个split_group后再划分训练/验证集。组织/字号等偏好须按用途筛选。本导出不训练网络，也未生成训练准确率。\n', encoding='utf8')
        write(export / 'manifest.json', dict(schema='wps-ui-training/v1', run=str(run), created=now(),
                                             explicit_corrections=len(labels), model_candidates=len(model_labels), not_labels=len(retained),
                                             program_model_api_calls=0, agent_tokens=None,
                                             files=[dict(path=p.relative_to(export).as_posix(), sha256=digest(p)) for p in export.rglob('*') if p.is_file()]))
        self.note('训练数据', '已保存人工修正 ' + str(len(labels)) + ' 条、模型候选 ' + str(len(model_labels)) + ' 条；保留项另存', directory=str(export))
        return dict(path=str(export), explicit_corrections=len(labels), model_candidates=len(model_labels), not_labels=len(retained))


class Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = False


class ServiceLock:
    """One service owns a UI data directory; no parallel stale workspace writers."""
    def __init__(self, data):
        data.mkdir(parents=True, exist_ok=True)
        self.stream = (data / '.ui_service.lock').open('a+b')
        if self.stream.seek(0, 2) == 0:
            self.stream.write(b'0'); self.stream.flush()
        self.stream.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            self.stream.close(); self.stream = None
            raise ValueError('此数据目录已有界面服务运行，请打开原窗口；切换Python前先退出原服务。') from exc
        atexit.register(self.close)

    def close(self):
        if self.stream is not None:
            try:
                self.stream.seek(0)
                if os.name == 'nt':
                    import msvcrt
                    msvcrt.locking(self.stream.fileno(), msvcrt.LK_UNLCK, 1)
            finally:
                self.stream.close(); self.stream = None


def serve(core, data, port=0, browser=True):
    guard = ServiceLock(data)
    app = Application(core, data); token = secrets.token_urlsafe(32)
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args): pass

        def send(self, status, content, mime='application/json; charset=utf-8', **headers):
            if not isinstance(content, bytes): content = json.dumps(content, ensure_ascii=False).encode('utf8')
            self.send_response(status); self.send_header('Content-Type', mime)
            self.send_header('Content-Length', str(len(content))); self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff'); self.send_header('Referrer-Policy', 'no-referrer')
            self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'")
            for key, value in headers.items(): self.send_header(key.replace('_', '-'), value)
            self.end_headers(); self.wfile.write(content)

        def authorized(self):
            origin = 'http://127.0.0.1:' + str(self.server.server_port)
            if self.headers.get('Host') != '127.0.0.1:' + str(self.server.server_port): return False
            if self.headers.get('Origin') not in (None, origin): return False
            return secrets.compare_digest(self.headers.get('X-WPS-UI-Token', ''), token)

        def do_GET(self):
            parsed = urlparse(self.path); path = parsed.path; args = parse_qs(parsed.query)
            if path in ('/', '/app.js', '/style.css'):
                file = HERE / 'web' / ('index.html' if path == '/' else path[1:])
                mime = 'text/html; charset=utf-8' if path == '/' else 'text/javascript; charset=utf-8' if path.endswith('.js') else 'text/css; charset=utf-8'
                return self.send(200, file.read_bytes(), mime)
            if not self.authorized(): return self.send(403, dict(message='本机界面授权失效，请从启动器重新打开。'))
            try:
                with app.lock:
                    if path == '/api/state': value = app.summary()
                    elif path == '/api/form': value = app.form_data(args['kind'][0])
                    elif path == '/api/context': value = app.context(int(args['row'][0]))
                    elif path == '/api/logs':
                        start = max(0, int(args.get('after', ['0'])[0])); value = dict(lines=[json.loads(x) for x in app.log_lines[start:]], count=len(app.log_lines))
                    else: return self.send(404, dict(message='未知接口'))
                self.send(200, value)
            except Exception as exc: self.error(exc)

        def error(self, exc):
            data = dict(message=str(exc), exception=type(exc).__name__)
            if isinstance(exc, Conflict): data.update(exc.data)
            app.note('错误', str(exc))
            self.send(409 if isinstance(exc, Conflict) else 400, data)

        def do_POST(self):
            if not self.authorized(): return self.send(403, dict(message='本机界面授权失效'))
            try:
                length = int(self.headers.get('Content-Length', '0'))
                if not 0 <= length <= 4_000_000: raise ValueError('请求过大')
                payload = json.loads(self.rfile.read(length) or b'{}'); path = urlparse(self.path).path
                if path == '/api/attach':
                    with app.lock: value = app.attach(payload['run'])
                elif path == '/api/sync':
                    with app.lock:
                        value = app.sync(payload['kind'], payload['version'], payload.get('changes'), payload.get('resolutions'), payload.get('import_path'), payload.get('origin', 'user'))
                elif path == '/api/submit': value = app.start_job(payload['operation'], payload)
                elif path == '/api/open-word': value = app.open_documents(payload.get('row'), payload.get('side', 'both'), payload.get('plain', False))
                elif path == '/api/open-template-tool': value = app.open_template_tool()
                elif path == '/api/stop':
                    if app.active: app.jobs[app.active]['stop_requested'] = True
                    value = dict(status='STOP_REQUESTED')
                elif path == '/api/choose':
                    if payload.get('kind') not in ('directory', 'docx', 'xlsx'): raise ValueError('选择类型无效')
                    result = subprocess.run([sys.executable, '-B', str(HERE / 'ui_tool.py'), 'choose', '--kind', payload['kind']],
                                            capture_output=True, text=True, encoding='utf8', timeout=180,
                                            creationflags=0x08000000 if os.name == 'nt' else 0)
                    if result.returncode: raise ValueError(result.stderr or '文件选择器未完成')
                    value = json.loads(result.stdout)
                elif path == '/api/open-form':
                    run, work, disk, _ = app.form(payload['kind'])
                    if os.name != 'nt': raise ValueError('系统打开需要Windows')
                    os.startfile(str(run / FORMS[payload['kind']][1])); value = dict(status='OPENED')
                elif path == '/api/export-form':
                    content = app.exported_form(payload['kind'])
                    return self.send(200, content, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', Content_Disposition='attachment; filename="' + payload['kind'] + '_form.xlsx"')
                elif path == '/api/reset-form':
                    with app.lock: value = app.reset_form(payload['kind'], payload['version'])
                elif path == '/api/finalize':
                    with app.lock: value = app.finalize(bool(payload.get('accept_pending')), payload.get('origin', 'user'))
                elif path == '/api/training':
                    with app.lock: value = app.training_export()
                elif path == '/api/shutdown':
                    app.assert_idle(); self.send(200, dict(status='STOPPING'))
                    threading.Thread(target=self.server.shutdown, daemon=True).start(); return
                else: return self.send(404, dict(message='未知接口'))
                self.send(200, value)
            except Exception as exc: self.error(exc)
    server = Server(('127.0.0.1', port), Handler)
    url = 'http://127.0.0.1:' + str(server.server_port) + '/#token=' + token
    write(data / 'server.json', dict(url=url, port=server.server_port, pid=os.getpid(), python=sys.executable, version=VERSION, ui_build=BUILD_ID, core=str(core)))
    print(json.dumps(dict(status='SERVING', url=url, python=sys.executable), ensure_ascii=False), flush=True)
    if browser: webbrowser.open(url)
    try: server.serve_forever(poll_interval=.4)
    except KeyboardInterrupt: pass
    finally:
        server.server_close()
        guard.close()
        if app.native and app.native.poll() is None:
            try: app.native.stdin.close()
            except Exception: pass
    return 0

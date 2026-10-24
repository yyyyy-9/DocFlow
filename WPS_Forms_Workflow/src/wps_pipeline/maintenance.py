"""Read-only product identity and explicit retirement of completed Forms runs.

Uses the standard library only. Cleanup is independent of recognition/runtime
fingerprints so old runs can be retired without pretending to migrate them.
"""
from contextlib import contextmanager
from pathlib import Path
from uuid import uuid4
import ast
import hashlib
import json
import os
import re
import shutil
import stat
from .job_support import ROOT, VERSION, read, write, sha, stamp, contained

GENERATED_DIRS = {'source_scan', 'preprocessed', 'prepared', 'wps_pages'}
GENERATED_FILES = {
    '01_输入.xlsx', '02_扫描核查.xlsx', '03_反馈提交.xlsx', '反馈处理回执.xlsx',
    'submitted_classification.xlsx', 'submitted_feedback.xlsx', 'approved_plan.json',
    'feedback_export.json', 'feedback_ledger.json', 'forms_state.json', 'inventory.json',
    'page_facts.json', 'preprocessing.json', 'DELIVERY.md',
}
PENDING = 'CLEANUP_PENDING.json'
CLOSED = 'CLEANED_TASK.json'
ARCHIVE = '阶段总稿'


def identity():
    """Name the wrapper and protocol separately, and expose inconsistent copies."""
    build = read(ROOT / 'BUILD_INFO.json')
    config = read(ROOT / 'config.json')
    tree = ast.parse((ROOT / 'src/wps_pipeline/forms_workflow.py').read_text('utf-8-sig'))
    forms = next(ast.literal_eval(n.value) for n in tree.body
                 if isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'FORMS_VERSION' for t in n.targets))
    errors = []
    if build.get('forms_version') != forms: errors.append('BUILD_INFO.forms_version 与实际三表代码不同')
    if build.get('protocol') != VERSION or config.get('version') != VERSION:
        errors.append('BUILD_INFO/config 与实际内核协议不同')
    if not build.get('build_id'): errors.append('缺少构建号')
    if config.get('release_status') != 'FORMS_WORKFLOW_RELEASE': errors.append('非当前三表发行展示标签')
    return dict(product='WPS_Forms_Workflow', product_name='WPS 三表文档整理',
                forms_version=forms, kernel_protocol=VERSION, build_id=build.get('build_id'),
                release_status=config.get('release_status'),
                identity_status='CONSISTENT' if not errors else 'IDENTITY_MISMATCH', identity_errors=errors)


def is_link(path):
    info = path.lstat()
    return stat.S_ISLNK(info.st_mode) or bool(getattr(info, 'st_file_attributes', 0) & 0x400)


def real_directory(path):
    value = Path(path).expanduser().absolute()
    for p in (value, *value.parents):
        if p.exists() and is_link(p): raise ValueError('目录含链接/重解析点，请使用实际目录：' + str(p))
    value = value.resolve()
    if not value.is_dir(): raise ValueError('任务目录不存在：' + str(value))
    return value


def local_path(run, relative):
    """Validate the lexical path before resolving away any in-run junction."""
    resolved = contained(run, relative)
    current = Path(run) / str(relative).replace('\\', '/')
    while current != run:
        try:
            if is_link(current): raise ValueError('任务路径含链接/重解析点，保留现场：' + str(current))
        except FileNotFoundError:
            pass
        current = current.parent
    return resolved


@contextmanager
def task_lease(run):
    """One non-blocking OS lease shared by CLI revisions and cleanup."""
    run = real_directory(run)
    path = local_path(run, '.maintenance.lock')
    stream = path.open('a+b')
    locked = False
    try:
        stream.seek(0, 2)
        if stream.tell() == 0: stream.write(b'0'); stream.flush()
        stream.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            locked = True
        except OSError as exc:
            raise ValueError('该任务正在提交或清理，请等待当前操作结束：' + str(run)) from exc
        yield
    finally:
        if locked:
            stream.seek(0)
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)
        stream.close()
        # Keep the tiny lock file: unlinking while another opener owns the old
        # inode would let a third process acquire a different lock file.


def assert_open(run):
    for name in (CLOSED, PENDING):
        if (Path(run) / name).exists():
            raise ValueError('此任务已清理或正在清理；请查看“阶段总稿”和清理回执。旧02/03不能继续提交；需要整理时从原始来源新建任务。')


def generated_name(name):
    return (name in GENERATED_DIRS or name in GENERATED_FILES or
            bool(re.fullmatch(r'draft_\d{3,}(?:\.status\.json)?', name)))


def snapshot(run, names):
    result = {}
    for name in names:
        if not generated_name(name): raise ValueError('非程序产物不能删除：' + name)
        target = local_path(run, name)
        if not target.exists(): continue
        if is_link(target): raise ValueError('程序产物中含链接：' + str(target))
        directory = name in GENERATED_DIRS or bool(re.fullmatch(r'draft_\d{3,}',name))
        if directory != target.is_dir(): raise ValueError('程序路径类型改变，保留现场：' + str(target))
        if target.is_file(): paths = [target]
        else:
            paths = []
            for base, dirs, names_here in os.walk(target, followlinks=False):
                for child in [*dirs, *names_here]:
                    p = Path(base, child)
                    if is_link(p): raise ValueError('产物内部含链接，保留该任务：' + str(p))
                if 'forms_state.json' in names_here: raise ValueError('产物内还有独立任务，保留该任务：' + str(base))
                paths.extend(Path(base, n) for n in names_here)
        for p in paths:
            info = p.stat()
            result[p.relative_to(run).as_posix()] = [info.st_size, info.st_mtime_ns]
    return result


def discover(root):
    runs = []; skipped = []
    for base, dirs, names in os.walk(root, followlinks=False):
        p = Path(base)
        if 'pyvenv.cfg' in names: dirs[:] = []; continue
        dirs[:] = [d for d in dirs if d not in {'.git', '.codex', 'node_modules', '__pycache__', ARCHIVE}
                   and not is_link(p / d)]
        if CLOSED in names: dirs[:] = []; continue
        if PENDING in names:
            try:
                plan = read(p / PENDING)
                if plan.get('schema') != 'wps-cleanup-pending-1' or Path(plan['run']).resolve() != p:
                    raise ValueError('清理暂存记录不匹配')
                runs.append(dict(path=p, state=plan['task'], pending=plan))
            except Exception as exc: skipped.append(dict(run=str(p), reason=str(exc)))
            dirs[:] = []; continue
        if 'forms_state.json' not in names: continue
        try:
            state = read(p / 'forms_state.json')
            if state.get('schema') != 'wps-three-forms-1': raise ValueError('不是支持的三表任务')
            runs.append(dict(path=p, state=state))
        except Exception as exc: skipped.append(dict(run=str(p), reason=str(exc)))
        dirs[:] = []
    return runs, skipped


def terminal(state):
    return str(state.get('status', '')).startswith('COMPLETED') or state.get('status') == 'NEEDS_REPAIR'


def source_proofs(run, state, inventory):
    proofs = {}
    data_root = Path(inventory['root']).resolve()
    for source in inventory['sources']:
        original = source.get('preprocessing_original') or {}
        relative = original.get('relative', source.get('relative'))
        expected = original.get('sha256', source['sha256'])
        path = contained(data_root, relative)
        if not path.is_file() or sha(path) != expected:
            raise ValueError('输入原件缺失或已改变，保留冻结副本所在任务：' + str(path))
        proofs[str(path)] = expected
    frozen = run / 'source_scan/inputs/template.docx'
    if not frozen.is_file() or sha(frozen) != state['frozen_template_sha256']:
        raise ValueError('冻结模板缺失或改变，保留现场')
    source_state = read(run / 'source_scan/state.json')
    raw = source_state.get('job', {}).get('template', '')
    candidates = [Path(raw)] if raw and Path(raw).is_absolute() else []
    candidates.append(ROOT / 'templates/target_template.docx')
    match = next((p.resolve() for p in candidates if p.is_file() and sha(p) == state['frozen_template_sha256']
                  and not p.resolve().is_relative_to(run)), None)
    if match is None: raise ValueError('未找到任务外相同的用户模板，保留冻结模板所在任务')
    proofs[str(match)] = state['frozen_template_sha256']
    return proofs


def plan_task(item):
    run, state = item['path'], item['state']
    if item.get('pending'): return item['pending']
    if not terminal(state) or not state.get('drafts'): raise ValueError('尚未交稿或仍在执行，保留任务')
    inventory_path = run / 'inventory.json'
    if sha(inventory_path) != state['inventory_sha256']: raise ValueError('来源清单绑定改变，保留现场')
    inventory = read(inventory_path)
    proofs = source_proofs(run, state, inventory)
    masters = []; manual_outputs = []
    known_ids = {s['id'] for s in inventory['sources']}
    approved = read(run / 'approved_plan.json') if (run / 'approved_plan.json').is_file() else {}
    known_ids.update(e['id'] for e in approved.get('entries', []))
    for draft in state['drafts']:
        if not re.fullmatch(r'draft_\d{3,}', draft['run']): raise ValueError('异常稿件目录')
        core = contained(run, draft['run']); core_state = read(core / 'state.json')
        known_ids.update(s['id'] for s in core_state.get('sources', []))
        if not terminal(core_state): raise ValueError('内核仍在执行，保留任务')
        declared = {contained(core, d['file']): d for d in core_state.get('deliverables', [])}
        for p, d in declared.items():
            if d.get('id') != 'master' and p.is_file() and sha(p) != d.get('sha256'):
                manual_outputs.append(str(p))
        if manual_outputs: raise ValueError('单篇稿已被手工改动，保留所在任务：' + manual_outputs[0])
        for build in sorted(core.iterdir()):
            if not build.is_dir() or not re.fullmatch(r'build_\d{3,}', build.name): continue
            paths = [p for p in (build / 'documents').glob('master*.docx')
                     if p.name in ('master.docx', 'master.assembled-draft.docx') and p.is_file()]
            if not paths and (build / 'assembled/master.docx').is_file(): paths = [build / 'assembled/master.docx']
            for p in paths:
                if is_link(p): raise ValueError('总稿不能是链接')
                value = sha(p); declared_hash = declared.get(p, {}).get('sha256')
                dest = f'{ARCHIVE}/{core.name}_{build.name}_{p.stem}_{value[:12]}.docx'
                masters.append(dict(source=p.relative_to(run).as_posix(), file=dest, sha256=value,
                                    bytes=p.stat().st_size, manually_changed=bool(declared_hash and declared_hash != value),
                                    assembled_only=p.parent.name == 'assembled'))
    if not masters: raise ValueError('没有实际总稿，保留任务及失败记录')
    names = sorted(p.name for p in run.iterdir() if generated_name(p.name))
    signature = snapshot(run, names)
    # Unknown Office files may be a user's own draft placed in a generated
    # folder. Do not treat the entire directory as proof of file ownership.
    for relative in signature:
        path = Path(relative)
        if path.suffix.lower() not in ('.doc','.docx','.xlsx'): continue
        if len(path.parts)==1 and path.name in GENERATED_FILES: continue
        stem = path.stem.removesuffix('.assembled-draft')
        registered = stem in known_ids | {'template'}
        if len(path.parts)>=4 and re.fullmatch(r'draft_\d{3,}',path.parts[0]) and path.parent.name in ('documents','assembled'):
            registered = stem in known_ids | {'master'}
        if path.parts[0]=='preprocessed' and 'wps' in path.parts:
            registered = path.stem in {'input','source','read_only_source','readonly_source','converted','converted_source'}
        if not registered: raise ValueError('产物目录有未登记的Office文件，保留该任务：' + relative)
    return dict(schema='wps-cleanup-pending-1', run=str(run), task={k:state.get(k) for k in ('run_id','created','status','parent','title','version')},
                roots=names, files=signature, source_proofs=proofs, masters=masters,
                state_sha256=sha(run / 'forms_state.json'), inventory_sha256=sha(inventory_path))


def validate_plan(run, plan, resume=False):
    if Path(plan['run']).resolve() != run: raise ValueError('清理路径与任务不符')
    for name in plan['roots']:
        if not generated_name(name) or '/' in name or '\\' in name: raise ValueError('清理记录含非程序路径')
    for path, value in plan['source_proofs'].items():
        p = Path(path)
        if not p.is_absolute() or not p.is_file() or sha(p) != value: raise ValueError('原件/模板已改变，暂不删除：' + path)
    current = snapshot(run, plan['roots'])
    if (current != plan['files'] if not resume else any(plan['files'].get(k) != v for k,v in current.items())):
        raise ValueError('任务文件已变化或有新增文件，暂不删除')
    if not resume:
        if sha(run / 'forms_state.json') != plan['state_sha256'] or sha(run / 'inventory.json') != plan['inventory_sha256']:
            raise ValueError('任务状态已改变，重新预览后清理')


def execute_plan(run, plan, resume=False):
    with task_lease(run):
        if (run / CLOSED).exists(): return dict(run=str(run), status='ALREADY_CLEANED')
        validate_plan(run, plan, resume)
        for master in plan['masters']:
            source = local_path(run, master['source']); dest = local_path(run, master['file'])
            if not master['file'].startswith(ARCHIVE + '/'): raise ValueError('总稿保存路径非法')
            if dest.exists():
                if is_link(dest) or sha(dest) != master['sha256']: raise ValueError('已有保留总稿不同，禁止覆盖')
            else:
                if not source.is_file() or sha(source) != master['sha256']: raise ValueError('源总稿已改变')
                dest.parent.mkdir(parents=True, exist_ok=True)
                if is_link(dest.parent): raise ValueError('总稿目录不能是链接')
                temp = dest.with_name('.' + uuid4().hex + '.tmp')
                try:
                    shutil.copyfile(source, temp)
                    if sha(temp) != master['sha256']: raise ValueError('总稿复制校验失败')
                    with temp.open('r+b') as stream: os.fsync(stream.fileno())
                    # Windows rename fails if the destination already exists;
                    # POSIX hard-link creation supplies the same no-overwrite
                    # atomic publication of a complete sibling file.
                    if os.name == 'nt': os.rename(temp, dest)
                    else: os.link(temp, dest)
                finally:
                    if temp.exists(): temp.unlink()
            if sha(dest) != master['sha256']: raise ValueError('保留总稿哈希不一致')
        # Every master is durable and verified before the first deletion.
        write(run / PENDING, plan)
        for name in plan['roots']:
            path = local_path(run, name)
            if path.exists():
                if is_link(path): raise ValueError('删除前发现链接，停止本任务')
                if path.is_dir(): shutil.rmtree(path)
                else: path.unlink()
        result = dict(schema='wps-cleaned-task-1', status='CLEANED', time=stamp(), run=str(run), task=plan['task'],
                      masters=plan['masters'], deleted_files=len(plan['files']),
                      removed_bytes=sum(x[0] for x in plan['files'].values()),
                      retained_master_bytes=sum(x['bytes'] for x in plan['masters']),
                      estimated_reclaimed_bytes=max(0,sum(x[0] for x in plan['files'].values())-sum(x['bytes'] for x in plan['masters'])),
                      can_submit_old_forms=False, note='只保留各阶段总稿及最小清理记录；原件和任务外模板未删除。重新整理请从来源新建任务。')
        write(run / CLOSED, result);(run / PENDING).unlink()
        return result


def cleanup_tasks(root, apply=False, keep_latest=False):
    root = real_directory(root); runs, skipped = discover(root)
    bypath = {x['path']:x for x in runs}; protected = set()
    def protect_chain(item):
        seen = set()
        while item and item['path'] not in seen:
            seen.add(item['path']); protected.add(item['path'])
            parent = item['state'].get('parent')
            if not parent: break
            parent_path = Path(parent).resolve()
            item = bypath.get(parent_path)
            if item is None:
                # Revisions can be stored outside the selected parent folder
                # and then point back into it. Read that bridge, never delete
                # outside the explicitly selected root.
                try:
                    parent_path = real_directory(parent)
                    parent_state = read(local_path(parent_path, 'forms_state.json'))
                    if parent_state.get('schema') != 'wps-three-forms-1': raise ValueError('未知父任务')
                    item = dict(path=parent_path, state=parent_state)
                except Exception:
                    protected.update(bypath)
                    break
    for item in runs:
        if not terminal(item['state']): protect_chain(item)
    if keep_latest and runs:
        protect_chain(max(runs, key=lambda x:(str(x['state'].get('created','')),str(x['path']))))
    plans = []
    for item in runs:
        if item['path'] in protected:
            skipped.append(dict(run=str(item['path']), reason='保留最新/未完成任务及其修订依赖链'));continue
        try: plans.append((item['path'], plan_task(item), bool(item.get('pending'))))
        except Exception as exc: skipped.append(dict(run=str(item['path']), reason=str(exc)))
    # A source/template used by any discovered task is never removed as a
    # generated folder belonging to another task selected in this same batch.
    external = []
    for _,plan,_ in plans: external.extend(Path(p).resolve() for p in plan['source_proofs'])
    for item in runs:
        try:
            if item.get('pending'): external.extend(Path(p).resolve() for p in item['pending']['source_proofs'])
            else:
                inv = read(item['path'] / 'inventory.json')
                external.extend(contained(Path(inv['root']).resolve(), (s.get('preprocessing_original') or {}).get('relative',s.get('relative'))) for s in inv['sources'])
                template=read(item['path']/'source_scan/state.json').get('job',{}).get('template')
                if template and Path(template).is_absolute(): external.append(Path(template).resolve())
        except Exception:
            # Do not prune ancestors of a task whose provenance cannot be read.
            protect_chain(item)
    accepted = []
    for run, plan, resume in plans:
        targets = [contained(run,n) for n in plan['roots']]
        if run in protected or any(p==t or p.is_relative_to(t) for p in external for t in targets):
            skipped.append(dict(run=str(run), reason='仍被其他任务当作原件/依赖，保留任务'));continue
        accepted.append((run,plan,resume))
    results = []
    for run, plan, resume in accepted:
        if not apply:
            results.append(dict(run=str(run), status='PLANNED', masters=len(plan['masters']),
                                generated_files=len(plan['files']), generated_bytes=sum(v[0] for v in plan['files'].values()),
                                master_files=[m['file'] for m in plan['masters']]));continue
        try: results.append(execute_plan(run,plan,resume))
        except Exception as exc:
            results.append(dict(run=str(run), status='CLEANUP_INCOMPLETE', error=str(exc),
                                note='总稿先保留；如已有CLEANUP_PENDING.json，可再次执行同一清理命令接续。'))
    failed = any(r['status']=='CLEANUP_INCOMPLETE' for r in results)
    status = ('PREVIEW' if not apply else 'PARTIAL' if failed else
              'NOTHING_TO_CLEAN' if not runs else 'NO_ELIGIBLE_TASKS' if not results else
              'COMPLETED_WITH_SKIPPED_TASKS' if skipped else 'COMPLETED')
    result = dict(status=status, root=str(root),
                  mode='KEEP_LATEST_CHAIN' if keep_latest else 'STAGE_MASTERS_ONLY', apply=apply,
                  cleaned_tasks=sum(r['status']=='CLEANED' for r in results), planned_tasks=len(accepted),
                  skipped_tasks=skipped, tasks=results, program_model_api_calls=0,
                  note='仅留总稿模式会结束旧任务的02/03接续能力；原输入/用户模板和未知根文件保留。--keep-latest保留最新任务及其依赖链。')
    if apply and results:
        try:
            with local_path(root, 'cleanup_history.jsonl').open('a',encoding='utf8') as stream:
                stream.write(json.dumps(dict(time=stamp(), mode=result['mode'],cleaned_tasks=result['cleaned_tasks'],
                                             runs=[dict(run=r['run'],status=r['status']) for r in results]),ensure_ascii=False)+'\n')
        except Exception as exc:
            result['history_log_error'] = str(exc)
            result['note'] += ' 父目录累计日志未写入；各任务CLEANED_TASK及本次返回结果仍有效。'
    return result

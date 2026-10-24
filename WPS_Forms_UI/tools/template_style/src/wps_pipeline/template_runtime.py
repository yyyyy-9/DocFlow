"""Template-tool process adapter. Document rules stay in template_prepare."""
from pathlib import Path
from zipfile import ZipFile
import hashlib
import importlib.metadata
import json
import os
import platform
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from contextlib import contextmanager

from . import resource_root

TOOL_VERSION = '1.0.3'
REQUEST_SCHEMA = 'wps-template-runtime-v1'
ROLE_NAMES = {'document': '输入文档', 'source': '带封面的模板', 'style_source': '样式来源'}


def runtime_facts():
    try:
        lxml_version = importlib.metadata.version('lxml')
    except importlib.metadata.PackageNotFoundError:
        try:
            from lxml import etree
            lxml_version = '.'.join(map(str, etree.LXML_VERSION))
        except (ImportError, OSError):
            lxml_version = None
    return {'tool_version': TOOL_VERSION, 'release_channel': 'stable',
            'executable': sys.executable, 'python': platform.python_version(),
            'frozen': bool(getattr(sys, 'frozen', False)), 'pid': os.getpid(),
            'lxml': lxml_version}


def probe_file(path, role):
    """Hash and inspect one read handle in this process; never dump document text."""
    result = {'role': role, 'label': ROLE_NAMES.get(role, role),
              'path': str(Path(path).resolve()), 'observed_at': datetime.now(timezone.utc).isoformat()}
    try:
        digest = hashlib.sha256()
        count = 0
        with Path(path).open('rb') as stream:
            result['stat_bytes'] = os.fstat(stream.fileno()).st_size
            header = stream.read(8)
            stream.seek(0)
            while chunk := stream.read(1024 * 1024):
                count += len(chunk)
                digest.update(chunk)
            result.update(bytes_read=count, sha256=digest.hexdigest(), header_hex=header.hex())
            stream.seek(0)
            try:
                with ZipFile(stream) as archive:
                    names = set(archive.namelist())
                    result.update(zip_valid=True, has_document='word/document.xml' in names,
                                  has_styles='word/styles.xml' in names)
            except Exception as exc:
                result.update(zip_valid=False, zip_error_type=type(exc).__name__, zip_error=str(exc))
    except Exception as exc:
        result.update(read_error_type=type(exc).__name__, read_error=str(exc))
    return result


def builtin_style_source():
    root = resource_root()
    config = json.loads((root / 'config.json').read_text('utf-8-sig'))
    return str((root / config['template']).resolve())


def _path(value, role):
    if not value or not str(value).strip():
        raise ValueError('请选择' + ROLE_NAMES.get(role, role))
    return str(Path(value).resolve())


def execute_request(request):
    """Execute the same backend in source, embedded Python, or selected Python."""
    result = {'runtime': runtime_facts()}
    inputs = {}
    phase = 'request'
    failed_role = None
    try:
        if request.get('schema') != REQUEST_SCHEMA:
            raise ValueError('Unsupported template request schema')
        action = request['action']
        params = request.get('parameters', {})
        if action not in ('inspect', 'inspect-pair', 'prepare', 'probe'):
            raise ValueError('Unsupported template action')
        if action == 'inspect':
            inputs['document'] = _path(params.get('document'), 'document')
        else:
            if params.get('source') or action != 'probe':
                inputs['source'] = _path(params.get('source'), 'source')
            inputs['style_source'] = _path(params.get('style_source') or builtin_style_source(), 'style_source')
        if action == 'probe':
            result.update(status='PROBED', inputs=[probe_file(path, role) for role, path in inputs.items()])
            try:
                from lxml import etree
                result['backend_import'] = {'status': 'READY', 'lxml': etree.LXML_VERSION}
            except Exception as exc:
                result['backend_import'] = {'status': 'FAIL', 'error_type': type(exc).__name__, 'error': str(exc)}
            return result
        phase = 'load_backend'
        from .template_prepare import inspect_document, inspect_styles, prepare_template
        if action == 'inspect':
            failed_role = 'document'
            phase = 'read_document'
            document = inspect_document(inputs['document'])
            phase = 'read_styles'
            styles = inspect_styles(inputs['document'])
            result.update(status='INSPECTED', document=document, styles=styles)
        elif action == 'inspect-pair':
            failed_role, phase = 'style_source', 'read_style_source'
            styles = inspect_styles(inputs['style_source'])
            failed_role, phase = 'source', 'read_cover_template'
            document = inspect_document(inputs['source'])
            result.update(status='INSPECTED', document=document, styles=styles)
        else:
            phase = 'prepare_template'
            prepared = prepare_template(inputs['source'], params['output'], params['body_start'],
                                        params['expected_sha256'], style_source=inputs['style_source'],
                                        expected_style_sha256=params.get('expected_style_sha256'))
            result.update(prepared)
        return result
    except Exception as exc:
        details = {'phase': phase, 'error_type': type(exc).__name__, 'failed_role': failed_role,
                   'failed_path': inputs.get(failed_role),
                   'inputs': [probe_file(path, role) for role, path in inputs.items()],
                   'probe_timing': 'Read again after the failed operation; not an atomic snapshot of it.'}
        hint = None
        if phase == 'load_backend':
            missing_lxml = isinstance(exc, ModuleNotFoundError) and (exc.name or '').startswith('lxml')
            hint = ('所选 Python 无法加载 lxml。' if missing_lxml else '无法加载模板处理模块。')
            hint += '请选择已经能运行 template_tool.py 的 Python 环境。'
        location = '\n文件：' + inputs[failed_role] if failed_role in inputs else ''
        label = ROLE_NAMES.get(failed_role, '模板处理')
        message = f'{label}失败（{phase}）{location}\n{type(exc).__name__}: {exc}'
        if hint:
            message += '\n' + hint
        result.update(status='FAIL', error=message, error_details=details)
        return result


class TemplateBackend:
    def __init__(self, python=None):
        self.python = str(python).strip() if python else None

    def run(self, action, **parameters):
        # The child uses a temporary working directory. Resolve user paths here
        # so relative CLI/entry-field paths retain their original meaning.
        for key in ('document', 'source', 'style_source', 'output'):
            if parameters.get(key) is not None and str(parameters[key]).strip():
                parameters[key] = str(Path(parameters[key]).resolve())
        request = {'schema': REQUEST_SCHEMA, 'action': action, 'parameters': parameters}
        if not self.python:
            return execute_request(request)
        runtime = {'tool_version': TOOL_VERSION, 'requested_executable': self.python, 'mode': 'selected-python'}
        try:
            python = Path(self.python).resolve()
            if not python.is_file():
                raise ValueError('请选择实际的 python.exe 路径')
            if getattr(sys, 'frozen', False) and python == Path(sys.executable).resolve():
                raise ValueError('请选择 Python 解释器，不能选择工具 EXE 自身')
            worker = resource_root() / 'template_tool.py'
            if not worker.is_file():
                raise FileNotFoundError('Matching template worker source is missing: ' + str(worker))
            with tempfile.TemporaryDirectory(prefix='wps-template-runtime-') as directory:
                directory = Path(directory)
                request_path, result_path = directory / 'request.json', directory / 'result.json'
                request_path.write_text(json.dumps(request, ensure_ascii=False), encoding='utf-8')
                command = [str(python), '-X', 'utf8', '-B', str(worker),
                           '--worker-request', str(request_path), '--result-json', str(result_path)]
                creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                with external_environment() as environment:
                    process = subprocess.run(command, cwd=directory, capture_output=True, timeout=120,
                                             creationflags=creationflags, env=environment)
                if not result_path.is_file():
                    detail = process.stderr.decode('utf-8', 'replace')[-2500:]
                    raise RuntimeError('Python 未返回处理结果（exit ' + str(process.returncode) + '）。\n' + detail)
                result = json.loads(result_path.read_text('utf-8-sig'))
                if result.get('runtime', {}).get('tool_version') != TOOL_VERSION:
                    raise ValueError('Python worker version does not match this tool')
                if process.returncode != (1 if result.get('status') == 'FAIL' else 0):
                    raise RuntimeError('Python exit code and result disagree')
                result['runtime']['mode'] = 'selected-python'
                return result
        except Exception as exc:
            message = str(exc)
            if isinstance(exc, subprocess.TimeoutExpired):
                message = 'Python 处理超时，输出状态未知。请检查目标文件；工具不会自动重试。'
            return {'status': 'FAIL', 'error': message, 'runtime': runtime,
                    'error_details': {'phase': 'launch_python', 'error_type': type(exc).__name__}}

    def diagnose(self, **parameters):
        embedded = execute_request({'schema': REQUEST_SCHEMA, 'action': 'probe', 'parameters': parameters})
        selected = self.run('probe', **parameters) if self.python else None
        comparison = []
        if selected and selected.get('status') == 'PROBED':
            other = {item['role']: item for item in selected['inputs']}
            for item in embedded.get('inputs', []):
                match = other.get(item['role'], {})
                comparison.append({'role': item['role'], 'path': item['path'],
                                   'same_read_bytes': item.get('bytes_read') == match.get('bytes_read')
                                   and item.get('bytes_read') is not None,
                                   'same_sha256': item.get('sha256') == match.get('sha256')
                                   and item.get('sha256') is not None})
        return {'schema': 'wps-template-diagnostics-v1', 'status': 'DIAGNOSTICS',
                'tool_version': TOOL_VERSION, 'embedded': embedded, 'selected_python': selected,
                'comparison': comparison,
                'note': 'Hashes, lengths, first 8 bytes and process facts only; no document body. Inputs are read independently.'}


@contextmanager
def external_environment():
    """Undo the frozen loader's library search changes for the selected Python.

    https://pyinstaller.org/en/stable/common-issues-and-pitfalls.html#launching-external-programs-from-the-frozen-application
    Restore this process after the child exits; preserve the user's environment
    outside the PyInstaller extraction directory (including conda DLL paths).
    """
    environment = os.environ.copy()
    if not getattr(sys, 'frozen', False):
        yield environment
        return
    bundle = Path(sys._MEIPASS).resolve()
    def bundled_path(value):
        try:
            return bool(value) and Path(value).resolve().is_relative_to(bundle)
        except (OSError, ValueError):
            return False
    for key in list(environment):
        if key.startswith('_PYI_') or key == '_MEIPASS2':
            del environment[key]
    for key in ('PATH', 'PYTHONPATH'):
        if key in environment:
            environment[key] = os.pathsep.join(part for part in environment[key].split(os.pathsep) if not bundled_path(part))
    if bundled_path(environment.get('PYTHONHOME')):
        environment.pop('PYTHONHOME', None)
    kernel, previous = None, None
    if sys.platform == 'win32':
        import ctypes
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.GetDllDirectoryW.argtypes = [ctypes.c_uint32, ctypes.c_wchar_p]
        kernel.GetDllDirectoryW.restype = ctypes.c_uint32
        kernel.SetDllDirectoryW.argtypes = [ctypes.c_wchar_p]
        kernel.SetDllDirectoryW.restype = ctypes.c_int
        buffer = ctypes.create_unicode_buffer(32768)
        if kernel.GetDllDirectoryW(len(buffer), buffer):
            previous = buffer.value
        if not kernel.SetDllDirectoryW(None):
            raise ctypes.WinError(ctypes.get_last_error())
    try:
        yield environment
    finally:
        if kernel is not None:
            kernel.SetDllDirectoryW(previous)


def settings_path():
    base = Path(os.environ.get('LOCALAPPDATA') or Path.home() / '.config')
    return base / 'WPS_Template_Style_Tool' / 'settings.json'


def load_python_preference():
    explicit = os.environ.get('WPS_TEMPLATE_PYTHON')
    if explicit:
        return explicit
    try:
        settings = json.loads(settings_path().read_text('utf-8'))
        return (settings.get('python') or '') if isinstance(settings, dict) else ''
    except (OSError, ValueError):
        return ''


def save_python_preference(python):
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps({'python': str(python) if python else ''}, ensure_ascii=False), 'utf-8')
    temporary.replace(path)

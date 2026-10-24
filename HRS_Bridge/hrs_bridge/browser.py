"""Real local HRS server and original browser engine; never clicks UI controls."""
from contextlib import contextmanager
import importlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
import threading

from .runtime import BridgeError, new_path


@contextmanager
def session(runtime, project=None):
    studio = importlib.import_module('studio')
    portable = importlib.import_module('portable_host')
    with tempfile.TemporaryDirectory(prefix='hrs-bridge-') as directory:
        folder = Path(directory)
        target = folder / 'project'
        # The real engine sees a private snapshot. Original projects never receive
        # drafts, registry, output files or service lifecycle writes.
        if project:
            target.mkdir()
            shutil.copyfile(Path(project) / 'report.md', target / 'report.md')
            shutil.copytree(Path(project) / 'assets', target / 'assets')
        server = studio.StudioServer(('127.0.0.1', 0), target, folder / 'workspace')
        class QuietHandler(portable.PortableHandler):
            def log_message(self, *args):
                pass
        server.RequestHandlerClass = QuietHandler
        server.portable_lifecycle = portable.WindowSessions('bridge', enabled=False)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            yield 'http://127.0.0.1:' + str(server.server_port)
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)


def call(runtime, operation, *, project=None, **arguments):
    if not runtime.node:
        raise BridgeError('Node.js and playwright-core are required for HRS browser operations')
    studio = importlib.import_module('studio')
    browser = runtime.browser or studio.pick_browser()
    if not browser:
        raise BridgeError('Choose an installed Chrome/Chromium executable')
    with session(runtime, project) as base:
        request = dict(operation=operation, base=base, browser=browser, **arguments)
        result = subprocess.run([runtime.node, str(Path(__file__).with_name('browser_worker.cjs'))],
                                input=json.dumps(request, ensure_ascii=False), capture_output=True,
                                text=True, encoding='utf8', timeout=180)
    if result.returncode:
        raise BridgeError(result.stderr.strip()[-3000:] or 'Browser operation failed')
    return json.loads(result.stdout)


def _html_api(runtime):
    try:
        api = importlib.import_module('html_cli')
    except ImportError as error:
        raise BridgeError('Use HRS PF13.8 or newer for the HRS-owned HTML command interface') from error
    if Path(api.__file__).resolve().parent != runtime.hrs.resolve():
        raise BridgeError('Do not mix two HRS HTML command versions in the same process')
    return api


def recover(runtime, source, output):
    try:
        return _html_api(runtime).recover_html(source, new_path(output), browser=runtime.browser)
    except ValueError as error:
        raise BridgeError(str(error)) from error


def export(runtime, project, output, quality='original', mode='recompilable', *, accept_numbered_headings=False):
    try:
        return _html_api(runtime).export_html(project, new_path(output), browser=runtime.browser,
                                              quality=quality, source_mode=mode,accept_numbered_headings=accept_numbered_headings)
    except ValueError as error:
        raise BridgeError(str(error)) from error

"""Delegate reads and explicit plan/apply to the HRS-owned interfaces."""
import importlib
from pathlib import Path
from .runtime import BridgeError


def read(runtime, source, action='objects', object_id=None, index=None, expected=None):
    agent=importlib.import_module('agent_readonly')
    ir=agent.analyze_file(Path(source))
    return agent.project(ir,action,object_id=object_id,block_index=index,expected_sha256=expected)


def _commands(runtime):
    api=importlib.import_module('block_cli')
    if Path(api.__file__).resolve().parent!=runtime.hrs:raise BridgeError('Do not mix two HRS block command versions')
    return api


def plan(runtime, source, expected, replacement, *, object_id=None, index=None, operation='replace'):
    try:return _commands(runtime).plan(source,expected,replacement,object_id=object_id,index=index,operation=operation)
    except (ValueError,OSError) as error:raise BridgeError(str(error)) from error


def apply(runtime, plan_path, expected_plan_sha256):
    try:return _commands(runtime).apply_plan(plan_path,expected_plan_sha256)
    except (ValueError,OSError) as error:raise BridgeError(str(error)) from error

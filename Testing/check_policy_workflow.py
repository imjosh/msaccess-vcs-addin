"""Qualify request-owned policies against actual disposable contents and history.

Exercises external API, APIAsync, and public MCP single-object tools under every
policy, core and separate-component paths, ordinary cycles, history migration,
deletion limits, and mixed conflict lists. Keep evidence outside the checkout.
Required calls and setup are never hidden in assert expressions.
"""
import argparse
import gc
import hashlib
import json
import os
from pathlib import Path
import sys

import pythoncom
import win32api
import win32com.client
import win32event

from check_noninteractive_policies import Endpoint
from request_policy_cases import Cases, require
from policy_artifact_cases import run_artifacts


def run(args):
    out = args.evidence.resolve()
    out.mkdir(parents=True, exist_ok=True)
    folder = out / (args.label + '-user-host')
    folder.mkdir(exist_ok=False)
    environment = {'ACCESS_VCS_PROJECT_DIR': str(folder),
                   'ACCESS_VCS_DIAGNOSTIC_LOG_DIR': str(folder / 'mcp-diagnostics'),
                   'ACCESS_VCS_LOG_DIR': str(folder / 'mcp-logs')}
    for key, value in environment.items():
        Path(value).mkdir(parents=True, exist_ok=True)
        os.environ[key] = value
    # Configure fresh project/log roots before any MCP import resolves user config.
    from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none, pid_from_access_app
    require(list_access_pids_or_none() == set(), 'Serialize native work')
    driver = str(Path(os.environ['APPDATA']) / 'MSAccessVCS/Version Control')
    record = {'cases': {}, 'legacy_control': args.legacy_control, 'environment': environment,
              'client_source': str(args.client_source.resolve()),
              'installed_sha256': hashlib.sha256(Path(driver + '.accda').read_bytes()).hexdigest(),
              'coverage': {'case_set': args.case_set, 'transports': args.transports,
                           'qualified_types_if_entire_run_passes':
                               (['query', 'db_property', 'table_data', 'table'] if args.case_set in ['all', 'core'] else []) +
                               (['command_bar', 'theme', 'table_data_macro', 'table'] if args.case_set in ['all', 'artifacts'] and not args.legacy_control else []) +
                               (['query', 'table', 'table_data'] if args.case_set == 'deletions' else []),
                           'grouped_table_macro_transports': ['external'],
                           'other_types': 'Unqualified by this driver; report separately.'}}
    def save():
        (out / (args.label + '.json')).write_text(json.dumps(record, indent=2, default=str), encoding='utf-8')
    app = handle = endpoint = cases = None
    pid = stamp = None
    pythoncom.CoInitialize()
    try:
        endpoint = Endpoint()
        app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Access.Application'))
        pid = pid_from_access_app(app)
        stamp = process_create_time(pid)
        require(bool(pid and stamp), 'Owned Access process identity unavailable')
        handle = win32api.OpenProcess(0x1000 | 0x100000, False, pid)
        record['process'] = {'pid': pid, 'create_time': stamp}
        cases = Cases(app, driver, folder, endpoint, record, save, args.legacy_control)
        if args.case_set in ['all', 'core']:
            cases.matrix(args.transports)
            if not args.legacy_control:
                cases.cycles_and_migration()
        if not args.legacy_control and args.case_set in ['all', 'core', 'deletions']:
            cases.deletion_matrix()
        if not args.legacy_control and args.case_set in ['all', 'artifacts']:
            run_artifacts(cases, args.transports)
        record['verified'] = True
    except BaseException as error:
        record['error'] = repr(error)
        save()
        raise
    finally:
        if app is not None:
            try:
                require(pid_from_access_app(app) == pid and process_create_time(pid) == stamp, 'Owned process identity changed')
                app.Quit(win32com.client.constants.acQuitSaveNone)
            except BaseException as error:
                record['quit_error'] = repr(error)
        if cases is not None:
            cases.app = None
        cases = app = None
        gc.collect()
        if handle is not None:
            record['original_handle_exit'] = win32event.WaitForSingleObject(handle, 10000) == win32event.WAIT_OBJECT_0
            handle.Close()
        if endpoint is not None:
            endpoint.close()
        pythoncom.CoUninitialize()
        save()
    require(record.get('original_handle_exit'), 'Owned Access process has not exited')
    print(json.dumps({'verified': record.get('verified'), 'case_count': len(record['cases']),
                      'original_handle_exit': record['original_handle_exit']}))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--evidence', type=Path, required=True)
    parser.add_argument('--label', required=True)
    parser.add_argument('--client-source', type=Path, required=True)
    parser.add_argument('--transports', nargs='+', choices=['external', 'async', 'mcp'], default=['external', 'async', 'mcp'])
    parser.add_argument('--legacy-control', action='store_true')
    parser.add_argument('--case-set', choices=['all', 'core', 'artifacts', 'deletions'], default='all')
    args = parser.parse_args()
    require(Path(args.label).name == args.label, 'Label must be one path component')
    require(not (args.legacy_control and args.case_set in ['deletions', 'artifacts']), 'Legacy control requires core or all cases')
    require(not args.evidence.resolve().is_relative_to(Path(__file__).resolve().parents[1]), 'Keep evidence outside checkout')
    sys.path.insert(0, str(args.client_source.resolve()))
    run(args)

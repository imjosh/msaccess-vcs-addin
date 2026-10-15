"""Check source-built diagnostic writes in a fresh owned Access host on Windows."""
import argparse
import asyncio
import ctypes
import gc
import hashlib
import json
import math
import os
import re
import sys
import shutil
import subprocess
import time
import traceback
import zipfile
from pathlib import Path

import pythoncom
import win32api
import win32com.client
import win32event
import win32file
from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none, pid_from_access_app
from msaccess_vcs_mcp.cli import run_mcp_tool, _tool_result_text, parse_result_payload

REPO = Path(__file__).resolve().parents[1]
INSTALLED = Path.home()/'AppData/Roaming/MSAccessVCS/Version Control.accda'
EXCLUSIONS = ['-modTestConflicts.TestModuleImport_IndexesEachFileOnSharedInstance',
              '-modTestConflicts.TestModuleImportFast_IndexesEachFileOnSharedInstance']

def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def process_image(handle):
    size = ctypes.c_ulong(32768)
    buffer = ctypes.create_unicode_buffer(size.value)
    query = ctypes.WinDLL('kernel32', use_last_error=True).QueryFullProcessImageNameW
    query.argtypes = [ctypes.c_void_p, ctypes.c_ulong, ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_ulong)]
    query.restype = ctypes.c_int
    require(query(int(handle), 0, buffer, ctypes.byref(size)), 'Cannot identify recovery executable')
    return buffer.value


def validate_inputs(out, label):
    require(bool(re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{0,63}', label)), 'Invalid evidence label')
    require(label.upper() not in {'CON', 'PRN', 'AUX', 'NUL', *('COM'+str(i) for i in range(1,10)),
                                  *('LPT'+str(i) for i in range(1,10))}, 'Reserved evidence label')
    for path in (out, *out.parents):
        require(not path.is_symlink() and not path.is_junction(), 'Redirected evidence path')
    out = out.resolve()
    repo = REPO.resolve()
    require(not out.is_relative_to(repo) and not repo.is_relative_to(out),
            'Evidence must be outside the contribution and cannot contain it')
    require('.git' not in out.parts, 'Evidence cannot be in Git metadata')
    return out


def save_record(out, label, record):
    target = out/(label+'.json')
    temporary = out/(label+'.json.tmp')
    temporary.write_text(json.dumps(record, indent=2, default=str), encoding='utf-8')
    temporary.replace(target)


def qualification_checks(record):
    return {'no_error': not record.get('error') and not record.get('close_error'),
            'initial_compile': record.get('initial_compile') is True,
            'fresh_reopen_compile': record.get('fresh_reopen_compiled') is True,
            'components_preserved': record.get('components_preserved') is True,
            'library_preserved': record.get('library_preserved') is True,
            'references_preserved': record.get('references_preserved') is True,
            'initial_exit': record.get('initial_process', {}).get('original_handle_exit') is True,
            'reopen_exit': record.get('reopen_process', {}).get('original_handle_exit') is True,
            'empty_inventory': record.get('after_pids') == []}


def cleanup_target(out, label, record):
    out = validate_inputs(out, label)
    container = out/'hosts'/label
    folder = container/'repo'
    # Reject junctions and symlinks before resolving, including a redirected hosts root.
    for path in (out/'hosts', container, folder):
        require(not path.is_symlink() and not path.is_junction(), 'Redirected cleanup path')
    for path in container.rglob('*'):
        require(not path.is_symlink() and not path.is_junction()
                and path.resolve().is_relative_to(container.resolve()), 'Redirected descendant in cleanup tree')
    require(container.resolve().parent == (out/'hosts').resolve()
            and folder.resolve().parent == container.resolve()
            and folder.resolve() == Path(record['host']).resolve().parent,
            'Cleanup path escaped its evidence container')
    require(record.get('qualified') is True and not record.get('timed_out')
            and all(qualification_checks(record).values())
            and all(behavior_checks(record).values()), 'Unqualified host cannot be deleted')
    require(list_access_pids_or_none() == set(), 'Access remains before cleanup')
    archive = out/(label+'-host.zip')
    require(archive.is_file() and sha(archive) == record.get('archive_sha256'), 'Missing or changed host archive')
    return container, folder


def cleanup(out, label):
    record = json.loads((out/(label+'.json')).read_bytes())
    container, folder = cleanup_target(out, label, record)
    record['active_phase'] = 'remove_worktree'
    save_record(out, label, record)
    subprocess.run(['git','-C',str(REPO),'worktree','remove','--force',str(folder)], check=True, timeout=20)
    # Revalidate after Git removal before recursive deletion.
    container, folder = cleanup_target(out, label, record)
    shutil.rmtree(container)
    record['disposed'] = not container.exists()
    require(record['disposed'], 'Host disposal incomplete')
    record['active_phase'] = 'complete'
    save_record(out, label, record)


def recover_processes(record, seconds):
    """Use OS handles only. Never call COM while recovering a blocked worker."""
    results = []
    deadline = time.monotonic() + seconds
    for name in ('initial_process', 'reopen_process'):
        identity = record.get(name)
        if not identity:
            continue
        result = {'phase': name, **identity}
        results.append(result)
        handle = None
        try:
            pid, stamp = identity['pid'], identity['create_time']
            if process_create_time(pid) is None:
                result['outcome'] = 'already_exited'
                continue
            require(process_create_time(pid) == stamp, 'Recovery PID creation time changed')
            handle = win32api.OpenProcess(0x1000 | 0x100000 | 0x1, False, pid)
            require(process_create_time(pid) == stamp, 'Recovery PID reused while opening handle')
            image = process_image(handle)
            result['image'] = image
            require(Path(image).name.lower() == 'msaccess.exe', 'Recovery process is not Access')
            # Terminate through this validated handle, never by a later PID lookup.
            win32api.TerminateProcess(handle, 1)
            remaining = max(0, int((deadline-time.monotonic())*1000))
            result['exited'] = win32event.WaitForSingleObject(handle, remaining) == win32event.WAIT_OBJECT_0
            result['outcome'] = 'terminated' if result['exited'] else 'termination_wait_expired'
        except Exception as exc:
            result['outcome'] = 'recovery_refused_or_failed'
            result['error'] = repr(exc)
        finally:
            if handle is not None:
                handle.Close()
    return results


def supervise(command, out, label, seconds, recovery_seconds):
    """Bound the entire worker, including its COM shutdown, GC, and evidence work."""
    start = time.monotonic()
    with (out/(label+'-worker-console.txt')).open('a', encoding='utf-8') as log:
        worker = subprocess.Popen(command, stdout=log, stderr=subprocess.STDOUT)
        try:
            code = worker.wait(timeout=seconds)
            timed_out = False
        except subprocess.TimeoutExpired:
            timed_out = True
            kill_error = None
            try:
                worker.kill()
            except OSError as exc:
                kill_error = repr(exc)
            try:
                worker.wait(timeout=recovery_seconds)
            except subprocess.TimeoutExpired:
                pass
            code = worker.poll()
    record = json.loads((out/(label+'.json')).read_bytes())
    record.setdefault('supervision', []).append({'worker_pid': worker.pid, 'exit_code': code,
        'deadline_seconds': seconds, 'elapsed_seconds': time.monotonic()-start, 'timed_out': timed_out})
    if timed_out:
        record['worker_termination_error'] = kill_error
    if timed_out or code != 0:
        record.update(qualified=False, retained_failed_host=True, disposed=False,
                      timed_out=timed_out, failed_phase=record.get('failed_phase') or record.get('active_phase'))
        record['error'] = record.get('error') or ('worker_timeout' if timed_out else 'worker_failed')
        # Save failure before recovery so even a recovery failure leaves a receipt.
        save_record(out, label, record)
        # Recovery has its own disposable process and deadline as well.
        with (out/(label+'-worker-console.txt')).open('a', encoding='utf-8') as log:
            recovery = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '--_recover',
                str(out), label, str(recovery_seconds)], stdout=log, stderr=subprocess.STDOUT)
            try:
                recovery_code = recovery.wait(timeout=recovery_seconds)
            except subprocess.TimeoutExpired:
                recovery.kill()
                try:
                    recovery.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    pass
                recovery_code = None
        record = json.loads((out/(label+'.json')).read_bytes())
        record['recovery_worker'] = {'pid': recovery.pid, 'exit_code': recovery_code,
                                     'deadline_seconds': recovery_seconds}
        record['recovery_cleanup'] = 'host retained; recursive cleanup prohibited'
        save_record(out, label, record)
        raise RuntimeError(record['error'] + '; see ' + str(out/(label+'.json')))
    return record


def run(out, label, mode, filters, transport, timeout_seconds=300, recovery_seconds=10, cleanup_seconds=60):
    out = validate_inputs(out, label)
    for value in (timeout_seconds, recovery_seconds, cleanup_seconds):
        require(math.isfinite(value) and 0 < value <= 3600, 'Deadline must be finite and between 0 and 3600 seconds')
    require(list_access_pids_or_none() == set(), 'Native use must be serialized')
    out.mkdir(parents=True, exist_ok=True)
    require(not (out/'hosts').is_symlink() and not (out/'hosts').is_junction(), 'Redirected hosts root')
    require(not (out/'hosts'/label).exists(), 'Host label already exists')
    record = {'qualified': False, 'active_phase': 'worker_start', 'mode': mode,
              'filters': filters, 'transport': transport, 'supervisor_pid': os.getpid()}
    # Exclusive reservation also protects historical receipts after disposed hosts.
    with (out/(label+'.json')).open('x', encoding='utf-8') as file:
        json.dump(record, file, indent=2)
    config = out/(label+'-request.json')
    with config.open('x', encoding='utf-8') as file:
        json.dump({'out': str(out), 'label': label, 'mode': mode, 'filters': filters, 'transport': transport}, file)
    command = [sys.executable, str(Path(__file__).resolve()), '--_worker', str(config)]
    record = supervise(command, out, label, timeout_seconds, recovery_seconds)
    require(record.get('qualified') is True, 'Worker did not qualify')
    supervise([sys.executable, str(Path(__file__).resolve()), '--_cleanup', str(config)],
              out, label, cleanup_seconds, recovery_seconds)
    return json.loads((out/(label+'.json')).read_bytes())

def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def git(repo, *args): return subprocess.check_output(['git', '-C', str(repo), *args], text=True).strip()

def behavior_checks(record):
    if record['mode'] == 'suite':
        tests = record.get('tests', {})
        summary = tests.get('summary', {})
        cases = tests.get('tests', {})
        required = ['TestTryAppendToFileAppendsLines', 'TestTryAppendToFileFailsQuietly', 'TestTryAppendToFileLocked',
                    'TestTryAppendToFileWithLiveErrorWrites', 'TestTryAppendToFileWithLiveErrorFailsQuietly']
        checks = {'passing_assertions': summary.get('passed', 0) > 0,
                'no_failed_assertions': summary.get('failed') == 0,
                'no_errored_tests': summary.get('errored') == 0,
                'all_passed': tests.get('allPassed') is True,
                'exclusions_applied': all(f[1:] not in cases for f in record['filters'] if f.startswith('-'))}
        for name in required:
            case = cases.get('modTestFileAccess.'+name, {})
            assertions = case.get('assertions', [])
            checks[name] = (case.get('status') == 'PASSED' and bool(assertions)
                            and all(a.get('passed') is True and 'SKIP:' not in str(a) for a in assertions))
        return checks
    probe = record.get('probe', {})
    checks = {'callback_active': probe.get('active') is True,
              'callback_released': probe.get('callback_released') is True,
              'failed_registration_released': probe.get('failed_registration_released') is True,
              'registration_error_consumed': probe.get('registration_error_consumed') is True,
              'failed_registration_fields_cleared': probe.get('failed_registration_fields_cleared') is True,
              'before_root': probe.get('status_before') == 0,
              'normal_mode': probe.get('mode_before') == 0,
              'no_journal_entries': probe.get('journal_delta') == 0,
              'no_errors': probe.get('error_delta') == 0,
              'level_preserved': probe.get('level_unchanged') is True,
              'status_preserved': probe.get('status_unchanged') is True,
              'mode_preserved': probe.get('mode_unchanged') is True,
              'original_mode_restored': probe.get('original_mode_restored') is True,
              'no_probe_error': 'probe_error' not in probe,
              'public_result_preserved': record.get('public_version_after') == record.get('version')}
    if record['mode'] in ['blocker', 'locked']:
        checks['bytes_preserved'] = record.get('original_bytes_preserved') is True
    else:
        checks['one_bom'] = record.get('one_bom') is True
        checks['marker_present'] = record.get('marker_present') is True
        checks['registration_error_details_preserved'] = record.get('registration_error_details_preserved') is True
    return checks

def native_run(out, label, mode, filters, transport):
    out = out.resolve()
    validate_inputs(out, label)
    record = json.loads((out/(label+'.json')).read_bytes())
    record.update(worker_pid=os.getpid(), active_phase='prepare_host')
    save_record(out, label, record)
    out.mkdir(parents=True, exist_ok=True)
    container = out/'hosts'/label
    require(not container.exists(), 'Host label already exists')
    folder = container/'repo'
    folder.parent.mkdir(parents=True)
    require(list_access_pids_or_none() == set(), 'Native use must be serialized')
    pin = git(REPO, 'rev-parse', 'HEAD')
    subprocess.run(['git', '-C', str(REPO), 'worktree', 'add', '--detach', str(folder), pin], check=True)
    host = folder/'Version Control.accda'
    source = Path(str(host)+'.src')
    shutil.copyfile(REPO/'Version Control.accda', host)
    shutil.copytree(REPO/'Version Control.accda.src', source, dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns('logs', 'mcp', 'test-results'))
    # Existing command-bar tests also resolve fixtures one directory above CodeProject.Path.
    shutil.copytree(folder/'Testing/Fixtures/menus/regression', container/'Testing/Fixtures/menus/regression')
    record.update({'qualified': False, 'active_phase': 'prepared', 'pin': pin, 'mode': mode, 'filters': filters, 'transport': transport,
              'host': str(host), 'binary_sha256': sha(host), 'runner_sha256': sha(INSTALLED),
              'source_manifest': {p.relative_to(source).as_posix():sha(p) for p in source.rglob('*') if p.is_file()}})
    app = handle = project = module = lock = None
    phase = None

    def save(): save_record(out, label, record)
    def checkpoint(name):
        record['active_phase'] = name
        save()
    def api(method, *args):
        checkpoint('API.' + method)
        value = app.Run(str(INSTALLED.with_suffix(''))+'.API', method, *args)
        return value[0] if isinstance(value, tuple) else value
    def open_owned(name):
        nonlocal app, handle, phase
        require(list_access_pids_or_none() == set(), 'Competing Access process')
        checkpoint(name + '.DispatchEx')
        app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Access.Application'))
        pid = pid_from_access_app(app)
        stamp = process_create_time(pid)
        require(pid and stamp, 'Missing process identity')
        handle = win32api.OpenProcess(0x1000|0x100000, False, pid)
        phase = name
        record[name] = {'pid':pid, 'create_time':stamp}
        checkpoint(name + '.OpenCurrentDatabase')
        app.OpenCurrentDatabase(str(host))
        app.Visible = True
        require(Path(app.CurrentProject.FullName).resolve() == host.resolve(), 'Unexpected native host')
    def compile_project():
        nonlocal project
        checkpoint('compile_project')
        project = next(p for p in app.VBE.VBProjects if Path(p.FileName).resolve() == host.resolve())
        app.VBE.ActiveVBProject = project
        app.DoCmd.RunCommand(win32com.client.constants.acCmdCompileAndSaveAllModules)
        project = None
        return bool(app.IsCompiled)
    def inspect(target=host):
        nonlocal project, module
        checkpoint('inspect_components')
        project = next(p for p in app.VBE.VBProjects if Path(p.FileName).resolve() == target.resolve())
        result = {}
        for component in project.VBComponents:
            module = component.CodeModule
            result[component.Name] = hashlib.sha256(str(module.Lines(1,module.CountOfLines)).encode()).hexdigest() if module.CountOfLines else ''
            module = None
        project = None
        return result
    def close_owned():
        nonlocal app, handle, project, module
        project = module = None
        checkpoint(phase + '.close')
        require(pid_from_access_app(app) == record[phase]['pid'], 'Close PID mismatch')
        require(process_create_time(record[phase]['pid']) == record[phase]['create_time'], 'Close creation-time mismatch')
        require(Path(app.CurrentProject.FullName).resolve() == host.resolve(), 'Unexpected native host')
        app.VBE.MainWindow.Visible = False
        app.Quit(win32com.client.constants.acQuitSaveNone)
        app = None
        gc.collect()
        record[phase]['original_handle_exit'] = win32event.WaitForSingleObject(handle,5000) == win32event.WAIT_OBJECT_0
        handle.Close()
        handle = None
        require(record[phase]['original_handle_exit'], 'Original Access handle did not exit')

    pythoncom.CoInitialize()
    try:
        open_owned('initial_process')
        record['version'] = api('GetVCSVersion')
        record['access_version'] = app.Version
        record['references'] = [{'name':r.Name,'path':r.FullPath,'broken':r.IsBroken} for r in app.References]
        record['library_before'] = inspect(INSTALLED)
        record['initial_compile'] = compile_project()
        require(record['initial_compile'], 'Initial compilation failed')
        if mode == 'suite':
            if transport == 'callback':
                args = {'database_path':str(host), 'timeout_seconds':300, 'filter':','.join(filters)}
                def progress(value,total,message):
                    if message: print(message, flush=True)
                checkpoint('callback.RunTestsHeadless')
                raw = asyncio.run(run_mcp_tool('vcs_run_tests',args,on_progress=progress))
                record['tests'] = parse_result_payload(_tool_result_text(raw))
                require(list_access_pids_or_none() == {record[phase]['pid']}, 'Callback changed Access process inventory')
                require(process_create_time(record[phase]['pid']) == record[phase]['create_time'], 'Callback process identity changed')
            else:
                app.Run(str(INSTALLED.with_suffix(''))+'.SetInteractionMode',1)
                record['tests'] = json.loads(api('RunTestsHeadless',*filters))
            record['skip_assertions'] = [a for v in record['tests'].get('tests',{}).values() for a in v.get('assertions',[]) if 'SKIP:' in str(a)]
            record['end_session'] = api('EndSession')
            print(json.dumps(record['tests'].get('summary',{})), flush=True)
        else:
            checkpoint('import_probe')
            project = next(p for p in app.VBE.VBProjects if Path(p.FileName).resolve() == host.resolve())
            module = project.VBComponents.Import(str(folder/'Testing/Fixtures/modules/vcs_diagnostic_write_probe.bas'))
            record['probe_module_name'] = module.Name
            require(module.Name == 'vcs_diagnostic_write_probe', 'Unexpected probe module')
            app.DoCmd.Save(win32com.client.constants.acModule,module.Name)
            module = project = None
            record['probe_compile'] = compile_project()
            require(record['probe_compile'], 'Probe import did not establish compilation')
            logs = source/'logs'
            if logs.exists():
                shutil.move(str(logs),str(source/'probe-original-logs'))
            path = logs/'MCP_Debug.log'
            if mode == 'blocker':
                logs.write_bytes(b'blocked directory\x00\xff')
                before = logs.read_bytes()
            elif mode == 'locked':
                logs.mkdir()
                path.write_bytes(b'\xef\xbb\xbforiginal\r\n')
                before = path.read_bytes()
                lock = win32file.CreateFile(str(path),0x80000000|0x40000000,0,None,3,0,None)
            callback = json.dumps({'callback_url':'http://127.0.0.1:1/unused', 'operation_id':label})
            start = time.monotonic()
            checkpoint('DiagnosticWriteProbe')
            value = app.Run(str(host.with_suffix(''))+'.DiagnosticWriteProbe',callback,'diagnostic probe one')
            record['elapsed_seconds'] = time.monotonic()-start
            record['probe'] = json.loads(value[0] if isinstance(value,tuple) else value)
            if lock: lock.Close(); lock = None
            if mode in ['blocker','locked']:
                record['original_bytes_preserved'] = (logs if mode=='blocker' else path).read_bytes() == before
            else:
                raw = path.read_bytes()
                record['log_hex'] = raw.hex()
                record['one_bom'] = raw.startswith(b'\xef\xbb\xbf') and raw.count(b'\xef\xbb\xbf') == 1
                record['marker_present'] = b'diagnostic probe one\r\n' in raw
                record['registration_error_details_preserved'] = b'clsMCP.RegisterCallback: ERROR - Error parsing JSON:' in raw
            record['public_version_after'] = api('GetVCSVersion')
        checkpoint('behavior_complete')
        record['before_close_components'] = inspect()
        record['library_after'] = inspect(INSTALLED)
        record['library_preserved'] = record['library_before'] == record['library_after']
        close_owned()
        open_owned('reopen_process')
        record['reopen_components'] = inspect()
        record['reopen_references'] = [{'name':r.Name,'path':r.FullPath,'broken':r.IsBroken} for r in app.References]
        record['references_preserved'] = record['references'] == record['reopen_references']
        record['fresh_reopen_compiled'] = compile_project()
        record['components_preserved'] = record['before_close_components'] == record['reopen_components']
        close_owned()
    except Exception as exc:
        record['error'] = repr(exc)
        record['traceback'] = traceback.format_exc()
        record['failed_phase'] = record['active_phase']
        save()
    finally:
        if lock: lock.Close()
        project = module = None
        if app is not None and handle is not None:
            try: close_owned()
            except Exception as exc: record['close_error'] = repr(exc)
        app = None
        gc.collect()
        pythoncom.CoUninitialize()
        record['final_inventory_samples'] = []
        for _ in range(5):
            remaining = list_access_pids_or_none()
            record['final_inventory_samples'].append(None if remaining is None else sorted(remaining))
            if remaining == set(): break
            time.sleep(.5)
        record['after_pids'] = None if remaining is None else sorted(remaining)
        save()
    record['behavior_checks'] = behavior_checks(record)
    record['qualification_checks'] = qualification_checks(record)
    record['qualified'] = all(record['behavior_checks'].values()) and all(record['qualification_checks'].values())
    save()
    if record['qualified']:
        checkpoint('archive_host')
        archive = out/(label+'-host.zip')
        with zipfile.ZipFile(archive,'x',zipfile.ZIP_DEFLATED) as z:
            for p in container.rglob('*'):
                if p.is_file(): z.write(p,p.relative_to(container))
        record['archive_sha256'] = sha(archive)
        record['git_status_after'] = git(folder,'status','--porcelain')
    else:
        record['retained_failed_host'] = True
    checkpoint('native_complete')
    print(json.dumps({k:v for k,v in record.items() if k not in ['source_manifest','tests','before_close_components','reopen_components','references']},indent=2),flush=True)
    require(record['qualified'], 'Native qualification failed; host retained')
    return record


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--_recover':
        out, label, seconds = Path(sys.argv[2]), sys.argv[3], float(sys.argv[4])
        record = json.loads((out/(label+'.json')).read_bytes())
        record['recovery'] = recover_processes(record, seconds)
        remaining = list_access_pids_or_none()
        record['recovery_after_pids'] = None if remaining is None else sorted(remaining)
        save_record(out, label, record)
    elif len(sys.argv) == 3 and sys.argv[1] in ('--_worker', '--_cleanup'):
        config = json.loads(Path(sys.argv[2]).read_bytes())
        config['out'] = Path(config['out'])
        try:
            if sys.argv[1] == '--_worker':
                native_run(**config)
            else:
                cleanup(config['out'], config['label'])
        except Exception as exc:
            record = json.loads((config['out']/(config['label']+'.json')).read_bytes())
            record.update(qualified=False, error=repr(exc), traceback=traceback.format_exc(), retained_failed_host=True)
            save_record(config['out'], config['label'], record)
            raise
    else:
        parser = argparse.ArgumentParser(description=__doc__)
        parser.add_argument('out',type=Path)
        parser.add_argument('label')
        parser.add_argument('--mode',choices=['suite','normal','blocker','locked'],default='suite')
        parser.add_argument('--filter',default='modTestFileAccess')
        parser.add_argument('--broad',action='store_true',help='Run all tests except the two unsafe base import/deletion tests')
        parser.add_argument('--transport',choices=['sync','callback'],default='sync')
        parser.add_argument('--timeout-seconds',type=float,default=300)
        parser.add_argument('--recovery-seconds',type=float,default=10)
        parser.add_argument('--cleanup-seconds',type=float,default=60)
        args = parser.parse_args()
        run(args.out,args.label,args.mode,EXCLUSIONS if args.broad else (args.filter.split(',') if args.filter else []),
            args.transport,args.timeout_seconds,args.recovery_seconds,args.cleanup_seconds)

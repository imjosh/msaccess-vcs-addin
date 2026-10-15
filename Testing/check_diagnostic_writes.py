"""Check source-built diagnostic writes in a fresh owned Access host on Windows."""
import argparse
import asyncio
import gc
import hashlib
import json
import shutil
import subprocess
import time
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

def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def git(repo, *args): return subprocess.check_output(['git', '-C', str(repo), *args], text=True).strip()

def run(out, label, mode, filters, transport):
    out = out.resolve()
    assert not out.is_relative_to(REPO.resolve()), 'Keep evidence outside the contribution'
    assert Path(label).name == label and label not in ['.', '..']
    out.mkdir(parents=True, exist_ok=True)
    container = out/'hosts'/label
    assert not container.exists()
    folder = container/'repo'
    folder.parent.mkdir(parents=True)
    assert list_access_pids_or_none() == set(), 'Native use must be serialized'
    pin = git(REPO, 'rev-parse', 'HEAD')
    subprocess.run(['git', '-C', str(REPO), 'worktree', 'add', '--detach', str(folder), pin], check=True)
    host = folder/'Version Control.accda'
    source = Path(str(host)+'.src')
    shutil.copyfile(REPO/'Version Control.accda', host)
    shutil.copytree(REPO/'Version Control.accda.src', source, dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns('logs', 'mcp', 'test-results'))
    # Existing command-bar tests also resolve fixtures one directory above CodeProject.Path.
    shutil.copytree(folder/'Testing/Fixtures/menus/regression', container/'Testing/Fixtures/menus/regression')
    record = {'pin': pin, 'mode': mode, 'filters': filters, 'transport': transport,
              'host': str(host), 'binary_sha256': sha(host), 'runner_sha256': sha(INSTALLED),
              'source_manifest': {p.relative_to(source).as_posix():sha(p) for p in source.rglob('*') if p.is_file()}}
    app = handle = project = module = lock = None
    phase = None

    def save(): (out/(label+'.json')).write_text(json.dumps(record, indent=2, default=str), encoding='utf-8')
    def api(method, *args):
        value = app.Run(str(INSTALLED.with_suffix(''))+'.API', method, *args)
        return value[0] if isinstance(value, tuple) else value
    def open_owned(name):
        nonlocal app, handle, phase
        assert list_access_pids_or_none() == set()
        app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Access.Application'))
        pid = pid_from_access_app(app)
        stamp = process_create_time(pid)
        assert pid and stamp
        handle = win32api.OpenProcess(0x1000|0x100000, False, pid)
        phase = name
        record[name] = {'pid':pid, 'create_time':stamp}
        save()
        app.OpenCurrentDatabase(str(host))
        app.Visible = True
        assert Path(app.CurrentProject.FullName).resolve() == host.resolve()
    def compile_project():
        nonlocal project
        project = next(p for p in app.VBE.VBProjects if Path(p.FileName).resolve() == host.resolve())
        app.VBE.ActiveVBProject = project
        app.DoCmd.RunCommand(win32com.client.constants.acCmdCompileAndSaveAllModules)
        project = None
        return bool(app.IsCompiled)
    def inspect():
        nonlocal project, module
        project = next(p for p in app.VBE.VBProjects if Path(p.FileName).resolve() == host.resolve())
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
        assert pid_from_access_app(app) == record[phase]['pid']
        assert process_create_time(record[phase]['pid']) == record[phase]['create_time']
        assert Path(app.CurrentProject.FullName).resolve() == host.resolve()
        app.VBE.MainWindow.Visible = False
        app.Quit(2)
        app = None
        gc.collect()
        record[phase]['original_handle_exit'] = win32event.WaitForSingleObject(handle,5000) == win32event.WAIT_OBJECT_0
        handle.Close()
        handle = None
        assert record[phase]['original_handle_exit']

    pythoncom.CoInitialize()
    try:
        open_owned('initial_process')
        record['version'] = api('GetVCSVersion')
        record['access_version'] = app.Version
        record['references'] = [{'name':r.Name,'path':r.FullPath,'broken':r.IsBroken} for r in app.References]
        record['initial_compile'] = compile_project()
        assert record['initial_compile']
        if mode == 'suite':
            if transport == 'callback':
                args = {'database_path':str(host), 'timeout_seconds':300, 'filter':','.join(filters)}
                def progress(value,total,message):
                    if message: print(message, flush=True)
                raw = asyncio.run(run_mcp_tool('vcs_run_tests',args,on_progress=progress))
                record['tests'] = parse_result_payload(_tool_result_text(raw))
            else:
                app.Run(str(INSTALLED.with_suffix(''))+'.SetInteractionMode',1)
                record['tests'] = json.loads(api('RunTestsHeadless',*filters))
            record['skip_assertions'] = [a for v in record['tests'].get('tests',{}).values() for a in v.get('assertions',[]) if 'SKIP:' in str(a)]
            record['end_session'] = api('EndSession')
            print(json.dumps(record['tests'].get('summary',{})), flush=True)
        else:
            project = next(p for p in app.VBE.VBProjects if Path(p.FileName).resolve() == host.resolve())
            module = project.VBComponents.Import(str(folder/'Testing/Fixtures/modules/vcs_diagnostic_write_probe.bas'))
            assert module.Name == 'vcs_diagnostic_write_probe'
            app.DoCmd.Save(5,module.Name)
            module = project = None
            assert compile_project()
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
            record['public_version_after'] = api('GetVCSVersion')
        record['before_close_components'] = inspect()
        close_owned()
        open_owned('reopen_process')
        record['reopen_components'] = inspect()
        record['fresh_reopen_compiled'] = compile_project()
        record['components_preserved'] = record['before_close_components'] == record['reopen_components']
        close_owned()
    except Exception as exc:
        record['error'] = repr(exc)
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
    assert not record.get('error') and not record.get('close_error'), record.get('error') or record.get('close_error')
    assert record['fresh_reopen_compiled'] and record['components_preserved'] and list_access_pids_or_none() == set()
    archive = out/(label+'-host.zip')
    with zipfile.ZipFile(archive,'x',zipfile.ZIP_DEFLATED) as z:
        for p in container.rglob('*'):
            if p.is_file(): z.write(p,p.relative_to(container))
    record['archive_sha256'] = sha(archive)
    record['git_status_after'] = git(folder,'status','--porcelain')
    target = folder.resolve()
    assert target.parent == container.resolve() and container.resolve().parent == (out/'hosts').resolve()
    subprocess.run(['git','-C',str(REPO),'worktree','remove','--force',str(target)],check=True)
    shutil.rmtree(container)
    record['disposed'] = not container.exists()
    save()
    print(json.dumps({k:v for k,v in record.items() if k not in ['source_manifest','tests','before_close_components','reopen_components','references']},indent=2),flush=True)
    return record

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('out',type=Path)
    parser.add_argument('label')
    parser.add_argument('--mode',choices=['suite','normal','blocker','locked'],default='suite')
    parser.add_argument('--filter',default='modTestFileAccess')
    parser.add_argument('--broad',action='store_true',help='Run all tests except the two unsafe base import/deletion tests')
    parser.add_argument('--transport',choices=['sync','callback'],default='sync')
    args = parser.parse_args()
    run(args.out,args.label,args.mode,EXCLUSIONS if args.broad else args.filter.split(','),args.transport)

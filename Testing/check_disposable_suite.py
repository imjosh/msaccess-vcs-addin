"""Run MCP tests in one owned disposable host, then validate fresh reopen.

Requires the Windows MCP Python environment and a matching source-built binary.
Never replaces modules in a loaded library or deletes a live host.
"""
import argparse
import asyncio
import gc
import hashlib
import json
import os
import shutil
import subprocess
import time
import zipfile
from pathlib import Path

import pythoncom
import win32api
import win32com.client
import win32event
from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none, pid_from_access_app
from msaccess_vcs_mcp.cli import run_mcp_tool, _tool_result_text, parse_result_payload

REPO=Path(__file__).resolve().parents[1]
OUT=None
INSTALLED=Path.home()/'AppData/Roaming/MSAccessVCS/Version Control.accda'
# Full-suite quits can take well over five seconds to finish releasing the process.
EXIT_WAIT_MS=60000
INVENTORY_SETTLE_MS=30000
FIXTURES=['vcs_test_import_alpha','vcs_test_import_beta',
          'vcs_test_import_fast_alpha','vcs_test_import_fast_beta']

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def git(repo,*args):return subprocess.check_output(['git','-C',str(repo),*args],text=True).strip()

def run(label,binary,pin,filters,fault=None,sync=False):
    assert OUT is not None and OUT.is_dir()
    assert not OUT.resolve().is_relative_to(REPO.resolve())
    assert Path(label).name==label and label not in ['.','..']
    container=OUT/'hosts'/label
    assert not container.exists()
    folder=container/'repo'
    folder.parent.mkdir(parents=True,exist_ok=False)
    assert not folder.exists()
    assert list_access_pids_or_none()==set(), 'Exclusive native use required'
    subprocess.run(['git','-C',str(REPO),'worktree','add','--detach',str(folder),pin],check=True)
    host=folder/'Version Control.accda'
    shutil.copyfile(binary,host)
    src=Path(str(host)+'.src')
    shutil.copytree(Path(str(binary)+'.src'),src,dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns('logs','mcp','test-results'))
    # Older command-bar tests resolve these fixtures beside CodeProject.Path.
    legacy=container/'Testing/Fixtures/menus/regression'
    shutil.copytree(folder/'Testing/Fixtures/menus/regression',legacy)
    # Controls modify only the idle disposable snapshot's fixture source.
    if fault=='assertion':
        p=folder/'Testing/Fixtures/modules/vcs_test_import_alpha.bas'
        p.write_bytes(p.read_bytes().replace(b'vcs_test_import_alpha',b'vcs_test_import_assert_control'))
    if fault=='import':
        p=folder/'Testing/Fixtures/modules/vcs_test_import_beta.bas'
        p.write_bytes(b'Attribute VB_Name = "vcs_test_import_beta"\r\n<<<<<<< ours\r\nOption Explicit\r\n=======\r\nOption Compare Database\r\n>>>>>>> theirs\r\n')
    data={'host':str(host),'pin':pin,'resolved_pin':git(folder,'rev-parse','HEAD'),'filters':filters,'fault':fault,'transport':'sync-json' if sync else 'mcp',
          'binary_input_sha256':sha(host),'installed_input_sha256':sha(INSTALLED),
          'source_manifest':{p.relative_to(folder).as_posix():sha(p) for p in src.rglob('*')
                             if p.is_file() and not any(x in p.parts for x in ['logs','mcp','test-results'])},
          'client_version':__import__('msaccess_vcs_mcp').__version__}
    data['legacy_fixture_context']={p.relative_to(container).as_posix():sha(p) for p in legacy.rglob('*') if p.is_file()}
    app=handle=project=module=None
    current_phase=None  # Phase whose Access instance this run still holds, if any.
    opened_phases=[]
    def save():
        (OUT/(label+'.json')).write_text(json.dumps(data,indent=2,default=str),encoding='utf-8')
    def open_owned(phase):
        """Create the owned instance, recording ownership as each step succeeds.

        The DispatchEx reference is kept before any wrapping, and the PID and create
        time are recorded before the process handle is opened, so cleanup can still
        quit the instance when a later step fails.
        """
        nonlocal app,handle,current_phase
        assert list_access_pids_or_none()==set()
        app=handle=None
        current_phase=phase
        opened_phases.append(phase)
        info=data[phase]={}
        try:
            app=win32com.client.DispatchEx('Access.Application')
            app=win32com.client.gencache.EnsureDispatch(app)
            pid=pid_from_access_app(app)
            info['pid']=pid
            assert pid,'Owned Access PID unavailable'
            info['create_time']=process_create_time(pid)
            assert info['create_time'],'Owned Access create time unavailable'
            handle=win32api.OpenProcess(0x1000|0x100000,False,pid)
            app.OpenCurrentDatabase(str(host));app.Visible=True
            assert Path(app.CurrentProject.FullName).resolve()==host.resolve()
        except Exception as exc:
            info['init_error']=repr(exc)
            raise
    def identity_mismatch(info):
        """Return what contradicts the recorded identity of the held instance, if anything.

        A check that cannot run is recorded but does not block quitting: the held
        reference is the instance this phase created with DispatchEx.
        """
        pid=info.get('pid')
        if not pid:return None
        try:
            current_pid=pid_from_access_app(app)
            if current_pid is None:info['pid_check_unavailable']=True
            elif current_pid!=pid:return 'pid'
        except Exception as exc:info['pid_check_error']=repr(exc)
        if info.get('create_time'):
            try:
                current_create_time=process_create_time(pid)
                if current_create_time is None:info['create_time_check_unavailable']=True
                elif current_create_time!=info['create_time']:return 'create_time'
            except Exception as exc:info['create_time_check_error']=repr(exc)
        return None
    def wait_for_no_access(timeout_ms):
        deadline=time.monotonic()+timeout_ms/1000
        while True:
            if list_access_pids_or_none()==set():return True
            if time.monotonic()>=deadline:return False
            time.sleep(0.5)
    def shutdown_owned(phase):
        """Quit the instance this phase created and record how it ended. Never raises.

        Only the held COM instance is quit, and not when its recorded PID or create
        time contradicts it. Database, identity, quit, exit-timeout and handle failures
        are recorded under separate keys, and a failed check never prevents the quit.
        """
        nonlocal app,handle,project,module,current_phase
        info=data.setdefault(phase,{})
        project=module=None  # Release code-module references before quitting.
        gc.collect()
        quit_attempted=False
        if app is not None:
            try:info['database_at_close']=str(Path(app.CurrentProject.FullName).resolve())
            except Exception as exc:info['database_check_error']=repr(exc)
            mismatch=identity_mismatch(info)
            if mismatch:
                info['identity_mismatch']=mismatch
            else:
                quit_attempted=True
                try:app.Quit(2)
                except Exception as exc:info['quit_error']=repr(exc)
        app=None;gc.collect()
        wait_ms=EXIT_WAIT_MS if quit_attempted else 0
        info['exit_wait_ms']=wait_ms
        exited=False
        if handle is not None:
            try:
                result=win32event.WaitForSingleObject(handle,wait_ms)
                exited=result==win32event.WAIT_OBJECT_0
                info['original_handle_exit']=exited
                info['exit_timeout']=result==win32event.WAIT_TIMEOUT
            except Exception as exc:info['exit_wait_error']=repr(exc)
            try:handle.Close()
            except Exception as exc:info['handle_close_error']=repr(exc)
            handle=None
        else:
            # No handle was opened: the exclusive-use precondition makes an empty process list the evidence.
            exited=wait_for_no_access(wait_ms)
            info['process_list_exit']=exited
        # A confirmed original-process exit can precede the disappearance of
        # transient Access helpers from the process inventory. Observe a bounded
        # settlement window; never terminate extra processes or accept an unknown
        # inventory. Both original handles and retained-fixture checks stay required.
        samples=[]
        settle_start=time.monotonic()
        settle_deadline=settle_start+(INVENTORY_SETTLE_MS/1000 if exited else 0)
        empty_samples=0
        inventory_unknown=False
        while True:
            pids=list_access_pids_or_none()
            samples.append({'elapsed_ms':round((time.monotonic()-settle_start)*1000),
                            'pids':sorted(pids) if pids is not None else None})
            if pids is None:
                inventory_unknown=True
                break
            empty_samples=empty_samples+1 if pids==set() else 0
            if empty_samples>=2 or time.monotonic()>=settle_deadline:
                break
            time.sleep(0.25)
        info['inventory_settle_ms']=INVENTORY_SETTLE_MS if exited else 0
        info['inventory_samples']=samples
        info['pids_after_exit']=sorted(pids) if pids is not None else None
        info['closure_confirmed']=bool(exited and not inventory_unknown and empty_samples>=2)
        current_phase=None
    def close_owned(phase):
        shutdown_owned(phase)
        info=data[phase]
        assert not info.get('identity_mismatch'),'Owned Access identity changed'
        assert Path(info.get('database_at_close',''))==host.resolve(),'Owned Access database changed'
        assert info['closure_confirmed'],'Owned Access closure not confirmed'
    def api(method,*args):
        value=app.Run(str(INSTALLED.with_suffix(''))+'.API',method,*args)
        return value[0] if isinstance(value,tuple) else value
    def snapshot():
        value=app.Run(str(host.with_suffix(''))+'.DisposableHostSnapshot','')
        return json.loads(value[0] if isinstance(value,tuple) else value)
    def inspect(phase):
        nonlocal project,module
        project=next(p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve()==host.resolve())
        components={}
        for name in FIXTURES+['modTestConflicts','modTestClassInvocation','clsTestTargetLocalInvocation','modTestAssert']:
            try:
                module=project.VBComponents(name).CodeModule
                body=str(module.Lines(1,module.CountOfLines))
                components[name]={'lines':module.CountOfLines,'body_sha256':hashlib.sha256(body.encode()).hexdigest()}
                if name=='modTestAssert':components[name]['body']=body
            except Exception as exc:components[name]={'error':repr(exc)}
            module=None
        data[phase]={'components':components,'fixture_names':sorted(c.Name for c in project.VBComponents if c.Name.startswith('vcs_test_import_')),
                     'projects':[str(p.FileName) for p in app.VBE.VBProjects],
                     'references':[{'name':r.Name,'path':r.FullPath,'broken':r.IsBroken} for r in app.References]}
        library=next((p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve()==INSTALLED.resolve()),None)
        if library is not None:
            hashes={}
            for component in library.VBComponents:
                cm=component.CodeModule
                hashes[component.Name]=hashlib.sha256(str(cm.Lines(1,cm.CountOfLines)).encode()).hexdigest() if cm.CountOfLines else ''
                cm=None
            data[phase]['library_components']=hashes
            library=None
        project=None
    pythoncom.CoInitialize()
    try:
        open_owned('initial_process')
        data['installed_version']=api('GetVCSVersion')
        project=next(p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve()==host.resolve())
        module=project.VBComponents.Import(str(folder/'Testing/Fixtures/modules/vcs_disposable_host_probe.bas'))
        assert module.Name=='vcs_disposable_host_probe'
        app.DoCmd.Save(5,'vcs_disposable_host_probe')
        module=project=None
        # The import tests refuse any host without this explicit designation.
        value=app.Run(str(host.with_suffix(''))+'.DesignateDisposableHost','')
        data['designation']=value[0] if isinstance(value,tuple) else value
        assert Path(data['designation']).resolve()==host.resolve(),'host designation mismatch'
        data['initial_compile']=api('CompileVBA',True)
        data['initial_is_compiled']=bool(app.IsCompiled)
        data['state_before']=snapshot()
        inspect('initial_project')
        if fault in ['class-error','legacy-invoker']:
            project=next(p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve()==host.resolve())
            module=project.VBComponents('modTestAssert').CodeModule
            try:
                line=module.ProcStartLine('TestClassInvoke',0)
                module.DeleteLines(line,module.ProcCountLines('TestClassInvoke',0))
            except Exception:pass
            if fault=='legacy-invoker':
                text='Public Sub TestClassInvoke(ByVal obj As Object, ByVal strProcName As String)\r\nCallByName obj, strProcName, VbMethod\r\nEnd Sub\r\n'
            else:
                text='Public Function TestClassInvoke(ByVal obj As Object, ByVal strProcName As String) As String\r\n\' TestClassInvoke helper version 2\r\nOn Error GoTo Failed\r\nCallByName obj, "NoSuchTestMethod", VbMethod\r\nExit Function\r\nFailed:\r\nTestClassInvoke = CStr(Err.Number) & vbTab & Err.Description\r\nEnd Function\r\n'
            module.InsertLines(module.CountOfLines+1,text)
            data['control_helper']=text
            app.DoCmd.Save(5,'modTestAssert')
            module=project=None
        args={'database_path':str(host),'timeout_seconds':300}
        if filters:args['filter']=filters
        data['mcp_arguments']=args;save()
        def progress(value,total,message):
            if message:print(message,flush=True)
        if sync:
            option_result=api('SetOption','DefaultTestFilter',filters)
            data['sync_filter_setup']=option_result
            setup=json.loads(option_result) if isinstance(option_result,str) else option_result
            if not isinstance(setup,dict) or setup.get('success') is not True:
                raise RuntimeError('Synchronous filter setup failed: '+repr(option_result))
            data['sync_policy']='block'
            data['sync_raw_result']=api('RunFilteredTests','block')
            data['tests']=json.loads(data['sync_raw_result'])
            data['tests']['success']=data['tests'].get('allPassed',False)
        else:
            result=asyncio.run(run_mcp_tool('vcs_run_tests',args,on_progress=progress))
            data['tests']=parse_result_payload(_tool_result_text(result))
        print(json.dumps(data['tests'].get('summary',{})),flush=True)
        data['skip_assertions']=[{'test':key,'assertion':a} for key,value in data['tests'].get('tests',{}).items()
                                 for a in value.get('assertions',[]) if 'SKIP:' in str(a)]
        data['legacy_fixture_context_unchanged']=all(sha(container/name)==value for name,value in data['legacy_fixture_context'].items())
        marker=folder/'invocation-teardown.txt'
        data['teardown_markers']=marker.read_text().splitlines() if marker.exists() else []
        data['state_before_end_session']=snapshot()
        data['end_session']=api('EndSession')
        data['state_after']=snapshot()
        inspect('after_project')
        data['is_compiled_after_run']=bool(app.IsCompiled)
        for location,name in [(src,'host'),(Path(str(INSTALLED)+'.src'),'runner')]:
            for part in ['logs','test-results','mcp']:
                p=location/part
                if p.exists():shutil.copytree(p,OUT/(label+'-'+name)/part)
        close_owned('initial_process');save()
        open_owned('reopen_process')
        inspect('reopened_project')
        project=next(p for p in app.VBE.VBProjects if Path(str(p.FileName)).resolve()==host.resolve())
        app.VBE.ActiveVBProject=project
        app.DoCmd.RunCommand(win32com.client.constants.acCmdCompileAndSaveAllModules)
        data['native_reopen_compiled']=bool(app.IsCompiled)
        project=None
        data['independent_reopen_compile']=api('CompileVBA',True)
        data['reopen_is_compiled']=bool(app.IsCompiled)
        data['end_session_after_reopen']=api('EndSession')
        close_owned('reopen_process')
    except Exception as exc:
        data['error']=repr(exc)
    finally:
        # Quit whatever instance is still held, even if initialization stopped early.
        if current_phase is not None:
            try:shutdown_owned(current_phase)
            except Exception as exc:data['cleanup_error']=repr(exc)
        project=module=app=None;gc.collect()
        if handle is not None:
            try:handle.Close()
            except Exception as exc:data['handle_close_error']=repr(exc)
        pythoncom.CoUninitialize()
        pids=list_access_pids_or_none()
        data['after_pids']=sorted(pids) if pids is not None else None
        save()
    # Final process and cleanup checks are recorded, never asserted, so they cannot skip
    # the archive attempt below. They fail the command after the evidence is retained.
    failures=[]
    for phase in opened_phases:
        info=data[phase]
        bad=[k for k in ('init_error','identity_mismatch','quit_error','exit_timeout','exit_wait_error','handle_close_error')
             if info.get(k)]
        if bad:failures.append(phase+': '+', '.join(bad))
        if not info.get('closure_confirmed'):failures.append(phase+': closure not confirmed')
    for key in ('cleanup_error','handle_close_error'):
        if data.get(key):failures.append(key)
    if data.get('after_pids')!=[]:failures.append('Access processes remain or cannot be listed')
    data['closure_confirmed']=bool(opened_phases) and not failures
    # Retain all failure binaries and host files in verified archives before disposal.
    # The archive takes its final name only after verification, so a partial file is never
    # mistaken for evidence. Active files may block it; that is recorded and the host kept.
    archive_path=OUT/(label+'-host.zip')
    partial_path=OUT/(label+'-host.zip.partial')
    archive=data['archive']={'path':str(archive_path),'verified':False}
    try:
        assert not archive_path.exists(),'Archive already exists'
        manifest={}
        with zipfile.ZipFile(partial_path,'x',zipfile.ZIP_DEFLATED) as z:
            for p in container.rglob('*'):
                if p.is_file():
                    name=p.relative_to(container).as_posix();raw=p.read_bytes()
                    z.writestr(name,raw);manifest[name]=hashlib.sha256(raw).hexdigest()
        with zipfile.ZipFile(partial_path) as z:
            assert sorted(z.namelist())==sorted(manifest),'Archive members differ from the manifest'
            assert all(hashlib.sha256(z.read(k)).hexdigest()==v for k,v in manifest.items()),'Archive content differs'
        os.replace(partial_path,archive_path)
        archive.update(sha256=sha(archive_path),files=len(manifest),verified=True)
    except Exception as exc:
        archive['error']=repr(exc)
        failures.append('host archive not verified')
    try:data['git_status_after']=git(folder,'status','--porcelain')
    except Exception as exc:data['git_status_error']=repr(exc)
    # Preserve failed reopen hosts raw as well, for later diagnosis. A host is disposed of
    # only when closure is confirmed and its archive is verified.
    retain=list(failures)
    if not data.get('reopen_is_compiled') or data.get('error'):retain.append('lifecycle did not complete cleanly')
    if not retain:
        try:
            target=folder.resolve();owned_container=container.resolve()
            assert target.parent==owned_container and owned_container.parent==(OUT/'hosts').resolve()
            subprocess.run(['git','-C',str(REPO),'worktree','remove','--force',str(target)],check=True)
            shutil.rmtree(owned_container)
            data['disposed']=not owned_container.exists()
        except Exception as exc:
            data['dispose_error']=repr(exc);failures.append('host disposal failed');retain.append('disposal failed')
    if retain:
        data['retained_raw']=True;data['retained_reasons']=retain
    data['infrastructure_failures']=failures
    save()
    print(json.dumps({k:v for k,v in data.items() if k not in ['tests','source_manifest','initial_project','after_project','reopened_project']},indent=2),flush=True)
    assert not failures,'; '.join(failures)
    return data

if __name__=='__main__':
    p=argparse.ArgumentParser()
    p.add_argument('label');p.add_argument('binary',type=Path);p.add_argument('pin')
    p.add_argument('--evidence',type=Path,required=True)
    p.add_argument('--filters',default='');p.add_argument('--fault',choices=['assertion','import','class-error','legacy-invoker'])
    p.add_argument('--sync',action='store_true')
    a=p.parse_args();OUT=a.evidence.resolve();OUT.mkdir(parents=True,exist_ok=True)
    run(a.label,a.binary,a.pin,a.filters,a.fault,a.sync)

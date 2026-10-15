"""Run MCP tests in one owned disposable host, then validate fresh reopen.

Requires the Windows MCP Python environment and a matching source-built binary.
Never replaces modules in a loaded library or deletes a live host.
"""
import argparse
import asyncio
import gc
import hashlib
import json
import shutil
import subprocess
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
    def save():
        (OUT/(label+'.json')).write_text(json.dumps(data,indent=2,default=str),encoding='utf-8')
    def open_owned(phase):
        nonlocal app,handle
        assert list_access_pids_or_none()==set()
        app=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Access.Application'))
        pid=pid_from_access_app(app); stamp=process_create_time(pid)
        assert pid and stamp
        handle=win32api.OpenProcess(0x1000|0x100000,False,pid)
        data[phase]={'pid':pid,'create_time':stamp}
        app.OpenCurrentDatabase(str(host));app.Visible=True
        assert Path(app.CurrentProject.FullName).resolve()==host.resolve()
    def close_owned(phase):
        nonlocal app,handle,project,module
        assert process_create_time(data[phase]['pid'])==data[phase]['create_time']
        assert pid_from_access_app(app)==data[phase]['pid']
        assert Path(app.CurrentProject.FullName).resolve()==host.resolve()
        module=project=None
        app.Quit(2);app=None;gc.collect()
        data[phase]['original_handle_exit']=win32event.WaitForSingleObject(handle,5000)==win32event.WAIT_OBJECT_0
        handle.Close();handle=None
        assert data[phase]['original_handle_exit']
        assert list_access_pids_or_none()==set()
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
                text='Public Function TestClassInvoke(ByVal obj As Object, ByVal strProcName As String) As String\r\nOn Error GoTo Failed\r\nCallByName obj, "NoSuchTestMethod", VbMethod\r\nExit Function\r\nFailed:\r\nTestClassInvoke = CStr(Err.Number) & vbTab & Err.Description\r\nEnd Function\r\n'
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
            app.Run(str(INSTALLED.with_suffix(''))+'.SetInteractionMode',1)
            api('SetOption','DefaultTestFilter',filters)
            data['sync_raw_result']=api('RunFilteredTests')
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
        project=module=None
        if app is not None and handle is not None:
            try:
                phase='reopen_process' if 'reopen_process' in data else 'initial_process'
                close_owned(phase)
            except Exception as exc:data['close_error']=repr(exc)
        app=None;gc.collect()
        if handle:handle.Close()
        pythoncom.CoUninitialize()
        data['after_pids']=sorted(list_access_pids_or_none() or [])
        save()
    assert list_access_pids_or_none()==set() and not data.get('close_error'),data
    # Retain all failure binaries and host files in verified archives before disposal.
    archive=OUT/(label+'-host.zip')
    manifest={}
    with zipfile.ZipFile(archive,'x',zipfile.ZIP_DEFLATED) as z:
        for p in container.rglob('*'):
            if p.is_file():
                name=p.relative_to(container).as_posix();raw=p.read_bytes()
                z.writestr(name,raw);manifest[name]=hashlib.sha256(raw).hexdigest()
    with zipfile.ZipFile(archive) as z:
        assert all(hashlib.sha256(z.read(k)).hexdigest()==v for k,v in manifest.items())
    data['archive']={'path':str(archive),'sha256':sha(archive),'files':len(manifest),'verified':True}
    data['git_status_after']=git(folder,'status','--porcelain')
    # Preserve failed reopen hosts raw as well, for later diagnosis.
    if data.get('reopen_is_compiled') and not data.get('error'):
        target=folder.resolve();owned_container=container.resolve()
        assert target.parent==owned_container and owned_container.parent==(OUT/'hosts').resolve()
        subprocess.run(['git','-C',str(REPO),'worktree','remove','--force',str(target)],check=True)
        shutil.rmtree(owned_container)
        data['disposed']=not owned_container.exists()
    else:data['retained_raw']=True
    save()
    print(json.dumps({k:v for k,v in data.items() if k not in ['tests','source_manifest','initial_project','after_project','reopened_project']},indent=2),flush=True)
    return data

if __name__=='__main__':
    p=argparse.ArgumentParser()
    p.add_argument('label');p.add_argument('binary',type=Path);p.add_argument('pin')
    p.add_argument('--evidence',type=Path,required=True)
    p.add_argument('--filters',default='');p.add_argument('--fault',choices=['assertion','import','class-error','legacy-invoker'])
    p.add_argument('--sync',action='store_true')
    a=p.parse_args();OUT=a.evidence.resolve();OUT.mkdir(parents=True,exist_ok=True)
    run(a.label,a.binary,a.pin,a.filters,a.fault,a.sync)

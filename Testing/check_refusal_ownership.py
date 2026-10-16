"""Ensure refused requests preserve the active operation and its callback.

Use a source-built request-policy add-in and its paired pre-session MCP client.
"""
import argparse
import gc
import importlib.util
import json
import os
from pathlib import Path
import sys
import time
import uuid
import pythoncom
import win32api
import win32com.client
import win32event

def run(out,label,expect_red=False):
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none,pid_from_access_app
    from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
    assert list_access_pids_or_none()==set()
    folder=out/(label+'-host');folder.mkdir(exist_ok=False)
    spec=importlib.util.spec_from_file_location('completion',Path(__file__).with_name('check_noninteractive_policies.py'))
    common=importlib.util.module_from_spec(spec);spec.loader.exec_module(common)
    app=handle=endpoint=None;pid=stamp=None;record={}
    pythoncom.CoInitialize()
    try:
        endpoint=common.Endpoint()
        app=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Access.Application'))
        pid=pid_from_access_app(app);stamp=process_create_time(pid);assert pid and stamp
        handle=win32api.OpenProcess(0x1000|0x100000,False,pid)
        record.update(pid=pid,create_time=stamp)
        host=folder/'fixture.accdb';app.NewCurrentDatabase(str(host));app.Visible=True
        identity=uuid.uuid4().hex
        driver=str(Path(os.environ['APPDATA'])/'MSAccessVCS/Version Control')
        raw=app.Run(driver+'.RefusalOwnershipProbe',endpoint.info(identity))
        record['probe']=json.loads(raw[0] if isinstance(raw,tuple) else raw)
        until=time.monotonic()+.5
        while time.monotonic()<until:pythoncom.PumpWaitingMessages();time.sleep(.02)
        record['terminals']=endpoint.terminals()
        p=record['probe']
        refusals={'invalid_policy':'invalid_decision_policy','test_refusal':'operation_already_running',
                  'merge_refusal':'operation_already_running','ribbon_refusal':'operation_already_running'}
        record['checks']={'no_probe_error':'probe_error' not in p,'active_state_preserved':p.get('before')==p.get('after'),
            'expected_refusals':all(p.get(key,{}).get('success') is False and p[key].get('error_pattern')==pattern
                                    for key,pattern in refusals.items()),
            'no_refusal_terminal':p.get('attempts_during_refusals')==0,'one_owner_attempt':p.get('attempts_after_completion')==1,
            'one_owner_terminal':len(record['terminals'])==1 and record['terminals'][0].get('type')=='complete' and record['terminals'][0].get('operation_id')==identity,
            'root_released':p.get('completed',{}).get('status')=='ready' and not p.get('completed',{}).get('token')}
        record['verified']=all(record['checks'].values())
        if expect_red:
            assert record['checks']['no_probe_error'] and p.get('before',{}).get('token'),record['checks']
            assert p.get('attempts_during_refusals',0)>0 and not record['checks']['active_state_preserved'],record['checks']
        else:
            assert record['verified'],record['checks']
    finally:
        if app is not None:
            try:
                assert pid_from_access_app(app)==pid and process_create_time(pid)==stamp
                app.Quit(2)
            except Exception as exc:record['quit_error']=repr(exc)
        app=None;gc.collect()
        record['original_handle_exit']=handle is not None and win32event.WaitForSingleObject(handle,5000)==win32event.WAIT_OBJECT_0
        if handle is not None:handle.Close()
        if endpoint is not None:endpoint.close()
        pythoncom.CoUninitialize()
        (out/(label+'.json')).write_text(json.dumps(record,indent=2),encoding='utf-8')
        if pid is not None:assert record['original_handle_exit']
    print(json.dumps(record['checks']))

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--evidence',type=Path,required=True);p.add_argument('--label',required=True)
    p.add_argument('--client-source',type=Path,required=True);p.add_argument('--expect-red',action='store_true')
    a=p.parse_args();assert Path(a.label).name==a.label
    assert not a.evidence.resolve().is_relative_to(Path(__file__).resolve().parents[1])
    a.evidence.mkdir(parents=True,exist_ok=True);sys.path.insert(0,str(a.client_source.resolve()))
    run(a.evidence.resolve(),a.label,a.expect_red)

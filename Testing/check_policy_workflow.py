"""Exercise public policy workflows in an owned disposable Access database.

Keep evidence outside the checkout. Requires the Windows MCP Python environment.
The --legacy-control mode safely records the base's silent conflict behavior.
"""
import argparse
import gc
import hashlib
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


def run(out, label, legacy=False):
    from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none, pid_from_access_app
    assert list_access_pids_or_none()==set(), 'Serialize native work'
    folder=out/(label+'-user-host');folder.mkdir(exist_ok=False)
    spec=importlib.util.spec_from_file_location('policy_completion',Path(__file__).with_name('check_noninteractive_policies.py'))
    common=importlib.util.module_from_spec(spec);spec.loader.exec_module(common)
    endpoint=common.Endpoint()
    driver=str(Path(os.environ['APPDATA'])/'MSAccessVCS/Version Control')
    record={'legacy_control':legacy,'cases':{},'installed_sha256':hashlib.sha256(Path(driver+'.accda').read_bytes()).hexdigest()}
    def save(): (out/(label+'.json')).write_text(json.dumps(record,indent=2,default=str),encoding='utf-8')
    def unpack(value):
        value=value[0] if isinstance(value,tuple) else value
        return json.loads(value) if isinstance(value,str) and value.startswith('{') else value
    def api(method,*args):return unpack(app.Run(driver+'.API',method,*args))
    def query_value():
        database=app.CurrentDb()
        rows=database.OpenRecordset('qryPolicy')
        value=rows.Fields(0).Value
        rows.Close();rows=None;database=None
        return value
    def snapshot():return unpack(app.Run(driver+'.ModeOnlyScopeProbe','snapshot'))
    def invoke(method,*args):
        start=len(endpoint.callbacks);identity=uuid.uuid4().hex
        admitted=unpack(app.Run(driver+'.APIAsync',endpoint.info(identity),method,*args))
        deadline=time.monotonic()+45
        while time.monotonic()<deadline:
            pythoncom.PumpWaitingMessages();time.sleep(.05)
            terminal=[c for c in endpoint.callbacks[start:] if c['type'] in ['complete','error','cancelled']]
            if terminal:break
        until=time.monotonic()+.3
        while time.monotonic()<until:pythoncom.PumpWaitingMessages();time.sleep(.02)
        terminal=[c for c in endpoint.callbacks[start:] if c['type'] in ['complete','error','cancelled']]
        assert len(terminal)==1 and terminal[0]['operation_id']==identity,terminal
        return {'start':admitted,'terminal':terminal[0]}
    app=handle=None
    pythoncom.CoInitialize()
    try:
        app=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Access.Application'))
        pid=pid_from_access_app(app);stamp=process_create_time(pid)
        assert pid and stamp
        handle=win32api.OpenProcess(0x1000|0x100000,False,pid)
        record['process']={'pid':pid,'create_time':stamp}
        for policy in ['silent'] if legacy else ['block','prefer_source','prefer_database','skip','decline']:
            host=folder/(policy+'.accdb');app.NewCurrentDatabase(str(host));app.Visible=True
            assert Path(app.CurrentProject.FullName).resolve()==host.resolve()
            app.CurrentDb().CreateQueryDef('qryPolicy','SELECT 1 AS FixtureValue;')
            app.Run(driver+'.Preload')
            api('RegisterSession','workflow-'+uuid.uuid4().hex)
            assert api('SetOption','ExportFormatVersion',50100)['success']
            if legacy:app.Run(driver+'.SetInteractionMode',1)
            else:assert api('SetOperationPolicy','block')['success']
            item={'host':str(host),'baseline':invoke('FullExport')};record['cases'][policy]=item;save()
            assert item['baseline']['terminal']['type']=='complete',item
            source=Path(str(host)+'.src')/'queries/qryPolicy.sql'
            assert source.is_file()
            # Both sides change after the baseline. The export conflict is a newer
            # source file; the import conflict is a newer database object.
            source.write_bytes(b'\xef\xbb\xbfSELECT 99 AS FixtureValue;\r\n')
            os.utime(source,(time.time()+3,time.time()+3))
            app.CurrentDb().QueryDefs('qryPolicy').SQL='SELECT 2 AS FixtureValue;'
            original=source.read_bytes()
            if legacy:
                item['conflict']=invoke('FullExport')
                item['source_changed']=source.read_bytes()!=original
                item['source_text']=source.read_text(encoding='utf-8-sig')
                assert item['conflict']['terminal']['type']=='complete' and item['source_changed'],item
            else:
                assert api('SetOperationPolicy',policy)=={'success':True,'policy':policy}
                item['conflict']=invoke('FullExport')
                item['source_changed']=source.read_bytes()!=original
                item['source_text']=source.read_text(encoding='utf-8-sig')
                terminal=item['conflict']['terminal']
                if policy=='block':
                    assert terminal.get('decision_required') and not item['source_changed'],item
                else:
                    assert terminal['type']=='complete' and terminal.get('decisions'),item
                    # For export, prefer_source applies the conflict action and
                    # overwrites the changed source with the saved database query.
                    assert item['source_changed']==(policy=='prefer_source'),item
                item['after']=snapshot()
                assert item['after']['status']=='ready' and item['after']['mode']==2 and not item['after']['blocked'],item
                assert not item['after'].get('form_visible',False),item
                assert api('ClearOperationPolicy')['success']
                assert api('SetOperationPolicy','prefer_source')['success']
                # Final synchronous category and single-object results.
                item['category_export']=api('ExportByType','queries',True)
                assert item['category_export']['success'],item
                source.write_bytes(b'\xef\xbb\xbfSELECT 7 AS FixtureValue;\r\n')
                item['single_import']=api('ImportObject','query','qryPolicy')
                assert item['single_import']['success'],item
                item['imported_value']=query_value()
                assert item['imported_value']==7,item
                item['single_export']=api('ExportObject','query','qryPolicy')
                assert item['single_export']['success'],item
                source.write_bytes(b'\xef\xbb\xbfSELECT 8 AS FixtureValue;\r\n')
                item['category_import']=api('ImportByType','queries',True)
                assert item['category_import']['success'],item
                item['category_value']=query_value()
                assert item['category_value']==8,item
                item['next_operation']=invoke('FullExport')
                assert item['next_operation']['terminal']['type']=='complete',item
                assert api('ClearOperationPolicy')['success']
            save();app.CloseCurrentDatabase()
        if not legacy:
            host=folder/'probes.accdb';app.NewCurrentDatabase(str(host));app.Visible=True
            app.CurrentDb().CreateQueryDef('qryPolicy','SELECT 1 AS FixtureValue;')
            app.Run(driver+'.Preload')
            api('RegisterSession','workflow-'+uuid.uuid4().hex)
            assert api('SetOption','ExportFormatVersion',50100)['success']
            assert api('SetOperationPolicy','block')['success']
            record['initial_export']=invoke('FullExport')
            record['ownership']=unpack(app.Run(driver+'.ScopedSyncOwnershipCheck'))
            assert record['ownership']=='OK',record['ownership']
            for exporting in [False,True]:
                key='cancel_export' if exporting else 'cancel_import'
                record[key]=unpack(app.Run(driver+'.ScopedSyncCancelCheck',exporting))
                assert record[key]=='OK',record[key]
            assert api('ClearOperationPolicy')['success']
            for action in ['fault','refused']:
                assert api('SetInteractionMode',2)['success']
                record['mode_'+action]=unpack(app.Run(driver+'.ModeOnlyScopeProbe',action))
                assert not record['mode_'+action].get('probe_error'),record['mode_'+action]
            record['form_policy']=unpack(app.Run(driver+'.ExportFormPolicyCheck'))
            assert record['form_policy']=='OK',record['form_policy']
            assert api('SetOperationPolicy','block')['success']
            record['invalid_merge']=api('MergeBuild','invalid-policy')
            assert record['invalid_merge']['error_pattern']=='invalid_decision_policy'
            record['merge']=invoke('MergeBuild','prefer_source')
            assert record['merge']['terminal']['type']=='complete',record['merge']
            assert api('ClearOperationPolicy')['success']
            record['final_state']=snapshot()
            assert record['final_state']['status']=='ready' and record['final_state']['mode']==0
            assert not record['final_state']['decisions'] and not record['final_state']['blocked']
        record['verified']=True;save()
    except Exception as exc:
        record['error']=repr(exc);save();raise
    finally:
        if app is not None:
            try:
                assert pid_from_access_app(app)==pid and process_create_time(pid)==stamp
                app.Quit(win32com.client.constants.acQuitSaveNone)
            except Exception as exc:
                record['quit_error']=repr(exc)
        app=None;gc.collect()
        if handle:
            record['original_handle_exit']=win32event.WaitForSingleObject(handle,5000)==win32event.WAIT_OBJECT_0
            handle.Close()
        endpoint.close();pythoncom.CoUninitialize();save()
        assert record.get('original_handle_exit'), 'Owned process has not exited'
    print(json.dumps({'verified':record['verified'],'policies':list(record['cases']),'original_handle_exit':record['original_handle_exit']}))


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--evidence',type=Path,required=True)
    parser.add_argument('--label',required=True)
    parser.add_argument('--client-source',type=Path,required=True)
    parser.add_argument('--legacy-control',action='store_true')
    args=parser.parse_args()
    assert Path(args.label).name==args.label
    assert not args.evidence.resolve().is_relative_to(Path(__file__).resolve().parents[1])
    args.evidence.mkdir(parents=True,exist_ok=True)
    sys.path.insert(0,str(args.client_source.resolve()))
    run(args.evidence.resolve(),args.label,args.legacy_control)

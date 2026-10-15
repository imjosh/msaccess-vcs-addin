"""Check policy completion on a disposable user database and persistent callbacks.

Run with a policy-aware pre-session MCP client. Keep evidence outside the checkout.
See NONINTERACTIVE-POLICIES.md. This driver never edits a loaded library.
"""
import argparse
import asyncio
from functools import partial
import json
import os
from pathlib import Path
import sys
import threading
import time
from types import SimpleNamespace
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

def save(name, result):
    (OUT/(name+'.json')).write_text(json.dumps(result,indent=2,default=str),encoding='utf-8')

def parse_run(value):
    # Access's early-bound Application.Run includes its ByRef arguments in a tuple.
    return json.loads(value[0] if isinstance(value, tuple) else value)


class Endpoint:
    def __init__(self, cancelled=False):
        self.callbacks = []
        self.polls = []
        endpoint = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def reply(self, payload):
                data = json.dumps(payload).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def do_POST(self):
                endpoint.callbacks.append(json.loads(self.rfile.read(int(self.headers["Content-Length"]))))
                self.reply({"success": True})

            def do_GET(self):
                endpoint.polls.append(self.path)
                self.reply({"cancelled": cancelled})

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.url = f"http://127.0.0.1:{self.server.server_port}/callback"

    def info(self, operation_id):
        return json.dumps({"callback_url": self.url, "operation_id": operation_id})

    def terminals(self):
        return [c for c in self.callbacks if c["type"] in ("complete", "error", "cancelled")]

    def close(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()


def live(label):
    import pythoncom
    import win32com.client
    from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
    from msaccess_vcs_mcp.access_com.process_qos import pid_from_access_app, process_is_alive, terminate_pid
    import win32api
    import win32event
    folder = OUT / (label + '-user-host')
    folder.mkdir(exist_ok=False)
    pythoncom.CoInitialize()
    endpoint=app=handle=None
    pid=stamp=None
    initial_trapping=None
    fixture=folder/'fixture.accdb'
    result={'cases':{}}
    try:
        endpoint = Endpoint()
        wire_errors = []
        def unique_object(pairs):
            value = {}
            for key, item in pairs:
                if key in value:
                    raise ValueError('duplicate JSON key: ' + key)
                value[key] = item
            return value
        def strict_post(handler):
            raw = handler.rfile.read(int(handler.headers['Content-Length']))
            try:
                endpoint.callbacks.append(json.loads(raw, object_pairs_hook=unique_object))
            except ValueError as exc:
                wire_errors.append(str(exc))
            handler.reply({'success': not wire_errors})
        endpoint.server.RequestHandlerClass.do_POST = strict_post
        app = win32com.client.DispatchEx('Access.Application')
        pid = pid_from_access_app(app)
        stamp = process_create_time(pid)
        assert pid and stamp
        handle=win32api.OpenProcess(0x1000|0x100000,False,pid)
        scratch = SimpleNamespace(name=str(folder), cleanup=lambda: None)
        result.update(fixture_pid=pid, fixture_stamp=stamp, scratch=scratch.name)
        driver = str(Path(os.environ['APPDATA']) / 'MSAccessVCS/Version Control')

        def save_live():
            result['wire_parse_errors'] = wire_errors
            (OUT / (label + '-live.json')).write_text(json.dumps(result, indent=2), encoding='utf-8')

        def run(method, *args):
            print('API', method, flush=True)
            value = app.Run(driver + '.API', method, *args)
            value = value[0] if isinstance(value, tuple) else value
            return json.loads(value) if isinstance(value, str) and value.startswith('{') else value

        def snapshot(after=False):
            return parse_run(app.Run(driver + '.BlockedCompileSnapshot', after))

        def drain(seconds):
            until = time.monotonic() + seconds
            while time.monotonic() < until:
                pythoncom.PumpWaitingMessages()
                time.sleep(.05)

        async def tool_call(database):
            from msaccess_vcs_mcp import tools
            from msaccess_vcs_mcp.callback_server import CallbackServer
            from msaccess_vcs_mcp.operation_manager import OperationManager
            manager = OperationManager.get_instance()
            manager.set_event_loop(asyncio.get_running_loop())
            captured = []
            def route(operation_id, callback):
                captured.append(callback)
                manager.route_callback(operation_id, callback)
            server = CallbackServer(callback_router=partial(route), cancel_checker=manager.is_cancelled,
                cancel_requester=manager.request_cancel, host='127.0.0.1', port=0)
            server.start()
            os.environ['ACCESS_VCS_CALLBACK_URL'] = server.callback_url
            try:
                outcome = await tools.vcs_run_tests(str(database), filter='modTestPolicyFixture', timeout_seconds=30)
                assert manager.pending_count() == 0
                return outcome, [c for c in captured if c['type'] in ('complete', 'error', 'cancelled')]
            finally:
                server.stop()
                os.environ.pop('ACCESS_VCS_CALLBACK_URL', None)

        def case(name, transport, blocked, compile_failure, save_failure=False, completion_fault=0, delivery_failure=False, mode=0, results_override=''):
            bad_results = Path(results_override) if results_override else Path(scratch.name) / 'results-is-a-directory'
            if save_failure and not results_override:
                bad_results.mkdir(exist_ok=True)
            app.Run('ConfigureFixture', blocked, compile_failure,
                str(bad_results) if save_failure else '', completion_fault)
            start_index = len(endpoint.callbacks)
            before = snapshot()
            operation_id = 'policy-' + name + '-' + uuid.uuid4().hex
            tool_terminals = []
            if transport == 'mcp':
                outcome, tool_terminals = asyncio.run(tool_call(fixture))
                if tool_terminals:
                    operation_id = tool_terminals[0]['operation_id']
            elif transport == 'async':
                info = json.dumps({'callback_url': 'http://invalid..host/callback', 'operation_id': operation_id}) if delivery_failure else endpoint.info(operation_id)
                outcome = parse_run(app.Run(driver + '.APIAsync', info, 'RunFilteredTests', 'block'))
                until = time.monotonic() + 25
                while time.monotonic() < until:
                    pythoncom.PumpWaitingMessages()
                    if any(c['type'] in ('complete', 'error', 'cancelled') for c in endpoint.callbacks[start_index:]):
                        break
                    if delivery_failure:
                        state = snapshot()
                        if state['status'] == 'ready' and not state['token'] and state.get('completion', {}).get('run_error_pattern') == 'project_not_compiled':
                            break
                    time.sleep(.05)
            else:
                outcome = run('RunFilteredTests', 'block')
            drain(.6)
            terminals = tool_terminals or [c for c in endpoint.callbacks[start_index:] if c['type'] in ('complete', 'error', 'cancelled')]
            after = snapshot(True)
            item = {'transport': transport, 'operation_id': operation_id, 'before': before, 'outcome': outcome,
                'terminals': terminals, 'after': after}
            result['cases'][name] = item
            payload = outcome if transport == 'mcp' else (after['completion'] if delivery_failure else (terminals[0] if terminals else outcome))
            path = payload.get('results_path', payload.get('resultsPath')) or after.get('completion', {}).get('results_path')
            if path and Path(path).is_file():
                item['saved_results'] = json.loads(Path(path).read_text(encoding='utf-8-sig'))
            log = payload.get('log_path', payload.get('logPath'))
            if log and Path(log).is_file():
                item['log_excerpt'] = Path(log).read_text(encoding='utf-8-sig')[-6500:]
            save_live()
            failures = []
            def check(condition, message):
                if not condition:
                    failures.append(message)
            check(len(terminals) == (1 if transport in ('async', 'mcp') and not delivery_failure else 0), 'terminal delivery count')
            check(not wire_errors, 'no duplicate wire keys')
            check(after.get('callback_diagnostics', {}).get('terminalPostAttempts') == (1 if transport in ('async', 'mcp') else 0), 'exactly one terminal submission')
            if delivery_failure:
                check(after['callback_diagnostics']['failedPostAttempts'] > 0, 'failed delivery was exercised')
            if completion_fault:
                check(payload.get('completion_error_number') == completion_fault, 'construction failure was exercised')
            if transport == 'async' and terminals:
                check(payload['operation_id'] == operation_id, 'caller identity')
            if blocked:
                check(payload.get('success') is False and payload.get('decision_required') is True, 'blocked verdict')
                check(payload.get('error_pattern') == 'decision_required', 'decision primary')
                journal = payload.get('decisions', [])
                check(sum(d['kind'] == 'decision_required' for d in journal) == 1 and
                    any('confirmation' in d['message'] for d in journal), 'decision journal')
                check(payload.get('error') == 'A required decision was not covered by the decision policy.', 'primary text')
                if compile_failure:
                    check(payload.get('run_error_pattern') == 'project_not_compiled', 'secondary compile pattern')
                    check(payload.get('run_error') == 'Project has compile errors. Please resolve before running tests.', 'secondary compile text')
            elif compile_failure:
                check(payload.get('success') is False and payload.get('error_pattern') == 'project_not_compiled', 'compile verdict')
            else:
                check(payload.get('type', 'complete') == 'complete', 'next operation completes')
                check(not payload.get('decisions') and not payload.get('decision_required') and not payload.get('run_error_pattern'), 'no stale diagnostics')
            if compile_failure:
                check(payload.get('cancelled') is False, 'compile is not cancellation')
                partial = item.get('saved_results', payload)
                check(partial.get('allPassed') is False and partial.get('summary', {}).get('subs') == 0, 'zero partial summary')
            if save_failure:
                check(bool(payload.get('results_error')) and not path, 'persistence failure retained')
            else:
                check(bool(path) and 'saved_results' in item, 'own results file retained')
            check(bool(log) and 'log_excerpt' in item and Path(log).name.startswith('TestRun_'), 'own operation log retained')
            check(after['status'] == 'ready' and not after['token'] and after['mode'] == mode, 'root ready')
            check(after['policy'] == before['policy'], 'caller policy preserved')
            check(not after['blocked'] and not after['decisions'] and after['error_level'] == 0, 'owned scope released')
            check(not after['force_unattended'] and after['error_trapping'] == before['error_trapping'], 'trapping restored')
            check(not after.get('form_visible', False), 'form remains hidden')
            check(all(after.get(k) for k in ('log_released', 'runner_released', 'mcp_released', 'options_released')), 'infrastructure released')
            check(after['log_active'] == before['log_active'] and after['debug_suppressed'] == before['debug_suppressed'], 'logging restored')
            item['failures'] = failures
            save_live()
            print(name, 'PASS' if not failures else failures, flush=True)
            return failures
        fixture = Path(scratch.name) / 'fixture.accdb'
        print('fixture', str(fixture), 'PID', pid, flush=True)
        app.NewCurrentDatabase(str(fixture))
        app.Visible = True
        app.UserControl = True
        # Use the documented headless workflow, then restore the user's setting.
        initial_trapping = app.GetOption('Error Trapping')
        result['initial_error_trapping'] = initial_trapping
        app.SetOption('Error Trapping', 2)
        app.Run(driver + '.Preload')
        print('preloaded', flush=True)
        run('RegisterSession', 'policy-' + uuid.uuid4().hex)
        run('SetOperationPolicy', 'block')
        run('InstallTestAssertModule')
        run('ClearOperationPolicy')
        # Select by observed filename: never mutate the installed driver's project.
        project = next(p for p in app.VBE.VBProjects if Path(p.FileName) == fixture)
        helper = project.VBComponents('modTestAssert').CodeModule
        line = helper.ProcStartLine('GlobalTestSetup', 0)
        count = helper.ProcCountLines('GlobalTestSetup', 0)
        helper.DeleteLines(line, count)
        helper.AddFromString('Public Sub GlobalTestSetup()\r\n'
            f'Application.Run "{driver}.ConfigureBlockedCompileFixture", FixtureBlock, FixtureCompile, FixtureResults, FixtureFault\r\nEnd Sub\r\n')
        helper = None
        module = project.VBComponents.Add(1)
        module.Name = 'modTestPolicyFixture'
        module.CodeModule.AddFromString('Option Explicit\r\n\'@Folder("Tests")\r\n'
            'Public FixtureBlock As Boolean\r\nPublic FixtureCompile As Boolean\r\nPublic FixtureResults As String\r\nPublic FixtureFault As Long\r\n'
            'Public Sub ConfigureFixture(ByVal blocked As Boolean, ByVal failed As Boolean, ByVal results As String, ByVal fault As Long)\r\n'
            'FixtureBlock = blocked\r\nFixtureCompile = failed\r\nFixtureResults = results\r\nFixtureFault = fault\r\nEnd Sub\r\n'
            'Public Sub TestFixturePass()\r\nTestAssert True, "Policy fixture"\r\nEnd Sub\r\n')
        module = None
        project = None
        run('SetOption', 'DefaultTestFilter', 'modTestPolicyFixture')
        failures = case('combined_async', 'async', True, True)
        if failures:
            # Keep the polluted state as evidence; do not hide it with local resets.
            raise AssertionError(failures)
        for name, transport, blocked, failed, save_failed in (
            ('after_combined', 'async', False, False, False),
            ('combined_sync', 'sync', True, True, False),
            ('after_sync', 'sync', False, False, False),
            ('compile_only_async', 'async', False, True, False),
            ('compile_only_sync', 'sync', False, True, False),
            ('combined_save_failure', 'async', True, True, True),
            ('combined_sync_save_failure', 'sync', True, True, True),
            ('final_clean', 'async', False, False, False)):
            assert not case(name, transport, blocked, failed, save_failed), name
        assert not case('payload_failure', 'async', True, True, completion_fault=5)
        assert not case('after_payload_failure', 'async', False, False)
        assert not case('payload_failure_sync', 'sync', True, True, completion_fault=5)
        assert not case('after_payload_failure_sync', 'sync', False, False)
        assert not case('delivery_failure', 'async', True, True, delivery_failure=True)
        assert not case('after_delivery_failure', 'async', False, False)
        assert run('SetOperationPolicy', 'block')['success']
        assert not case('caller_session_combined', 'async', True, True, mode=2)
        assert not case('caller_session_subsequent', 'async', False, False, mode=2)
        assert run('ClearOperationPolicy')['success']
        assert not case('after_session_clear', 'async', False, False)
        for name, blocked, failed, save_failed, fault in (
            ('mcp_combined', True, True, False, 0),
            ('mcp_compile_only', False, True, False, 0),
            ('mcp_save_failure', True, True, True, 0),
            ('mcp_payload_failure', True, True, False, 5),
            ('mcp_final_clean', False, False, False, 0)):
            assert not case(name, 'mcp', blocked, failed, save_failed, fault), name
        import win32file
        existing=folder/'locked-results.json'
        existing.write_text('{"stale":true}',encoding='utf-8')
        original=existing.read_bytes()
        lock=win32file.CreateFile(str(existing),win32file.GENERIC_READ,win32file.FILE_SHARE_READ,
            None,win32file.OPEN_EXISTING,0,None)
        try:
            failures=case('locked_existing_results','sync',False,False,save_failure=True,results_override=str(existing))
        finally:
            lock.Close()
        result['locked_existing_bytes_preserved']=existing.read_bytes()==original
        save_live()
        assert not failures and result['locked_existing_bytes_preserved']
        result['verified'] = True
        save_live()
    finally:
        if app is not None:
            try:
                if initial_trapping is not None:
                    app.SetOption('Error Trapping',initial_trapping)
                    result['restored_error_trapping']=app.GetOption('Error Trapping')
                assert pid_from_access_app(app)==pid and process_create_time(pid)==stamp
                app.Quit(2)
            except Exception as exc:
                result['quit_error']=repr(exc)
        app=None
        import gc
        gc.collect()
        exited=handle is not None and win32event.WaitForSingleObject(handle,5000)==win32event.WAIT_OBJECT_0
        if handle is not None:handle.Close()
        if endpoint is not None:endpoint.close()
        pythoncom.CoUninitialize()
        save(label+'-cleanup',{'pid':pid,'stamp':stamp,'original_handle_exit':exited,'quit_error':result.get('quit_error')})
        if pid is not None:assert exited,'Owned host did not close; artifacts retained'


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--evidence',type=Path,required=True)
    parser.add_argument('--label',required=True)
    parser.add_argument('--client-source',type=Path,required=True)
    args=parser.parse_args()
    OUT=args.evidence.resolve()
    assert not OUT.is_relative_to(Path(__file__).resolve().parents[1])
    assert Path(args.label).name==args.label
    OUT.mkdir(parents=True,exist_ok=True)
    sys.path.insert(0,str(args.client_source.resolve()))
    os.environ['ACCESS_VCS_SKIP_SESSION_CLEANUP']='true'
    os.environ['ACCESS_VCS_LOG_DIR']=str(OUT/'usage')
    os.environ['ACCESS_VCS_DIAGNOSTIC_LOG_DIR']=str(OUT/'diagnostics')
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none
    assert list_access_pids_or_none()==set(), 'Serialize native Access work'
    live(args.label)

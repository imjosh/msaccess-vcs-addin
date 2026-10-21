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
        # The cancel-status answer. A case may switch it while the server runs.
        self.cancelled = cancelled
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
                self.reply({"cancelled": endpoint.cancelled})

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
    result={'cases':{}, 'seam_mode': SEAM_MODE}
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
                check((save_failure or not payload.get('decisions')) and not payload.get('decision_required') and not payload.get('run_error_pattern'), 'no stale diagnostics')
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

        def cancel_poll_case():
            # The real poll against this endpoint, through a private MCP in the driver.
            # The answer switches between probe calls, so a cached GET would show.
            name = 'cancel_poll'
            operation_id = 'policy-' + name + '-' + uuid.uuid4().hex
            info = endpoint.info(operation_id)
            path = '/cancel-status/' + operation_id
            item = {'operation_id': operation_id, 'steps': {}}
            result['cases'][name] = item
            def step(label):
                value = parse_run(app.Run(driver + '.CancelPollProbe', info, label))
                value['gets'] = sum(p == path for p in endpoint.polls)
                item['steps'][label] = value
                return value
            endpoint.cancelled = False
            try:
                first = step('first')
                endpoint.cancelled = True
                uncached = step('uncached')
                midnight = step('midnight')
            finally:
                endpoint.cancelled = False
            save_live()
            failures = []
            def check(condition, message):
                if not condition:
                    failures.append(message)
            steps = (first, uncached, midnight)
            check(not any('probe_error' in s for s in steps), 'probe ran')
            check(first.get('active') is True, 'private callback registered')
            checks = [c for s in steps for c in s['checks']]
            check(len(checks) == 4, 'four checks reported')
            if len(checks) == 4:
                poll, throttled, fresh, wrapped = checks
                check(poll['result'] is False and poll['cancelPolls'] == 1, 'first poll answers false')
                check(throttled['result'] is False and throttled['cancelPolls'] == 1, 'immediate check is throttled')
                check(first['gets'] == 1, 'one GET for the first two checks')
                check(fresh['result'] is True and fresh['cancelPolls'] == 2 and uncached['gets'] == 2,
                    'switched answer reaches the add-in uncached')
                check(wrapped['result'] is True and wrapped['cancelPolls'] == 3 and midnight['gets'] == 3,
                    'check after the midnight wrap polls')
                check(all(c['cancelPollErrors'] == 0 for c in checks), 'no failed polls')
            item['failures'] = failures
            save_live()
            print(name, 'PASS' if not failures else failures, flush=True)
            return failures
        def seam_case(name, method, args, cancel_on, expect='cancelled'):
            # The test seam answers the Nth cancel check of the run, through the
            # path a real MCP request takes. Zero runs the method unarmed.
            start_index = len(endpoint.callbacks)
            before = snapshot()
            operation_id = 'seam-' + name + '-' + uuid.uuid4().hex
            app.Run(driver + '.SetCancelPollForTest', cancel_on)
            try:
                outcome = parse_run(app.Run(driver + '.APIAsync', endpoint.info(operation_id), method, *args))
                until = time.monotonic() + 60
                while time.monotonic() < until:
                    pythoncom.PumpWaitingMessages()
                    if any(c['type'] in ('complete', 'error', 'cancelled') for c in endpoint.callbacks[start_index:]):
                        break
                    time.sleep(.05)
                drain(.6)
            finally:
                app.Run(driver + '.SetCancelPollForTest', 0)
            terminals = [c for c in endpoint.callbacks[start_index:] if c['type'] in ('complete', 'error', 'cancelled')]
            after = snapshot(True)
            item = {'method': method, 'args': list(args), 'cancel_on': cancel_on, 'operation_id': operation_id,
                'before': before, 'outcome': outcome, 'terminals': terminals, 'after': after}
            result['cases'][name] = item
            payload = terminals[0] if terminals else {}
            log = payload.get('log_path', payload.get('logPath'))
            if log and Path(log).is_file():
                item['log_excerpt'] = Path(log).read_text(encoding='utf-8-sig')[-6500:]
            path = payload.get('results_path', payload.get('resultsPath'))
            if path and Path(path).is_file():
                item['saved_results'] = json.loads(Path(path).read_text(encoding='utf-8-sig'))
            failures = []
            def check(condition, message):
                if not condition:
                    failures.append(message)
            check(outcome.get('success') is True and outcome.get('async') is True, 'async admission')
            check(len(terminals) == 1, 'terminal delivery count')
            check(payload.get('type') == expect, 'terminal type ' + expect)
            check(payload.get('operation_id') == operation_id, 'caller identity')
            check(bool(log) and 'log_excerpt' in item, 'own operation log retained')
            # Export, build and merge stop on a critical log entry; a test run stops
            # through TestRunner.Cancel instead.
            if method != 'RunFilteredTests':
                canceled = 'Canceled Operation' in item.get('log_excerpt', '')
                check(canceled == (expect == 'cancelled'), 'cancel logged only when cancelled')
            if expect == 'cancelled':
                check(payload.get('success') is False, 'cancel is not success')
            check(after['status'] == 'ready' and not after['token'], 'root ready')
            check(after['policy'] == before['policy'], 'caller policy preserved')
            check(not after['blocked'] and not after['decisions'], 'no decision recorded')
            check(all(after.get(k) for k in ('log_released', 'runner_released', 'mcp_released', 'options_released')), 'infrastructure released')
            item['failures'] = failures
            save_live()
            print(name, 'PASS' if not failures else failures, flush=True)
            return item, failures

        def stage(item, verb):
            # The last category header before the cancel names the loop that stopped.
            text = item.get('log_excerpt', '')
            text = text[:text.find('Canceled Operation')] if 'Canceled Operation' in text else text
            headers = [line.strip() for line in text.splitlines() if line.strip().startswith(verb)]
            item['stage'] = headers[-1] if headers else ''
            return item['stage']

        def source_files(source):
            return sorted(str(p.relative_to(source)) for p in source.rglob('*') if p.is_file())

        def scoped_route_cases():
            # F5 R1/R2: a saved user form delivers a trusted VBA request from
            # its Timer event while the synchronous scoped API call is running.
            # The form is outside the selected query category, so neither sync
            # closes or imports the event source.
            failures = []
            module = project.VBComponents('modTestPolicyFixture').CodeModule
            probe_code = '''
Public ProbeAttempts As Long
Public ProbeId As String
Public ProbeResponse As String
Public ProbeErrorText As String
Public Sub UP005ResetProbe()
    ProbeAttempts = 0
    ProbeId = vbNullString
    ProbeResponse = vbNullString
    ProbeErrorText = vbNullString
End Sub
Public Function UP005CancelScopedProbe() As Boolean
    On Error GoTo ProbeError
    If ProbeAttempts <> 0 Then Exit Function
    If GetSetting("MSAccessVCS", "State", "State", "0") <> "2" Then Exit Function
    ProbeId = GetSetting("MSAccessVCS", "State", "PublicOperationId", vbNullString)
    If Len(ProbeId) = 0 Then Exit Function
    ProbeAttempts = 1
    Forms("frmUP005CancelProbe").TimerInterval = 0
    ProbeResponse = CStr(Application.Run("@LIB@.DismissAddinDialog", "frmVCSMain", "cancel", ProbeId))
    UP005CancelScopedProbe = True
    Exit Function
ProbeError:
    ProbeErrorText = CStr(Err.Number) & ": " & Err.Description
End Function
Public Function UP005ProbeAttempts() As Long
    UP005ProbeAttempts = ProbeAttempts
End Function
Public Function UP005ProbeId() As String
    UP005ProbeId = ProbeId
End Function
Public Function UP005ProbeResponse() As String
    UP005ProbeResponse = ProbeResponse
End Function
Public Function UP005ProbeError() As String
    UP005ProbeError = ProbeErrorText
End Function
'''.replace('@LIB@', driver.replace('"', '""'))
            module.AddFromString(probe_code.replace('\n', '\r\n'))
            app.DoCmd.Save(5, 'modTestPolicyFixture')
            module = None
            form = app.CreateForm()
            temporary_name = form.Name
            form.OnTimer = '=UP005CancelScopedProbe()'
            form.TimerInterval = 0
            app.DoCmd.Save(2, temporary_name)
            app.DoCmd.Close(2, temporary_name, 0)
            form = None
            app.DoCmd.Rename('frmUP005CancelProbe', 2, temporary_name)
            app.DoCmd.OpenForm('frmUP005CancelProbe')
            database = app.CurrentDb()
            for number in range(300):
                database.CreateQueryDef(f'qUP005{number:03}', f'SELECT {number} AS n;')
            database = None
            debug_option = run('SetOption', 'ShowDebug', True)
            if debug_option.get('success') is not True:
                return ['could not enable object progress in disposable fixture: ' + repr(debug_option)]
            baseline = run('FullExport', 'block')
            source = Path(str(fixture) + '.src')
            source_sql = source / 'queries/qUP005000.sql'
            result['cases']['scoped_baseline'] = {'outcome': baseline, 'source_exists': source.is_dir(),
                'query_file_count': sum(1 for _ in source.rglob('qUP005*')) if source.is_dir() else 0}
            save_live()
            if baseline.get('success') is not True:
                return ['full export did not establish scoped source']
            # FullExport closes open forms while scanning every category.
            app.DoCmd.OpenForm('frmUP005CancelProbe')

            def scalar(name):
                value = app.Run(name)
                return value[0] if isinstance(value, tuple) else value

            for route, method in (('F5-R1', 'ExportByType'), ('F5-R2', 'ImportByType')):
                if route == 'F5-R1':
                    database = app.CurrentDb()
                    database.QueryDefs('qUP005000').SQL = 'SELECT 1000 AS n;'
                    database = None
                    before_object = source_sql.read_text(encoding='utf-8-sig')
                else:
                    before_object = app.CurrentDb().QueryDefs('qUP005000').SQL
                    source_sql.write_text('SELECT 2000 AS n;\n', encoding='utf-8-sig')
                app.Run('UP005ResetProbe')
                # The first log flush is pre-loop. With visible per-object debug
                # progress, a later throttled flush pumps the timer in the loop.
                app.Forms('frmUP005CancelProbe').TimerInterval = 5000 if route == 'F5-R1' else 6000
                outcome = run(method, 'query', True, 'block')
                app.Forms('frmUP005CancelProbe').TimerInterval = 0
                receipt = {'method': method, 'outcome': outcome,
                    'attempts': scalar('UP005ProbeAttempts'), 'operation_id': scalar('UP005ProbeId'),
                    'response': scalar('UP005ProbeResponse'), 'probe_error': scalar('UP005ProbeError')}
                after_object = (source_sql.read_text(encoding='utf-8-sig') if route == 'F5-R1'
                                else app.CurrentDb().QueryDefs('qUP005000').SQL)
                receipt['object_checkpoint'] = {'before': before_object, 'after': after_object,
                                                'changed': before_object != after_object}
                if route == 'F5-R2':
                    # Import merges .bas and .sql source entries. A mid-loop
                    # cancellation can precede q000's .sql merge even after
                    # many .bas entries were imported.
                    receipt['object_checkpoint']['processed_count'] = outcome.get('transferred_count', 0)
                    receipt['object_checkpoint']['changed'] = outcome.get('transferred_count', 0) > 0
                try:
                    receipt['response_json'] = json.loads(receipt['response'])
                except (TypeError, ValueError):
                    pass
                receipt['next_run'] = run(method, 'query', False, 'block')
                receipt['failures'] = [message for condition, message in (
                    (receipt['attempts'] == 1, 'timer delivered one request'),
                    (not receipt['probe_error'], 'timer VBA call had no error'),
                    ((receipt.get('response_json') or {}).get('requested') is True, 'dismiss accepted matching ID'),
                    ((receipt.get('response_json') or {}).get('operation_id') == outcome.get('operation_id') ==
                     receipt['operation_id'], 'dismiss targeted the scoped root'),
                    (outcome.get('success') is False and outcome.get('cancelled') is True,
                     'scoped root returned cancelled rather than failed'),
                    (receipt['object_checkpoint']['changed'], 'a scoped object was processed before cancel'),
                    (receipt['next_run'].get('success') is True and not receipt['next_run'].get('cancelled'),
                     'next scoped run was admitted clean')) if not condition]
                result['cases'][route] = receipt
                failures.extend(route + ': ' + message for message in receipt['failures'])
                save_live()
            app.DoCmd.Close(2, 'frmUP005CancelProbe', 0)
            return failures

        def seam_cases():
            # F5 R3-R6: each export, build, merge and test checkpoint in this owned host.
            failures = []
            app.Run('ConfigureFixture', False, False, '', 0)
            item, found = seam_case('seam_test_loop', 'RunFilteredTests', ('block',), 2)
            saved = item.get('saved_results', {})
            item['tests_run'] = saved.get('summary', {}).get('subs')
            found += [m for c, m in (
                (saved.get('cancelled') is True, 'saved test run reports cancelled'),
                (saved.get('allPassed') is False, 'cancelled run is not allPassed'),
                (item['tests_run'] == 1, 'partial results keep the test that ran')) if not c]
            failures += found
            source = Path(scratch.name) / 'fixture.accdb.src'
            # Save the fixture's VBE edits: an unsaved module makes an export raise
            # Access's own Save prompt, which no decision policy answers.
            for module_name in ('modTestAssert', 'modTestPolicyFixture'):
                app.DoCmd.Save(5, module_name)
            item, found = seam_case('seam_export_baseline', 'FullExport', ('block',), 0, 'complete')
            failures += found
            baseline = source_files(source)
            item['source_files'] = baseline
            item, found = seam_case('seam_export_scan', 'FullExport', ('block',), 1)
            found += [] if not stage(item, 'Exporting ') else ['cancel came after the scan']
            failures += found
            # 24 scan checks for an accdb, then one per exported object.
            item, found = seam_case('seam_export_object', 'FullExport', ('block',), 26)
            found += [] if stage(item, 'Exporting ') else ['cancel came before the object loop']
            kept = source_files(source)
            item['source_files'] = kept
            found += [] if set(baseline) <= set(kept) else ['partial export kept the source']
            failures += found
            item, found = seam_case('seam_export_after', 'FullExport', ('block',), 0, 'complete')
            failures += found
            if SEAM_MODE == 'build':
                # A full build runs the regular loop for project files and the
                # fast loop for modules. Its cancelled host is disposable.
                before_modules = sum((source / n).is_file() for n in
                    ('project.json', 'vbe-project.json', 'vbe-references.json', 'proj-properties.json'))
                item, found = seam_case('seam_build_regular', 'Build', (str(source), 'block'), 1)
                found += [] if not stage(item, 'Importing ').startswith('Importing modules') and item['stage'] else ['build stopped in its regular loop']
                failures += found
                item, found = seam_case('seam_build_fast', 'Build', (str(source), 'block'), before_modules + 2)
                item['files_before_modules'] = before_modules
                found += [] if stage(item, 'Importing ').startswith('Importing modules') else ['build stopped in its fast loop']
                failures += found
                save_live()
                return failures
            # The completed full export is the merge baseline. MergeBuild reads merge
            # availability from frmVCSMain, which sets it on load, so close a hidden
            # copy an earlier operation left loaded.
            app.DoCmd.Close(2, 'frmVCSMain')
            # A merge imports only changed files, so change one module's source.
            changed = next(source.rglob('modTestPolicyFixture.bas'))
            original = changed.read_bytes()
            changed.write_bytes(original + b"' Seam merge marker\r\n")
            try:
                item, found = seam_case('seam_merge_loop', 'MergeBuild', ('block',), 1)
            finally:
                changed.write_bytes(original)
            found += [] if stage(item, 'Merging ').startswith('Merging modules') else ['merge stopped in its file loop']
            failures += found
            # ExportByType and ImportByType: a cancel and an accepted request after Begin,
            # outside the runner's root, in this library host.
            checks = result['cases']['scoped_sync'] = {}
            for check_name in ('ScopedSyncCancelCheck', 'ScopedSyncRequestCancelCheck'):
                for exporting in (False, True):
                    key = check_name + ('(export)' if exporting else '(import)')
                    # Application.Run returns the result with its ByRef arguments.
                    value = app.Run(driver + '.' + check_name, exporting)
                    checks[key] = value[0] if isinstance(value, tuple) else value
                    if checks[key] != 'OK':
                        failures.append(key + ': ' + str(checks[key]))
            save_live()
            print('scoped_sync', checks, flush=True)
            save_live()
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
        mode_result = run('SetInteractionMode', 1)
        if not mode_result.get('success'):
            raise RuntimeError(mode_result)
        run('InstallTestAssertModule')
        mode_result = run('SetInteractionMode', 0)
        if not mode_result.get('success'):
            raise RuntimeError(mode_result)
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
            'Public Sub TestFixturePass()\r\nTestAssert True, "Policy fixture"\r\nEnd Sub\r\n'
            'Public Sub TestFixtureSecond()\r\nTestAssert True, "Policy fixture second"\r\nEnd Sub\r\n')
        module = None
        project = None
        run('SetOption', 'DefaultTestFilter', 'modTestPolicyFixture')
        if SEAM_MODE == 'scoped':
            # Reacquire this host's VBE project for the saved timer fixture.
            project = next(p for p in app.VBE.VBProjects if Path(p.FileName) == fixture)
            failures = scoped_route_cases()
            if failures:
                raise AssertionError(failures)
            result['verified'] = True
            save_live()
            return
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
            failures = case(name, transport, blocked, failed, save_failed)
            if failures:
                raise AssertionError(failures)
        failures = case('payload_failure', 'async', True, True, completion_fault=5)
        if failures:
            raise AssertionError(failures)
        failures = case('after_payload_failure', 'async', False, False)
        if failures:
            raise AssertionError(failures)
        failures = case('payload_failure_sync', 'sync', True, True, completion_fault=5)
        if failures:
            raise AssertionError(failures)
        failures = case('after_payload_failure_sync', 'sync', False, False)
        if failures:
            raise AssertionError(failures)
        failures = case('delivery_failure', 'async', True, True, delivery_failure=True)
        if failures:
            raise AssertionError(failures)
        failures = case('after_delivery_failure', 'async', False, False)
        if failures:
            raise AssertionError(failures)
        retired_policy = run('SetOperationPolicy', 'block')
        retired_mode = run('SetInteractionMode', 2)
        harmless_clear = run('ClearOperationPolicy')
        result['retired_calls'] = {'policy': retired_policy, 'mode': retired_mode, 'clear': harmless_clear}
        if retired_policy.get('error_pattern') != 'policy_activation_retired':
            raise AssertionError(retired_policy)
        if retired_mode.get('success') or retired_mode.get('effective_mode') != 0:
            raise AssertionError(retired_mode)
        if not harmless_clear.get('success'):
            raise AssertionError(harmless_clear)
        failures = case('after_retired_calls', 'async', False, False)
        if failures:
            raise AssertionError(failures)
        for name, blocked, failed, save_failed, fault in (
            ('mcp_combined', True, True, False, 0),
            ('mcp_compile_only', False, True, False, 0),
            ('mcp_save_failure', True, True, True, 0),
            ('mcp_payload_failure', True, True, False, 5),
            ('mcp_final_clean', False, False, False, 0)):
            failures = case(name, 'mcp', blocked, failed, save_failed, fault)
            if failures:
                raise AssertionError(failures)
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
        failures = cancel_poll_case()
        if failures:
            raise AssertionError(failures)
        failures = seam_cases()
        if failures:
            raise AssertionError(failures)
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
    parser.add_argument('--seam-mode',choices=('merge', 'build', 'scoped'),default='merge')
    args=parser.parse_args()
    SEAM_MODE=args.seam_mode
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

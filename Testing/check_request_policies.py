"""Qualify request ownership through a disposable host and source-built add-in.

Run with optimization disabled and a compatible pre-session MCP client:
  python Testing/check_request_policies.py --evidence <fresh folder> --label <name> --client-source <MCP src>

The supervisor must serialize native runs and install the candidate first. This
driver preserves the persistent operation-ID ledger and owns only its new host.
The preflight and ownership case sets diagnose lifecycle failures separately;
only the default all case set can produce a full verified result.
"""
import argparse
import gc
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import struct
import sys
import time
import uuid
import threading
from concurrent.futures import ThreadPoolExecutor

EXIT_WAIT_MS = 60000
INVENTORY_SETTLE_MS = 30000

PROBE_BRIDGE_SOURCE = """Attribute VB_Name = "modRequestPolicyBridge"
'---------------------------------------------------------------------------------------
' Module    : modRequestPolicyBridge
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Compare installed probe invocation through a public disposable-host module.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : RequestPolicyFixtureBridge
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Preserve supplied arguments and successful results, and return request-local
'           : probe_error diagnostics when invocation fails. This standalone fixture has
'           : no installed logger reference, so it reports inherited Err locally before
'           : installing trapping rather than logging into another operation's owner.
'---------------------------------------------------------------------------------------
'
Public Function RequestPolicyFixtureBridge(ByVal strTarget As String, _
    Optional varArg1 As Variant, Optional varArg2 As Variant, Optional varArg3 As Variant) As Variant
    Dim lngBridgeError As Long
    Dim strBridgeError As String
    Dim strBridgeSource As String
    Dim varDetails As Variant
    Dim lngPart As Long
    Dim lngCharacter As Long

    If Err.Number <> 0 Then Debug.Print "RequestPolicyFixtureBridge inherited error " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo BridgeError
    If Not IsMissing(varArg3) Then
        RequestPolicyFixtureBridge = Application.Run(strTarget, varArg1, varArg2, varArg3)
    ElseIf Not IsMissing(varArg2) Then
        RequestPolicyFixtureBridge = Application.Run(strTarget, varArg1, varArg2)
    ElseIf Not IsMissing(varArg1) Then
        RequestPolicyFixtureBridge = Application.Run(strTarget, varArg1)
    Else
        RequestPolicyFixtureBridge = Application.Run(strTarget)
    End If
    Exit Function

BridgeError:
    lngBridgeError = Err.Number
    strBridgeError = Err.Description
    strBridgeSource = Err.Source
    Resume ReportBridgeError

ReportBridgeError:
    On Error GoTo ReportError
    varDetails = Array(strBridgeError, strBridgeSource, strTarget)
    For lngPart = LBound(varDetails) To UBound(varDetails)
        varDetails(lngPart) = Replace(varDetails(lngPart), Chr$(92), Chr$(92) & Chr$(92))
        varDetails(lngPart) = Replace(varDetails(lngPart), Chr$(34), Chr$(92) & Chr$(34))
        For lngCharacter = 0 To 31
            varDetails(lngPart) = Replace(varDetails(lngPart), Chr$(lngCharacter), _
                Chr$(92) & "u" & Right$("0000" & Hex$(lngCharacter), 4))
        Next lngCharacter
    Next lngPart
    RequestPolicyFixtureBridge = "{""probe_error"":" & Chr$(34) & varDetails(0) & Chr$(34) & _
        ",""probe_error_number"":" & CStr(lngBridgeError) & _
        ",""probe_error_source"":" & Chr$(34) & varDetails(1) & Chr$(34) & _
        ",""probe_error_target"":" & Chr$(34) & varDetails(2) & Chr$(34) & "}"
    Exit Function

ReportError:
    Err.Clear
    RequestPolicyFixtureBridge = "{""probe_error"":""The diagnostic probe bridge failed; error formatting was unavailable.""}"
End Function
"""


def require(condition, message):
    if not condition:
        raise AssertionError(message)


def decode(value):
    return json.loads(value[0] if isinstance(value, tuple) else value)


def index_build_dates(path):
    """Read the actual index header without changing its stored history."""
    data = path.read_bytes()
    require(len(data) >= 38 and data[:4] == b"VCSI", "The fixture index header is invalid")
    require(int.from_bytes(data[4:6], "little") in (3, 4, 5), "The fixture index version is unsupported")
    return dict(zip(("MergeBuildDate", "FullBuildDate", "ExportDate", "FullExportDate"),
                    struct.unpack_from("<dddd", data, 6)))


def run(evidence, label, case_set="all", probe_bridge=False):
    import pythoncom
    import win32api
    import win32com.client
    import win32event
    from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none, pid_from_access_app
    from msaccess_vcs_mcp.access_com.instance_registry import process_create_time

    spec = importlib.util.spec_from_file_location(
        "completion_driver", Path(__file__).with_name("check_noninteractive_policies.py"))
    common = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(common)
    folder = evidence / (label + "-host")
    folder.mkdir(exist_ok=False)
    record = {"cases": {}, "checks": {}, "case_set": case_set,
              "diagnostic_only": case_set != "all" or probe_bridge,
              "probe_transport": "fixture_bridge" if probe_bridge else "direct"}
    app = endpoint = handle = None
    pid = stamp = None
    library = str(Path(os.environ["APPDATA"]) / "MSAccessVCS" / "Version Control")
    # Access resolves this unique public function in the verified current
    # fixture. Its file path is not a library qualifier for the current project.
    bridge_target = "RequestPolicyFixtureBridge"
    bridge_ready = False

    def call_probe(method, *arguments):
        if probe_bridge:
            require(bridge_ready, "The disposable probe bridge was not installed")
            require(Path(app.CurrentProject.FullName).resolve() == (folder / "fixture.accdb").resolve(),
                    "The disposable bridge invocation host changed")
            return app.Run(bridge_target, library + "." + method, *arguments)
        return app.Run(library + "." + method, *arguments)

    def verify_bridge(stage, *, save_module=False):
        details = {"host": str(app.CurrentProject.FullName), "saved": False}
        record["probe_bridge"].setdefault("verification", {})[stage] = details
        require(Path(details["host"]).resolve() == (folder / "fixture.accdb").resolve(),
                "The disposable bridge inspection host changed")
        vbe = projects = project = component = code_module = None
        try:
            vbe = win32com.client.gencache.EnsureDispatch(app.VBE)
            projects = vbe.VBProjects
            for index in range(1, int(projects.Count) + 1):
                project = projects.Item(index)
                if Path(str(project.FileName)).resolve() == (folder / "fixture.accdb").resolve():
                    break
                project = None
            require(project is not None, "The owned fixture VBA project was not found")
            details["project_file"] = str(project.FileName)
            component = project.VBComponents.Item("modRequestPolicyBridge")
            details["component_name"] = str(component.Name)
            details["component_type"] = int(component.Type)
            require(details["component_name"] == "modRequestPolicyBridge"
                    and details["component_type"] == win32com.client.constants.vbext_ct_StdModule, details)
            if save_module:
                app.DoCmd.Save(win32com.client.constants.acModule, "modRequestPolicyBridge")
                vbe.ActiveVBProject = project
                if not bool(app.IsCompiled):
                    app.DoCmd.RunCommand(win32com.client.constants.acCmdCompileAndSaveAllModules)
                details["saved"] = True
                details["compiled"] = bool(app.IsCompiled)
                require(details["compiled"], "The owned fixture bridge did not compile")
            code_module = component.CodeModule
            start = int(code_module.ProcStartLine("RequestPolicyFixtureBridge", win32com.client.constants.vbext_pk_Proc))
            count = int(code_module.ProcCountLines("RequestPolicyFixtureBridge", win32com.client.constants.vbext_pk_Proc))
            body = str(code_module.Lines(start, count))
            details.update(procedure="RequestPolicyFixtureBridge", start_line=start, line_count=count,
                           source_sha256=hashlib.sha256(body.encode("utf-8")).hexdigest())
            require(start > 0 and count > 0 and "Public Function RequestPolicyFixtureBridge(" in body, details)
            declarations = str(code_module.Lines(1, int(code_module.CountOfDeclarationLines)))
            require("Option Private Module" not in declarations, "The fixture bridge must remain public")
            details["readable"] = True
        except Exception as exc:
            details["error"] = repr(exc)
            raise
        finally:
            code_module = component = project = projects = vbe = None

    def verify_bridge_invocation(stage):
        value = call_probe("GetVCSVersion")
        value = value[0] if isinstance(value, tuple) else value
        record["probe_bridge"]["verification"][stage]["invocation_result"] = value
        require(isinstance(value, str) and value and not value.startswith("{"),
                "The public fixture bridge did not return the installed version")

    pythoncom.CoInitialize()
    try:
        endpoint = common.Endpoint()
        app = win32com.client.DispatchEx("Access.Application")
        pid = pid_from_access_app(app)
        stamp = process_create_time(pid)
        require(pid and stamp, "Could not identify the owned Access process")
        handle = win32api.OpenProcess(0x1000 | 0x100000, False, pid)
        record.update(pid=pid, process_create_time=stamp)
        app.NewCurrentDatabase(str(folder / "fixture.accdb"))
        app.Visible = True
        app.UserControl = True
        record["user_control"] = bool(app.UserControl)
        if probe_bridge:
            require(Path(app.CurrentProject.FullName).resolve() == (folder / "fixture.accdb").resolve(),
                    "The disposable bridge host changed")
            bridge_file = folder / "modRequestPolicyBridge.bas"
            bridge_bytes = b"\xef\xbb\xbf" + PROBE_BRIDGE_SOURCE.replace("\n", "\r\n").encode("utf-8")
            bridge_file.write_bytes(bridge_bytes)
            app.LoadFromText(win32com.client.constants.acModule, "modRequestPolicyBridge", str(bridge_file))
            record["probe_bridge"] = {"module": "modRequestPolicyBridge", "entrypoint": bridge_target,
                                      "source_file": str(bridge_file), "sha256": hashlib.sha256(bridge_bytes).hexdigest()}
            verify_bridge("after_load", save_module=True)
            bridge_ready = True
        app.Run(library + ".Preload")
        if probe_bridge:
            verify_bridge_invocation("after_load")

        def save():
            (evidence / (label + ".json")).write_text(
                json.dumps(record, indent=2, default=str), encoding="utf-8")

        def probe(case):
            identity = "request-policy-" + uuid.uuid4().hex
            start = len(endpoint.callbacks)
            result = decode(call_probe("RequestPolicyProbe", case, endpoint.info(identity)))
            until = time.monotonic() + 0.2
            while time.monotonic() < until:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.01)
            record["cases"][case] = {
                "operation_id": identity, "result": result,
                "callbacks": endpoint.callbacks[start:],
            }
            save()
            require("probe_error" not in result, {case: result})
            return result

        def f4_probe(case, policy="block"):
            identity = "f4-policy-" + uuid.uuid4().hex
            start = len(endpoint.callbacks)
            result = decode(call_probe("RequestPolicyF4Probe", case, endpoint.info(identity), policy))
            until = time.monotonic() + 0.2
            while time.monotonic() < until:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.01)
            key = "f4_" + case + "_" + policy
            record["cases"][key] = {
                "operation_id": identity, "result": result,
                "callbacks": endpoint.callbacks[start:],
            }
            save()
            require("probe_error" not in result, {key: result})
            return identity, result, record["cases"][key]["callbacks"]

        def terminal_callbacks(callbacks):
            return [c for c in callbacks if c["type"] in ("complete", "error", "cancelled")]

        def host_files():
            return {
                "host_path": str(app.CurrentProject.FullName),
                "host_exists": (folder / "fixture.accdb").is_file(),
                "top_level": sorted(p.name for p in folder.iterdir() if p.suffix.lower() != ".laccdb"),
                "backups": sorted(str(p.relative_to(folder)) for p in folder.rglob("*_VCSBackup*")),
            }

        # A51: requests with policy or session-owned state need a session/1 envelope.
        identity = decode(app.Run(library + ".APIIdentity"))
        envelope = json.dumps({
            "session_id": str(uuid.uuid4()), "server_instance": str(uuid.uuid4()),
            "connection_id": str(uuid.uuid4()), "server_version": "0.3.0-dev.18",
            "protocol": "msaccess-vcs.session/1", "addin_instance": identity["addin_instance"]})
        record["compatibility_handshake"] = decode(app.Run(library + ".APIHandshake", envelope))
        require(record["compatibility_handshake"].get("success") is True, record["compatibility_handshake"])

        def api(method, *arguments):
            raw = app.Run(library + ".APIExecute", envelope, method, *arguments)
            raw = raw[0] if isinstance(raw, tuple) else raw
            return json.loads(raw) if isinstance(raw, str) and raw.startswith("{") else raw

        fixture_source = folder / "fixture.accdb.src"
        fixture_source.mkdir(exist_ok=True)
        options_path = fixture_source / "vcs-options.json"
        fixture_options = {"Info": {}, "Options": {"ExportFolder": "", "ExportFormatVersion": "5.1.0",
                           "RunBeforeBuild": "", "RunAfterBuild": "", "RunBeforeMerge": "", "RunAfterMerge": "",
                           "ValidateAfterBuild": ""}}
        options_path.write_bytes(b"\xef\xbb\xbf" + (json.dumps(fixture_options, indent=2) + "\r\n").encode("utf-8"))
        record["fixture_options_seed"] = {"path": str(options_path),
                                          "sha256": hashlib.sha256(options_path.read_bytes()).hexdigest()}
        session = api("RegisterSession", "f3-driver-" + uuid.uuid4().hex)
        require(session.get("success"), session)
        database = app.CurrentDb()
        database.CreateQueryDef("qryF3RequestFixture", "SELECT 1 AS FixtureValue;")
        database = None
        # Prepare real sources for this owned database. The headless full build
        # establishes legitimate FullBuildDate history without RunExport's form path.
        metadata_types = ["project", "vbe_project"]
        if probe_bridge:
            metadata_types.append("modules")
        record["fixture_metadata_exports"] = {}
        for category in metadata_types:
            exported = api("ExportByType", category, True, "block")
            record["fixture_metadata_exports"][category] = exported
            save()
            require(exported.get("success") is True, {"fixture_metadata_export": category, "result": exported})
        query_seed = api("ExportByType", "queries", True, "block")
        record["fixture_query_source_export"] = query_seed
        save()
        require(query_seed.get("success") is True, query_seed)
        index_path = fixture_source / "vcs-index.idx"
        dates_before_build = index_build_dates(index_path)
        record["fixture_index_before_build"] = dates_before_build
        if probe_bridge:
            bridge_ready = False
        built = api("BuildHeadless", str(fixture_source))
        record["fixture_headless_full_build"] = built
        save()
        require(built.get("success") is True, built)
        require(Path(app.CurrentProject.FullName).resolve() == (folder / "fixture.accdb").resolve(),
                "The headless setup build changed the owned fixture path")
        rows = database = None
        try:
            database = app.CurrentDb()
            rows = database.OpenRecordset("qryF3RequestFixture")
            query_value = rows.Fields("FixtureValue").Value
            require(query_value == 1, {"rebuilt_fixture_query_value": query_value})
            record["fixture_query_value_after_build"] = query_value
        finally:
            try:
                if rows is not None:
                    rows.Close()
            finally:
                rows = database = None
        dates_after_build = index_build_dates(index_path)
        record["fixture_index_after_build"] = dates_after_build
        require(dates_after_build["FullBuildDate"] > dates_before_build["FullBuildDate"],
                {"before": dates_before_build, "after": dates_after_build})
        if probe_bridge:
            verify_bridge("after_build")
            bridge_ready = True
            verify_bridge_invocation("after_build")
        baseline = api("ExportByType", "queries", True, "block")
        record["fixture_export"] = baseline
        require(baseline.get("success"), baseline)
        prepared_dates = index_build_dates(index_path)
        require(prepared_dates["FullBuildDate"] == dates_after_build["FullBuildDate"] > 0, prepared_dates)
        record["fixture_prepared_full_build_dates"] = prepared_dates
        fixture_baselines = {"options": options_path,
                             "project": fixture_source / "project.json",
                             "vbe_project": fixture_source / "vbe-project.json",
                             "query": fixture_source / "queries/qryF3RequestFixture.sql",
                             "index": index_path}
        for key, path in fixture_baselines.items():
            require(path.is_file(), {"missing_fixture_baseline": key, "path": str(path)})
        baseline_hashes = {key: hashlib.sha256(path.read_bytes()).hexdigest()
                           for key, path in fixture_baselines.items()}
        record["fixture_baselines"] = {key: {"path": str(path), "sha256": baseline_hashes[key]}
                                       for key, path in fixture_baselines.items()}
        # A newly created database remains exclusive until its first close. The
        # in-place merge acceptance cases require an actually shared fixture.
        if probe_bridge:
            bridge_ready = False
        app.CloseCurrentDatabase()
        app.OpenCurrentDatabase(str(folder / "fixture.accdb"), False)
        app.Run(library + ".Preload")
        require(Path(app.CurrentProject.FullName).resolve() == (folder / "fixture.accdb").resolve(),
                "The reopened request fixture host changed")
        for key, path in fixture_baselines.items():
            require(hashlib.sha256(path.read_bytes()).hexdigest() == baseline_hashes[key],
                    {"fixture_baseline_changed_on_reopen": key, "path": str(path)})
        record["fixture_baselines_survived_reopen"] = True
        if probe_bridge:
            verify_bridge("after_reopen")
            bridge_ready = True
            verify_bridge_invocation("after_reopen")
        retired = probe("retired")
        require(json.loads(retired["setter"])["error_pattern"] == "policy_activation_retired", retired)
        require(json.loads(retired["mode"])["success"] is False, retired)
        require(json.loads(retired["mode"])["effective_mode"] == 0, retired)
        require(json.loads(retired["clear"])["success"] is True, retired)
        if case_set in ("all", "preflight"):
            # Full-build confirmation must be reached as a UI request. The test
            # helper acquires a request-owned policy before entering the actual form.
            f4_before = host_files()
            for policy in ("block", "decline", "prefer_source", "prefer_database", "skip"):
                identity, value, callbacks = f4_probe("full_build_confirmation", policy)
                completion = value["completion"]
                terminals = terminal_callbacks(callbacks)
                require(value["source_before_form_entry"] == "user_interface", value)
                require(value["policy_before_form_entry"] == policy and value["valid_source_options"] is True, value)
                require(completion["operation_id"] == identity and completion["initiating_source"] == "user_interface", completion)
                require(completion["decision_policy"] == policy and completion["success"] is False, completion)
                require(completion.get("errorNumber") != -2147219094, completion)
                decisions = completion.get("decisions", [])
                require(len(decisions) == 1 and decisions[0]["message"].startswith("Build fixture (fixture.accdb) from source?"), completion)
                require(decisions[0].get("detail") != "file_dialog", completion)
                if policy == "decline":
                    require(completion["type"] == "cancelled" and decisions[0]["resolution"] == "declined", completion)
                else:
                    require(completion["error_pattern"] == "decision_required" and decisions[0]["resolution"] == "decision_required", completion)
                require(value["after"]["active"] is False and value["after"]["pending"] is False, value)
                require(value["after"]["mode"] == 0 and value["after"]["decisions"] == [], value)
                require(len(terminals) == 1 and terminals[0]["operation_id"] == identity, terminals)
                require(terminals[0]["decisions"] == decisions, terminals)
                require(host_files() == f4_before, {"before": f4_before, "after": host_files()})
            record["f4_full_build_files"] = {"before": f4_before, "after": host_files()}

            identity, active_form, callbacks = f4_probe("active_form", "decline")
            require(active_form["before"] == active_form["after"], active_form)
            require(json.loads(active_form["public_refusal"])["error_pattern"] == "operation_already_running", active_form)
            require(active_form["ribbon_accepted"] is False, active_form)
            require(active_form["attempts_before_completion"] == 0 and active_form["attempts_after_completion"] == 1, active_form)
            terminals = terminal_callbacks(callbacks)
            require(len(terminals) == 1 and terminals[0]["operation_id"] == identity, terminals)
            require(len(terminals[0]["decisions"]) == 1 and terminals[0]["decisions"][0]["title"] == "Active F4 owner", terminals)

            # Preserve the prepared real full-build index while making source-folder
            # discovery fail. A merge-not-available refusal or a started real merge fails.
            source = Path(api("GetExportFolder"))
            options_path = source / "vcs-options.json"
            held_options = source / "vcs-options.f4-held.json"
            index_paths = sorted(source.glob("vcs-index.*"))
            require(options_path.is_file() and not held_options.exists() and index_paths, str(source))
            index_before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in index_paths}
            prior_id, prior, prior_callbacks = f4_probe("previous_journal")
            require(len(prior["completion"]["decisions"]) == 1 and prior["completion"]["decisions"][0]["title"] == "Previous F4 operation", prior)
            require(prior["after"]["decisions"] == [] and prior["after"]["active"] is False, prior)
            require(len(terminal_callbacks(prior_callbacks)) == 1, prior_callbacks)
            missing_id = "f4-missing-folder-" + uuid.uuid4().hex
            callback_start = len(endpoint.callbacks)
            missing = {"operation_id": missing_id, "previous_operation_id": prior_id, "index_before": index_before}
            record["f4_missing_folder"] = missing
            missing_files = host_files()
            options_path.replace(held_options)
            missing["source_options_absent_before_dispatch"] = not options_path.exists()
            try:
                missing["admission"] = decode(app.Run(library + ".APIExecuteAsync", envelope, endpoint.info(missing_id), "MergeBuild", "block"))
                require(missing["admission"].get("async") is True, missing)
                deadline = time.monotonic() + 30
                terminals = []
                while time.monotonic() < deadline:
                    pythoncom.PumpWaitingMessages()
                    terminals = terminal_callbacks(endpoint.callbacks[callback_start:])
                    if terminals:
                        break
                    time.sleep(0.02)
                # Allow any incorrect duplicate terminal delivery to become observable.
                until = time.monotonic() + 0.2
                while time.monotonic() < until:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.01)
                terminals = terminal_callbacks(endpoint.callbacks[callback_start:])
                missing["callbacks"] = endpoint.callbacks[callback_start:]
                missing["terminals"] = terminals
                save()
                require(len(terminals) == 1, missing)
                terminal = terminals[0]
                require(terminal["operation_id"] == missing_id and terminal["initiating_source"] == "mcp", terminal)
                require(terminal["success"] is False and terminal["error_pattern"] == "decision_required", terminal)
                require(terminal["decision_policy"] == "block", terminal)
                decisions = terminal.get("decisions", [])
                require(len(decisions) == 1 and decisions[0]["detail"] == "file_dialog", terminal)
                require(decisions[0]["kind"] == "decision_required" and decisions[0]["resolution"] == "blocked", terminal)
                require(decisions[0]["title"] != "Previous F4 operation", terminal)
                require(not options_path.exists(), "Missing-folder preflight unexpectedly recreated source options")
                missing["index_after"] = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in index_paths}
                require(missing["index_after"] == index_before, missing)
            finally:
                if options_path.exists():
                    options_path.replace(source / "vcs-options.f4-unexpected.json")
                if held_options.exists():
                    held_options.replace(options_path)
            _, missing["after"], _ = f4_probe("inspect")
            require(missing["after"]["active"] is False and missing["after"]["pending"] is False, missing)
            require(missing["after"]["mode"] == 0 and missing["after"]["decisions"] == [], missing)
            require(missing["after"].get("form_merge_enabled") is True, missing)
            require(host_files() == missing_files, {"before": missing_files, "after": host_files()})
            record["checks"]["f4_preflight_ownership"] = True
            save()

        if case_set in ("all", "ownership"):
            for case in ("active_isolation", "cancel_target", "stale_continuation"):
                value = probe(case)
                require(value["before"] == value["after"], {case: value})
                require(value["attempts_before_completion"] == 0, {case: value})
                require(value["attempts_after_completion"] == 1, {case: value})
                callbacks = record["cases"][case]["callbacks"]
                terminals = [c for c in callbacks if c["type"] in ("complete", "error", "cancelled")]
                require(len(terminals) == 1, {case: terminals})
                require(terminals[0]["operation_id"] == record["cases"][case]["operation_id"], terminals)
                if case == "cancel_target":
                    require(json.loads(value["missing"])["error_pattern"] == "operation_id_required", value)
                    require(json.loads(value["mismatch"])["error_pattern"] == "operation_id_mismatch", value)
            for case in ("timer_refusal", "stale_timer_event"):
                value = probe(case)
                require(json.loads(value["queue"])["success"] is True, value)
                require(value["before"] == value["after"], {case: value})
            failed = probe("failed_timer")
            require(json.loads(failed["queue"])["error_pattern"] == "timer_install_failed", failed)
            require(failed["pending"] is False, failed)
            restored = probe("recreation")
            require(restored["restored"] is True, restored)
            require(restored["operation_id"] == restored["before"]["operation_id"], restored)
            require(restored["token"] == restored["before"]["token"], restored)
            require(restored["policy"] == "block" and restored["source"] == "mcp", restored)
            require(restored["attended"] is False and restored["resume_wrong"] is False, restored)
            rechecked = probe("queue_recheck")
            require(json.loads(rechecked["queue"]).get("async") is True, rechecked)
            require(rechecked["before"] == rechecked["after"], rechecked)
            terminals = [c for c in record["cases"]["queue_recheck"]["callbacks"]
                         if c["type"] in ("complete", "error", "cancelled")]
            require(len(terminals) == 1 and terminals[0]["error_pattern"] == "operation_already_running", terminals)
            require(terminals[0]["operation_id"] == record["cases"]["queue_recheck"]["operation_id"], terminals)
            dynamic = probe("dynamic_failure")
            require(json.loads(dynamic["result"])["success"] is False and dynamic["pending"] is False, dynamic)
            expired = probe("inspect_expiry")
            require(expired["active"] is False and expired["completion"]["success"] is False, expired)
            require(expired["completion"]["operation_id"] == record["cases"]["inspect_expiry"]["operation_id"], expired)
            disconnected = probe("disconnect")
            require(disconnected["owned_policy_after_disconnect"] == "block", disconnected)
            require(disconnected["completed_id"] == disconnected["expected_id"], disconnected)
            require(disconnected["next_policy"] == "" and disconnected["next_mode"] == 0, disconnected)
            require(disconnected["next_source"] == "user_interface", disconnected)
            cancelled = probe("cancel_reuse")
            require(cancelled["before"] == cancelled["after_stale"], cancelled)
            require(json.loads(cancelled["stale_cancel"])["error_pattern"] == "operation_id_mismatch", cancelled)
            require(json.loads(cancelled["matching_cancel"])["success"] is True, cancelled)

            option = api("SetOption", "SkipReopenBeforeMerge", True)
            require(option.get("success"), option)
            for case in ("failed_build_queue", "failed_merge_reset", "failed_merge_resume"):
                value = probe(case)
                require(value["active"] is False and value["pending"] is False, value)
                require(value["reset_calls"] == 0, value)
                require(value["completion"].get("success") is False, value)
                require(value["completion"].get("errorNumber") == -2147219094, value)
                # vbObjectError + 2410 establishes this was the failed handoff,
                # rather than a source-folder, fixture or compile refusal.
            option = api("SetOption", "SkipReopenBeforeMerge", False)
            require(option.get("success"), option)
            reservation = probe("rebuild_reservation")
            reserved_result = json.loads(reservation["outcome"])
            require(reserved_result["helper_launched"] is False and reservation["launch_count"] == 0, reservation)
            require(reserved_result["error_pattern"] == "timer_reservation_failed", reservation)
            require(reservation["active_after_fixture_cleanup"] is False, reservation)
            same_id = probe("rebuild_child_identity")
            same_id_result = json.loads(same_id["outcome"])
            require(same_id_result["error_pattern"] == "rebuild_child_identity_required", same_id)
            require(same_id_result["helper_launched"] is False and same_id["launch_count"] == 0, same_id)
            require(same_id["active_after_fixture_cleanup"] is False, same_id)
            activation = probe("rebuild_activation")
            activated_result = json.loads(activation["outcome"])
            require(activated_result["helper_launched"] is True and activation["launch_count"] == 1, activation)
            require(activated_result["build_operation_id"] == activation["expected_build_operation_id"], activation)
            require(activated_result["build_operation_id"] != activated_result["operation_id"], activation)
            require(activated_result["error_pattern"] == "timer_install_failed_after_launch", activation)
            require(activation["reservation_held"] is True, activation)
            require(activation["active_after_due_heartbeat"] is True and activation["handoff_hold"] is True, activation)
            require(bool(activation["timer_reservation_after_due_heartbeat"]), activation)
            require(activation["held_root_restored"] is True and activation["restored_handoff_hold"] is True, activation)
            require(activation["restored_active_after_due_heartbeat"] is True, activation)
            require(activation["fixture_helper_original_handle_exit"] is True, activation)
            require(activation["fixture_only_reservation_released"] is True, activation)
            require(activation["active_after_fixture_cleanup"] is False, activation)
            ribbon = probe("ribbon_queue_failure")
            require(ribbon["accepted"] is False and ribbon["active"] is False and ribbon["pending"] is False, ribbon)
            require(ribbon["completion"].get("errorNumber") == -2147219094, ribbon)
            statements = decode(call_probe("RequestPolicyStatementProbe"))
            record["statement_calls"] = statements
            require(statements["direct"]["success"] and statements["api"]["success"] and not statements["active"], statements)
            require(not statements["pending"], statements)
            for key in ("direct_build", "api_build"):
                require(statements[key].get("success") is False, statements)
                require(statements[key].get("errorNumber") == -2147219094, statements)

            sync_id = "sync-mcp-" + uuid.uuid4().hex
            sync = decode(app.Run(library + ".APIExecuteMCP", envelope, sync_id, "ExportObject", "queries", "qryF3RequestFixture", "block"))
            reused = decode(app.Run(library + ".APIExecuteMCP", envelope, sync_id, "ExportObject", "queries", "qryF3RequestFixture", "block"))
            record["sync_mcp"] = {"id": sync_id, "outcome": sync, "reused": reused}
            require(sync["operation_id"] == sync_id and sync["initiating_source"] == "mcp", sync)
            require(reused["error_pattern"] == "operation_id_reused", reused)

            typed_id = "typed-async-" + uuid.uuid4().hex
            start = len(endpoint.callbacks)
            admission = decode(app.Run(library + ".APIExecuteAsync", envelope, endpoint.info(typed_id), "ExportByType", ["queries"], False, "block"))
            require(admission.get("async") is True, admission)
            deadline = time.monotonic() + 30
            while time.monotonic() < deadline:
                pythoncom.PumpWaitingMessages()
                terminals = [c for c in endpoint.callbacks[start:]
                             if c["type"] in ("complete", "error", "cancelled")]
                if terminals:
                    break
                time.sleep(0.02)
            record["typed_async_dispatch"] = {"admission": admission, "terminals": terminals}
            require(len(terminals) == 1 and terminals[0].get("success") is True, terminals)
            require(terminals[0]["operation_id"] == typed_id and terminals[0]["decision_policy"] == "block", terminals)
            require(terminals[0]["initiating_source"] == "mcp", terminals)
            concurrent_id = "concurrent-claim-" + uuid.uuid4().hex
            barrier = threading.Barrier(2)

            def concurrent_claim(number):
                pythoncom.CoInitialize()
                peer = peer_handle = None
                details = {"number": number}
                try:
                    peer = win32com.client.DispatchEx("Access.Application")
                    peer_pid = pid_from_access_app(peer)
                    peer_stamp = process_create_time(peer_pid)
                    require(peer_pid and peer_stamp, "Could not identify a concurrent claim host")
                    details.update(pid=peer_pid, create_time=peer_stamp)
                    peer_handle = win32api.OpenProcess(0x1000 | 0x100000, False, peer_pid)
                    peer.NewCurrentDatabase(str(folder / ("claim-" + str(number) + ".accdb")))
                    peer.Run(library + ".Preload")
                    barrier.wait(timeout=30)
                    details["outcome"] = decode(peer.Run(library + ".ClaimRequestIdentityProbe", concurrent_id))
                    require(pid_from_access_app(peer) == peer_pid and process_create_time(peer_pid) == peer_stamp,
                            "Concurrent claim host identity changed")
                    peer.Quit(2)
                finally:
                    if peer is not None:
                        try:
                            peer.Quit(2)
                        except Exception as exc:
                            details["quit_error"] = repr(exc)
                    peer = None
                    gc.collect()
                    details["original_handle_exit"] = peer_handle is not None and win32event.WaitForSingleObject(peer_handle, EXIT_WAIT_MS) == win32event.WAIT_OBJECT_0
                    if peer_handle is not None:
                        peer_handle.Close()
                    pythoncom.CoUninitialize()
                require(details["original_handle_exit"], details)
                return details

            with ThreadPoolExecutor(max_workers=2) as executor:
                claims = list(executor.map(concurrent_claim, (1, 2)))
            record["concurrent_claim"] = {"operation_id": concurrent_id, "hosts": claims}
            require(sum(item["outcome"].get("success") is True for item in claims) == 1, claims)
            require(sum(item["outcome"].get("error_pattern") == "operation_id_reused" for item in claims) == 1, claims)
            record["checks"]["request_ownership"] = True
        ended = api("EndSession")
        require(ended.get("success"), ended)
        record["checks_completed"] = True
        record["verified"] = False
        save()
    finally:
        # Keep a failed acceptance assertion primary even when shutdown also fails.
        # Each cleanup stage still records its independent evidence and error.
        primary_error = sys.exc_info()[1]
        if primary_error is not None:
            record["primary_error"] = {
                "type": type(primary_error).__name__, "message": str(primary_error),
            }
        cleanup_errors = []

        def cleanup_error(stage, error):
            cleanup_errors.append({"stage": stage, "error": repr(error)})

        cleanup_start = time.monotonic()
        record["cleanup_stages"] = []

        def cleanup_checkpoint(stage):
            record["cleanup_stages"].append({
                "stage": stage,
                "elapsed_ms": round((time.monotonic() - cleanup_start) * 1000),
            })
            try:
                (evidence / (label + ".json")).write_text(
                    json.dumps(record, indent=2, default=str), encoding="utf-8")
            except Exception as exc:
                cleanup_error("checkpoint_write", exc)

        quit_attempted = False
        record["close_current_database_attempted"] = False
        record["close_current_database_returned"] = False
        cleanup_checkpoint("cleanup_started")
        if app is not None:
            owned_identity = False
            try:
                require(pid_from_access_app(app) == pid and process_create_time(pid) == stamp,
                        "Owned Access process identity changed")
                owned_identity = True
                cleanup_checkpoint("owned_identity_verified")
            except Exception as exc:
                cleanup_error("owned_identity", exc)
            if owned_identity:
                # Record scalar names only; release the collection before closing
                # the database so this diagnostic cannot retain a form proxy.
                forms = None
                cleanup_checkpoint("forms_inspection_started")
                try:
                    forms = app.Forms
                    record["forms_before_close"] = [
                        str(forms.Item(index).Name) for index in range(int(forms.Count))
                    ]
                except Exception as exc:
                    record["forms_inspection_error"] = repr(exc)
                    cleanup_error("forms_inspection", exc)
                finally:
                    forms = None
                cleanup_checkpoint("forms_inspection_finished")
                try:
                    current_host = str(app.CurrentProject.FullName)
                    expected_host = folder / "fixture.accdb"
                    record["current_host_before_close"] = current_host
                    record["expected_host_before_close"] = str(expected_host)
                    require(Path(current_host).resolve() == expected_host.resolve(),
                            {"current_host": current_host, "expected_host": str(expected_host)})
                    # The library owns its hidden form independently of the current
                    # user database. Close that fixture window in its own context.
                    cleanup_library = str(Path(os.environ["APPDATA"]) / "MSAccessVCS" / "Version Control")
                    try:
                        cleanup_checkpoint("addin_form_close_started")
                        dismissed = decode(app.Run(cleanup_library + ".APIExecute", envelope, "DismissAddinDialog", "frmVCSMain", "close"))
                        record["addin_form_close_result"] = dismissed
                        cleanup_checkpoint("addin_form_close_returned")
                        require((dismissed.get("success") is True and dismissed.get("closed") is True)
                                or dismissed.get("error_pattern") == "not_open", dismissed)
                    except Exception as exc:
                        cleanup_error("addin_form_close", exc)
                    try:
                        cleanup_checkpoint("addin_form_inspection_started")
                        after_form_close = decode(call_probe("RequestPolicyF4Probe", "inspect", ""))
                        record["addin_after_form_close"] = after_form_close
                        cleanup_checkpoint("addin_form_inspection_returned")
                        require(after_form_close.get("active") is False
                                and after_form_close.get("pending") is False
                                and after_form_close.get("form_loaded") is False, after_form_close)
                    except Exception as exc:
                        cleanup_error("addin_form_inspection", exc)
                    record["close_current_database_attempted"] = True
                    cleanup_checkpoint("database_close_started")
                    app.CloseCurrentDatabase()
                    record["close_current_database_returned"] = True
                    cleanup_checkpoint("database_close_returned")
                except Exception as exc:
                    record["close_current_database_error"] = repr(exc)
                    cleanup_error("database_close", exc)
                # Capture scalar diagnostics only, and release every inspection
                # proxy before Quit so the observation cannot retain the host.
                vbe = projects = project = None
                cleanup_checkpoint("vba_inspection_started")
                vba_state = {"projects": []}
                record["vba_before_quit"] = vba_state
                try:
                    vbe = app.VBE
                    try:
                        vba_state["mode"] = int(vbe.Mode)
                    except Exception as exc:
                        vba_state["mode_error"] = repr(exc)
                    projects = vbe.VBProjects
                    for project_index in range(1, int(projects.Count) + 1):
                        project = projects.Item(project_index)
                        project_state = {}
                        for property_name in ("Name", "FileName", "Mode"):
                            try:
                                value = getattr(project, property_name)
                                project_state[property_name.lower()] = int(value) if property_name == "Mode" else str(value)
                            except Exception as exc:
                                project_state[property_name.lower() + "_error"] = repr(exc)
                        vba_state["projects"].append(project_state)
                        project = None
                except Exception as exc:
                    vba_state["inspection_error"] = repr(exc)
                finally:
                    value = project = projects = vbe = None
                cleanup_checkpoint("vba_inspection_finished")
                # A close failure must not skip Quit on this identified process.
                try:
                    require(pid_from_access_app(app) == pid and process_create_time(pid) == stamp,
                            "Owned Access process identity changed before Quit")
                    quit_attempted = True
                    cleanup_checkpoint("quit_started")
                    app.Quit(2)
                    cleanup_checkpoint("quit_returned")
                except Exception as exc:
                    record["quit_error"] = repr(exc)
                    cleanup_error("quit", exc)
        cleanup_checkpoint("application_release_started")
        database = forms = None
        app = None
        cleanup_checkpoint("application_released")
        cleanup_checkpoint("gc_started")
        try:
            gc.collect()
        except Exception as exc:
            cleanup_error("gc", exc)
        cleanup_checkpoint("gc_finished")
        cleanup_checkpoint("com_uninitialize_started")
        try:
            pythoncom.CoUninitialize()
            record["com_uninitialized"] = True
        except Exception as exc:
            cleanup_error("com_uninitialize", exc)
        cleanup_checkpoint("com_uninitialize_finished")
        exited = False
        wait_ms = EXIT_WAIT_MS if quit_attempted else 0
        record["exit_wait_ms"] = wait_ms
        if handle is not None:
            try:
                cleanup_checkpoint("original_process_wait_started")
                exit_result = win32event.WaitForSingleObject(handle, wait_ms)
                exited = exit_result == win32event.WAIT_OBJECT_0
                record["exit_timeout"] = exit_result == win32event.WAIT_TIMEOUT
                cleanup_checkpoint("original_process_wait_returned")
            except Exception as exc:
                cleanup_error("exit_wait", exc)
            try:
                handle.Close()
            except Exception as exc:
                cleanup_error("handle_close", exc)
        record["original_process_exited"] = exited

        # Original-handle exit remains mandatory. Observe transient helpers for a
        # bounded interval and require two consecutive known empty inventories.
        samples = []
        settle_start = time.monotonic()
        settle_deadline = settle_start + (INVENTORY_SETTLE_MS / 1000 if exited else 0)
        empty_samples = 0
        inventory_unknown = False
        while True:
            try:
                pids = list_access_pids_or_none()
            except Exception as exc:
                cleanup_error("process_inventory", exc)
                pids = None
            samples.append({"elapsed_ms": round((time.monotonic() - settle_start) * 1000),
                            "pids": sorted(pids) if pids is not None else None})
            if pids is None:
                inventory_unknown = True
                break
            empty_samples = empty_samples + 1 if pids == set() else 0
            if empty_samples >= 2 or time.monotonic() >= settle_deadline:
                break
            time.sleep(0.25)
        record["inventory_settle_ms"] = INVENTORY_SETTLE_MS if exited else 0
        record["inventory_samples"] = samples
        record["pids_after_exit"] = sorted(pids) if pids is not None else None
        closure_confirmed = bool(exited and not inventory_unknown and empty_samples >= 2)
        record["closure_confirmed"] = closure_confirmed
        if pid is not None and not closure_confirmed:
            cleanup_error("closure", AssertionError("Owned Access closure was not confirmed; host artifacts were retained"))
        if endpoint is not None:
            try:
                endpoint.close()
            except Exception as exc:
                cleanup_error("callback_endpoint", exc)
        record["cleanup_errors"] = cleanup_errors
        record["checks_completed"] = bool(record.get("checks_completed"))
        record["selected_checks_verified"] = bool(record["checks_completed"] and closure_confirmed and not cleanup_errors)
        record["verified"] = bool(record["selected_checks_verified"] and case_set == "all" and not probe_bridge)
        try:
            (evidence / (label + ".json")).write_text(
                json.dumps(record, indent=2, default=str), encoding="utf-8")
        except Exception as exc:
            cleanup_error("evidence_write", exc)
        if primary_error is None:
            require(not cleanup_errors, {"cleanup_errors": cleanup_errors, "evidence": str(evidence / (label + ".json"))})



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--evidence", type=Path, required=True)
    parser.add_argument("--label", required=True)
    parser.add_argument("--client-source", type=Path, required=True)
    parser.add_argument("--case-set", choices=("all", "preflight", "ownership"), default="all",
                        help="Select a diagnostic probe group; only all can qualify the complete contract")
    parser.add_argument("--probe-bridge", action="store_true",
                        help="Diagnose probe invocation through a public disposable-host module; cannot fully qualify")
    args = parser.parse_args()
    require(__debug__ and sys.flags.optimize == 0, "Run this driver without Python optimization")
    require(Path(args.label).name == args.label, "The label must be a filename")
    evidence = args.evidence.resolve()
    require(not evidence.is_relative_to(Path(__file__).resolve().parents[1]),
            "Evidence must be outside the source checkout")
    evidence.mkdir(parents=True, exist_ok=True)
    sys.path.insert(0, str(args.client_source.resolve()))
    run(evidence, args.label, args.case_set, args.probe_bridge)

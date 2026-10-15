"""Live Policy preflight regression, using a disposable project without modTestAssert.

Run with the MCP repo's Python environment after rebuilding this worktree:
  python tools/verify-caller-error-translation.py --binary "<closed source-built .accda>" --evidence "<new evidence folder>"

The development copy must be closed before copying it. This script owns one
Access instance, calls it through MCP, and quits only that instance in cleanup.
The library is a disposable copy of the rebuilt development add-in; the installed
add-in is never patched. VBA tests cover import error journals on the development copy.
"""

import asyncio
import os
import json
import shutil
import tempfile
from contextlib import ExitStack
from pathlib import Path

import pythoncom
import win32com.client

import argparse
import gc
import sys
from msaccess_vcs_mcp.access_com.process_qos import list_access_pids_or_none, pid_from_access_app
from msaccess_vcs_mcp.access_com.instance_registry import process_create_time
from msaccess_vcs_mcp import tools


REFUSAL = 'Le module "modTestAssert" manque. C:\\source\r\nInstallation refusee.'

WRAPPER = '''Option Compare Database
Option Explicit
Public Function CheckRefusal(ByVal strPolicy As String) As String
    Dim dResult As Object
    Dim cmp As Object
    Dim blnHelper As Boolean
    Operation.Source = eosUserInterface
    Translation.SetLanguage "POL_TEST"
    Set dResult = ParseJson(VCS.RunFilteredTests(strPolicy))
    dResult.Add "operationResult", CLng(Operation.Result)
    dResult.Add "operationReady", (Operation.Status = eosReady)
    dResult.Add "canceled", (Operation.Result = eorCanceled)
    dResult.Add "decisionRequiredOutcome", (Operation.Result = eorDecisionRequired)
    For Each cmp In CurrentVBProject.VBComponents
        If cmp.Name = "modTestAssert" Then blnHelper = True
    Next cmp
    dResult.Add "helperInstalled", blnHelper
    CheckRefusal = ConvertToJson(dResult)
End Function
'''


def add_translation(database):
    database.Execute("INSERT INTO tblLanguages (ID) VALUES ('POL_TEST')", 128)
    strings = database.OpenRecordset("tblStrings")
    strings.AddNew()
    strings.Fields("msgid").Value = "modTestAssert is not installed."
    strings.Update()
    strings.Bookmark = strings.LastModified
    string_id = strings.Fields("ID").Value
    strings.Close()
    translations = database.OpenRecordset("tblTranslation")
    translations.AddNew()
    translations.Fields("Language").Value = "POL_TEST"
    translations.Fields("StringID").Value = string_id
    translations.Fields("Translation").Value = REFUSAL
    translations.Update()
    translations.Close()


async def check(database_path, policy):
    payload = await tools.vcs_call_vba(str(database_path), "PolicyLibrary.CheckRefusal", args=[policy])
    assert payload["success"], payload
    decoded = json.loads(payload["result"])
    assert decoded["success"] is False, decoded
    assert decoded["operationReady"] is True, decoded
    assert decoded["helperInstalled"] is False, decoded
    if policy == "decline":
        assert decoded["error"] == REFUSAL, decoded
        assert "error_pattern" not in decoded, decoded
        assert decoded["canceled"] is True, decoded
    else:
        assert decoded["decision_required"] is True, decoded
        assert decoded["error_pattern"] == "decision_required", decoded
        assert decoded["decisionRequiredOutcome"] is True, decoded
        assert decoded["decisions"], decoded
    (EVIDENCE/(policy+".json")).write_text(json.dumps(decoded,indent=2),encoding="utf-8")
    print(json.dumps({"policy": policy, "decoded": decoded}, ensure_ascii=False), flush=True)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--binary',type=Path,required=True)
    parser.add_argument('--evidence',type=Path,required=True)
    args=parser.parse_args()
    global EVIDENCE
    EVIDENCE=args.evidence
    root=Path(__file__).resolve().parents[1]
    assert list_access_pids_or_none()==set()
    assert args.binary.is_file() and Path(str(args.binary)+'.src').is_dir()
    assert not args.evidence.resolve().is_relative_to(root)
    args.evidence.mkdir(parents=True,exist_ok=False)
    scratch = args.evidence
    scratch.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    app = None
    try:
        with ExitStack() as cleanup:
            directory = scratch
            folder = Path(directory)
            library = folder / "Version Control.accda"
            database_path = folder / "localized.accdb"
            shutil.copy2(args.binary, library)
            engine = win32com.client.Dispatch("DAO.DBEngine.120")
            database = engine.OpenDatabase(str(library))
            add_translation(database)
            database.Close()
            del database, engine
            app = win32com.client.DispatchEx("Access.Application")
            pid=pid_from_access_app(app)
            stamp=process_create_time(pid)
            assert pid and stamp
            app.OpenCurrentDatabase(str(library))
            project = next(
                project for project in app.VBE.VBProjects
                if Path(project.FileName).resolve() == library.resolve()
            )
            project.Name = "PolicyLibrary"
            component = project.VBComponents.Add(1)  # vbext_ct_StdModule
            component.Name = "modPolicy"
            if component.CodeModule.CountOfLines:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(WRAPPER.replace("\n", "\r\n"))
            app.DoCmd.Save(5, "modPolicy")  # acModule
            del component, project
            try:
                app.DoCmd.RunCommand(126)  # acCmdCompileAndSaveAllModules
                app.CloseCurrentDatabase()
                app.NewCurrentDatabase(str(database_path))
                app.References.AddFromFile(str(library))
                app.Visible = True
                app.UserControl = True
                async def sequence():
                    from msaccess_vcs_mcp.callback_server import CallbackServer
                    from msaccess_vcs_mcp.operation_manager import OperationManager
                    manager=OperationManager.get_instance()
                    manager.set_event_loop(asyncio.get_running_loop())
                    server=CallbackServer(callback_router=manager.route_callback,
                        cancel_checker=manager.is_cancelled,cancel_requester=manager.request_cancel,
                        host='127.0.0.1',port=0)
                    server.start()
                    previous=os.environ.get('ACCESS_VCS_CALLBACK_URL')
                    os.environ['ACCESS_VCS_CALLBACK_URL']=server.callback_url
                    try:
                        await check(database_path,"decline")
                        await check(database_path,"block")
                    finally:
                        server.stop()
                        if previous is None: os.environ.pop('ACCESS_VCS_CALLBACK_URL',None)
                        else: os.environ['ACCESS_VCS_CALLBACK_URL']=previous
                asyncio.run(sequence())
            finally:
                assert pid_from_access_app(app)==pid and process_create_time(pid)==stamp
                app.Quit(2)  # acQuitSaveNone
                app = None
        print("Policy localized preflight checks passed.")
    finally:
        if app is not None:
            app.Quit(2)
        gc.collect()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()

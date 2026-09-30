"""Live A26 preflight regression, using a disposable project without modTestAssert.

Run with the MCP repo's Python environment after rebuilding this worktree:
  python tools/verify-caller-error-translation.py

The development copy must be closed before copying it. This script owns one
Access instance, calls it through MCP, and quits only that instance in cleanup.
The library is a disposable copy of the rebuilt development add-in; the installed
add-in is never patched. VBA tests cover import error journals on the development copy.
"""

import asyncio
import json
import shutil
import tempfile
from contextlib import ExitStack
from pathlib import Path

import pythoncom
import win32com.client

from msaccess_vcs_mcp.cli import run_mcp_tool


REFUSAL = 'Le module "modTestAssert" manque. C:\\source\r\nInstallation refusee.'

WRAPPER = '''Option Compare Database
Option Explicit
Public Function CheckRefusal(ByVal strPolicy As String) As String
    Dim dResult As Object
    Dim cmp As Object
    Dim blnHelper As Boolean
    Operation.Source = eosUserInterface
    Translation.SetLanguage "A26_TEST"
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
    database.Execute("INSERT INTO tblLanguages (ID) VALUES ('A26_TEST')", 128)
    strings = database.OpenRecordset("tblStrings")
    strings.AddNew()
    strings.Fields("msgid").Value = "modTestAssert is not installed."
    strings.Update()
    strings.Bookmark = strings.LastModified
    string_id = strings.Fields("ID").Value
    strings.Close()
    translations = database.OpenRecordset("tblTranslation")
    translations.AddNew()
    translations.Fields("Language").Value = "A26_TEST"
    translations.Fields("StringID").Value = string_id
    translations.Fields("Translation").Value = REFUSAL
    translations.Update()
    translations.Close()


async def check(database_path, policy):
    result = await run_mcp_tool(
        "vcs_call_vba",
        {"database_path": str(database_path), "function_name": "A26Library.CheckRefusal", "args": [policy]},
    )
    payload = json.loads(next(item.text for item in result.content if item.type == "text"))
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
    print(json.dumps({"policy": policy, "decoded": decoded}, ensure_ascii=False), flush=True)


def main():
    root = Path(__file__).resolve().parents[1]
    scratch = root / "Testing" / "Fixtures" / "scratch"
    scratch.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    app = None
    try:
        with tempfile.TemporaryDirectory(prefix="a26-localization-", dir=scratch) as directory, ExitStack() as cleanup:
            folder = Path(directory)
            library = folder / "Version Control.accda"
            database_path = folder / "localized.accdb"
            shutil.copy2(root / "Version Control.accda", library)
            engine = win32com.client.Dispatch("DAO.DBEngine.120")
            database = engine.OpenDatabase(str(library))
            add_translation(database)
            database.Close()
            del database, engine
            app = win32com.client.DispatchEx("Access.Application")
            cleanup.callback(lambda: app.Quit(2) if app is not None else None)
            app.OpenCurrentDatabase(str(library))
            project = next(
                project for project in app.VBE.VBProjects
                if Path(project.FileName).resolve() == library.resolve()
            )
            project.Name = "A26Library"
            component = project.VBComponents.Add(1)  # vbext_ct_StdModule
            component.Name = "modA26"
            if component.CodeModule.CountOfLines:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(WRAPPER.replace("\n", "\r\n"))
            app.DoCmd.Save(5, "modA26")  # acModule
            del component, project
            try:
                app.DoCmd.RunCommand(126)  # acCmdCompileAndSaveAllModules
                app.CloseCurrentDatabase()
                app.NewCurrentDatabase(str(database_path))
                app.References.AddFromFile(str(library))
                app.Visible = True
                app.UserControl = True
                asyncio.run(check(database_path, "decline"))
                asyncio.run(check(database_path, "block"))
            finally:
                app.Quit(2)  # acQuitSaveNone
                app = None
        print("A26 localized preflight checks passed.")
    finally:
        if app is not None:
            app.Quit(2)
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()

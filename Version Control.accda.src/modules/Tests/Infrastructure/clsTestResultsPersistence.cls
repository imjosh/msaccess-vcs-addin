VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestResultsPersistence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Synchronous final test-result persistence regressions.
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Private Function PassingResult() As String
    PassingResult = "{""allPassed"":true,""passed"":1,""failed"":0,""totalAssertions"":2,""tests"":{""fixture"":{""status"":""PASS""}}}"
End Function

Public Sub TestPersistenceFailureRetainsAssertions()
    Dim dResult As Object
    Dim strError As String
    On Error GoTo ErrHandler
    strError = "Cannot write C:\fixture\logs: ""Permission denied""" & vbCrLf & "Error 70"
    Set dResult = ParseJson(OverlayTestResultsSaveError(PassingResult(), strError))
    TestAssert Not dResult("success"), "persistence failure cannot report success"
    TestAssert Not dResult("allPassed"), "passing assertions do not imply a successful run"
    TestAssert dResult("results_error") = strError, "actual persistence error survives JSON escaping"
    TestAssert InStr(dResult("error"), "results file could not be written") > 0, "explains the completed run's missing file"
    TestAssert dResult("passed") = 1 And dResult("failed") = 0, "retains test counts"
    TestAssert dResult("totalAssertions") = 2, "retains assertion count"
    TestAssert dResult("tests")("fixture")("status") = "PASS", "retains test details"
    Exit Sub
ErrHandler:
    TestAssert False, "final persistence result: " & Err.Number & ": " & Err.Description
End Sub

Public Sub TestPersistenceFailurePreservesPrimaryErrors()
    Dim dResult As Object
    Dim strJson As String
    On Error GoTo ErrHandler
    strJson = "{""success"":false,""decision_required"":true,""error_pattern"":""decision_required"",""error"":""Decision required"",""decisions"":[{""prompt"":""fixture""}],""runtime_error"":""runtime fixture"",""errorNumber"":91}"
    Set dResult = ParseJson(OverlayTestResultsSaveError(strJson, "save fixture"))
    TestAssert Not dResult("success") And Not dResult("allPassed"), "failed final outcome"
    TestAssert dResult("decision_required") And dResult("error_pattern") = "decision_required", "decision retains precedence"
    TestAssert dResult("error") = "Decision required", "primary explanation survives"
    TestAssert dResult("decisions")(1)("prompt") = "fixture", "decision journal survives"
    TestAssert dResult("runtime_error") = "runtime fixture" And dResult("errorNumber") = 91, "runtime error survives"
    TestAssert dResult("results_error") = "save fixture", "persistence error is secondary"
    strJson = "{""success"":false,""error"":""runtime fixture"",""errorNumber"":91}"
    Set dResult = ParseJson(OverlayTestResultsSaveError(strJson, "save fixture"))
    TestAssert dResult("error") = "runtime fixture" And dResult("errorNumber") = 91, "plain runtime failure remains primary"
    Exit Sub
ErrHandler:
    TestAssert False, "primary errors: " & Err.Number & ": " & Err.Description
End Sub

Public Sub TestSuccessfulPersistenceKeepsExistingShape()
    Dim strJson As String
    Dim dResult As Object
    On Error GoTo ErrHandler
    strJson = OverlayTestResultsSaveError(PassingResult(), "previous run")
    strJson = OverlayTestResultsSaveError(PassingResult(), vbNullString)
    TestAssert strJson = PassingResult(), "successful persistence preserves the exact result shape"
    Set dResult = ParseJson(strJson)
    TestAssert dResult("allPassed"), "passing control remains successful"
    TestAssert Not dResult.Exists("results_error"), "no stale save error"
    Exit Sub
ErrHandler:
    TestAssert False, "passing control: " & Err.Number & ": " & Err.Description
End Sub

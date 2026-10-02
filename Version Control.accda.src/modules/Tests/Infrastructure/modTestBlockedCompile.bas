Attribute VB_Name = "modTestBlockedCompile"
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")

Private m_Log As clsLog
Private m_Runner As clsTestRunner
Private m_MCP As clsMCP
Private m_Options As clsOptions

' Called only by the setup hook in a disposable host, outside a development
' test-runner root. The installed driver owns the actual prompt and compile gate.
Public Sub ConfigureBlockedCompileFixture(ByVal blnBlock As Boolean, _
    ByVal blnCompileFailure As Boolean, ByVal strResultsFile As String, _
    Optional ByVal lngCompletionFault As Long = 0)
    On Error GoTo ErrHandler
    If Operation.OperationType <> eotTestRun Or Not Operation.IsActive Then Exit Sub
    If Operation.InteractionMode <> eimNonInteractive Then Exit Sub
    Set m_Log = Log
    Set m_Runner = TestRunner
    Set m_MCP = MCP
    Set m_Options = Options
    Operation.CompletionFaultForTest = lngCompletionFault
    TestRunner.CompileFailureForTest = blnCompileFailure
    TestRunner.ResultsFileForTest = strResultsFile
    If blnBlock Then MsgBox2 "A35 confirmation", "Proceed with fixture?", , vbYesNo
    Exit Sub
ErrHandler:
    Debug.Print "A35 setup: " & Err.Number & ": " & Err.Description
End Sub

Public Function BlockedCompileSnapshot(ByVal blnAfter As Boolean) As String
    Dim dResult As Dictionary
    On Error GoTo ErrHandler
    Set dResult = ParseJson(ModeOnlyScopeProbe("snapshot"))
    dResult.Add "log_active", Log.Active
    dResult.Add "debug_suppressed", Log.SuppressDebugOutput
    If blnAfter Then
        If Not m_MCP Is Nothing Then
            dResult.Add "callback_diagnostics", m_MCP.GetDiagnosticData
            dResult.Add "terminal_posted", m_MCP.TerminalPosted
        End If
        dResult.Add "log_released", Not (m_Log Is Log)
        dResult.Add "runner_released", Not (m_Runner Is TestRunner)
        dResult.Add "mcp_released", Not (m_MCP Is MCP)
        dResult.Add "options_released", Not (m_Options Is Options)
        Set m_Log = Nothing
        Set m_Runner = Nothing
        Set m_MCP = Nothing
        Set m_Options = Nothing
    End If
    If Not Operation.LastCompletion Is Nothing Then dResult.Add "completion", Operation.LastCompletion
    BlockedCompileSnapshot = ConvertToJson(dResult)
    Exit Function
ErrHandler:
    BlockedCompileSnapshot = "{""probe_error"":""snapshot failed""}"
End Function

Public Sub TestCompletionConstructionFailureRestoresPrivateRoot()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dDone As Dictionary
    Dim lngErr As Long
    Dim strErr As String
    On Error GoTo ErrHandler
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    cOp.ResolvePrompt vbYesNo, "A35 private root", "Proceed?"
    cOp.CompletionFaultForTest = 5
    cRoot.Complete eorFailed
    Set dDone = cOp.LastCompletion
    TestAssert dDone("completion_error_number") = 5, "construction failure was exercised"
    TestAssert dDone("error_pattern") = ERR_DECISION_REQUIRED, "decision remains primary"
    TestAssert dDone("decisions").Count = 1, "journal survives construction failure"
    TestAssert cOp.Status = eosReady And Len(cOp.CurrentRootToken) = 0, "root released"
    TestAssert Not cOp.DecisionBlocked And cOp.Decisions.Count = 0, "owned scope released"
    TestAssert Not cOp.ForceUnattended And cOp.InteractionMode = eimNormal, "interaction restored"
    cRoot.Complete eorSuccess
    TestAssert cOp.LastCompletion Is dDone, "completion is idempotent after construction failure"
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "subsequent root succeeds"
    TestAssert cOp.LastDecisions.Count = 0, "subsequent journal is clean"
    TestAssert Not cOp.LastCompletion.Exists("completion_error"), "one-shot fault does not leak"
CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If lngErr <> 0 Then TestAssert False, "A35 private completion: " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub

Public Sub TestDecisionCompileDiagnosticsAreIdempotent()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dResult As Dictionary
    Dim lngErr As Long
    Dim strErr As String
    On Error GoTo ErrHandler
    Set cOp = New clsOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    cOp.ResolvePrompt vbYesNo, "A35", "Proceed?"
    Set dResult = ParseJson("{""success"":false,""cancelled"":false,""error_pattern"":""project_not_compiled"",""error"":""compile fixture"",""results_error"":""save fixture"",""logPath"":""own.log"",""resultsPath"":""partial.json""}")
    SetDecisionRequiredOutcome dResult, cOp.Decisions
    SetDecisionRequiredOutcome dResult, cOp.Decisions
    TestAssert dResult("error_pattern") = ERR_DECISION_REQUIRED, "primary decision pattern"
    TestAssert dResult("run_error_pattern") = ERR_PROJECT_NOT_COMPILED, "compile pattern preserved"
    TestAssert dResult("run_error") = "compile fixture", "compile explanation preserved"
    TestAssert Not dResult("success") And Not dResult("cancelled"), "failure without cancellation"
    TestAssert dResult("decisions").Count = 1, "journal preserved"
    TestAssert dResult("results_error") = "save fixture", "persistence diagnostic preserved"
    TestAssert dResult("logPath") = "own.log" And dResult("resultsPath") = "partial.json", "own paths preserved"
    TestAssert Not dResult.Exists("runtime_error") And Not dResult.Exists("errorNumber"), "compile failure is not a runtime exception"
CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If lngErr <> 0 Then TestAssert False, "A35 diagnostics: " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub

Public Sub TestCompletionFailureCannotReportSuccess()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dResult As Dictionary
    Dim lngErr As Long
    Dim strErr As String
    On Error GoTo ErrHandler
    Set cOp = New clsOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cOp.CompletionFaultForTest = 5
    cRoot.Complete eorSuccess
    Set dResult = ParseJson(OverlayTestCompletionFailure("{""allPassed"":true,""tests"":{}}", cOp.LastCompletion))
    TestAssert cOp.Result = eorFailed And cOp.LastCompletion("type") = "error", "construction failure is not success"
    TestAssert Not dResult("success") And Not dResult("allPassed"), "sync completion failure is not success"
    TestAssert dResult("completion_error_number") = 5, "sync diagnostic preserved"
    TestAssert cOp.Status = eosReady And Len(cOp.CurrentRootToken) = 0, "failed construction releases root"
    TestAssert dResult.Exists("tests"), "available sync results preserved"
CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If lngErr <> 0 Then TestAssert False, "A35 failure verdict: " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub

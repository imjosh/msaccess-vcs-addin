Attribute VB_Name = "modTestAsyncCallerIdentity"
'---------------------------------------------------------------------------------------
' Module    : modTestAsyncCallerIdentity
' Purpose   : Public launch/dispatch regressions. Parameterized probes run in the
'           : installed driver; assertions run in the development test host.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("integration")

Private m_strRefusedCallback As String
Private m_blnDispatchPending As Boolean
Private m_strLastProbe As String
Private m_blnExpectPendingPolicy As Boolean


' Parameterized helpers are not discovered as tests. The live driver supplies distinct
' loopback callbacks; ordinary suite runs use an immediately rejected host for refusals.
Public Function ConfigureAsyncCallerProbe(ByVal strCallbackInfo As String) As Boolean
    m_strRefusedCallback = strCallbackInfo
    ConfigureAsyncCallerProbe = True
End Function


' Both launches happen on one VBA stack, before Access can deliver the pending timer.
' Explicit dispatch uses the real timer entry point and removes timing from the test.
Public Function ProbePendingAsyncCallers(ByVal strOwnerCallback As String, _
    ByVal strRefusedCallback As String, ByVal strRefusedExportFolder As String) As String

    Dim dResult As New Dictionary
    Dim intPrior As eInteractionMode
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    intPrior = Operation.InteractionMode
    Operation.InteractionMode = eimSilent
    m_blnExpectPendingPolicy = True
    dResult.Add "owner_start", APIAsync(strOwnerCallback, "RunFilteredTests", "block", strSessionEnvelope:=CurrentCompatibilityEnvelope())
    dResult.Add "refused_start", APIAsync(strRefusedCallback, "BuildAs", strRefusedExportFolder, strRefusedExportFolder & ".accdb", strSessionEnvelope:=CurrentCompatibilityEnvelope())
    ' Dispatch belongs to its own session scope, after both launches unwind.
    LeaveCompatibilitySession
    If TimerIsPending Then WinAPITimerCallback

CleanUp:
    On Error Resume Next
    Operation.InteractionMode = intPrior
    m_blnExpectPendingPolicy = False
    If lngErr <> 0 Then dResult.Add "probe_error", CStr(lngErr) & ": " & strErr
    ProbePendingAsyncCallers = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function


' CStr(Null) fails during launch, before a timer is created. A later request must still
' be admitted, and the failed caller must receive its own terminal error.
Public Function ProbeAsyncAdmissionError(ByVal strCallbackInfo As String) As String
    Dim dResult As New Dictionary
    On Error GoTo ErrHandler
    dResult.Add "launch", APIAsync(strCallbackInfo, "RunFilteredTests", Null, strSessionEnvelope:=CurrentCompatibilityEnvelope())
    ProbeAsyncAdmissionError = ConvertToJson(dResult)
    Exit Function
ErrHandler:
    dResult.Add "caught_error", Err.Number
    ProbeAsyncAdmissionError = ConvertToJson(dResult)
End Function


' A request can be admitted while idle, then its timer delivered during a later
' synchronous API call. This exercises dispatch admission separately from launch admission.
Public Function ProbePendingBeforeSync(ByVal strRefusedCallback As String) As String
    Dim dResult As New Dictionary
    Dim lngErr As Long
    Dim strErr As String
    On Error GoTo ErrHandler
    m_blnDispatchPending = True
    dResult.Add "pending_start", APIAsync(strRefusedCallback, "BuildAs", "a33-must-not-build", "a33-must-not-build.accdb", strSessionEnvelope:=CurrentCompatibilityEnvelope())
    dResult.Add "owner_result", API("RunFilteredTests", "block", strSessionEnvelope:=CurrentCompatibilityEnvelope())
    dResult.Add "running_probe", ParseJson(m_strLastProbe)
CleanUp:
    On Error Resume Next
    m_blnDispatchPending = False
    If lngErr <> 0 Then dResult.Add "probe_error", CStr(lngErr) & ": " & strErr
    ProbePendingBeforeSync = ConvertToJson(dResult)
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function


Private Function RunningCallerSnapshot() As Dictionary
    Dim dResult As New Dictionary
    dResult.Add "operation_id", MCP.OperationId
    dResult.Add "operation_type", Operation.OperationType
    dResult.Add "callback_url", MCP.CallbackUrl
    dResult.Add "root_token", Operation.CurrentRootToken
    dResult.Add "status", Operation.Status
    dResult.Add "source", Operation.Source
    dResult.Add "result", Operation.Result
    dResult.Add "mode", Operation.InteractionMode
    dResult.Add "policy", Operation.DecisionPolicy
    dResult.Add "journal", ConvertToJson(Operation.Decisions)
    dResult.Add "log_path", Log.LogFilePath
    dResult.Add "saved_log_path", Log.SavedLogFilePath
    dResult.Add "terminal_posted", MCP.TerminalPosted
    Set RunningCallerSnapshot = dResult
End Function


' Called directly in the installed project while its API is executing a test in the
' development project. It does not acquire, finish, or release the driver's root.
Public Function ProbeRunningAsyncCaller(ByVal blnUnused As Boolean) As String
    Dim dResult As New Dictionary
    Dim dLaunch As Dictionary
    Dim cOutcome As clsPromptOutcome
    Dim strCallback As String
    Dim strEnvelope As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    Set cOutcome = Operation.ResolvePrompt(vbOKOnly, "A33 owner journal", "Owner acknowledgement")
    strCallback = m_strRefusedCallback
    If Len(strCallback) = 0 Then
        strCallback = "{""callback_url"":""http://invalid..host/callback"",""operation_id"":""a33-refused""}"
    End If
    dResult.Add "before", RunningCallerSnapshot()
    If m_blnDispatchPending Then
        dResult.Add "deferred_dispatch", TimerIsPending
        If TimerIsPending Then
            ' Keep the API root busy while the timer validates its own envelope.
            strEnvelope = CurrentCompatibilityEnvelope()
            LeaveCompatibilitySession
            WinAPITimerCallback
            strErr = EnterCompatibilitySession(strEnvelope)
            If Len(strErr) > 0 Then dResult.Add "probe_error", strErr
        End If
        dResult.Add "timer_pending_after", TimerIsPending
    Else
        Set dLaunch = ParseJson(APIAsync(strCallback, "BuildAs", "a33-must-not-build", "a33-must-not-build.accdb", _
            strSessionEnvelope:=CurrentCompatibilityEnvelope()))
        dResult.Add "launch", dLaunch
        ' The old implementation admits this timer, then overwrites MCP before API refuses.
        If dLaunch.Exists("async") Then WinAPITimerCallback
    End If
    dResult.Add "after", RunningCallerSnapshot()
    MCP.SkipCancelThrottleForTest
    dResult.Add "cancelled", MCP.CheckCancelled

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then dResult.Add "probe_error", CStr(lngErr) & ": " & strErr
    m_strLastProbe = ConvertToJson(dResult)
    ProbeRunningAsyncCaller = m_strLastProbe
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function


' The live driver records snapshots after the terminal result, outside any API stack.
Public Function LastAsyncCallerProbe(ByVal blnUnused As Boolean) As String
    LastAsyncCallerProbe = m_strLastProbe
End Function


' Read the actual driver's state: the development project's singleton is idle.
Public Function ReadAsyncOwnerForTest(ByVal blnUnused As Boolean) As String
    Dim dResult As Dictionary
    Set dResult = RunningCallerSnapshot()
    dResult.Add "expect_pending_policy", m_blnExpectPendingPolicy
    ReadAsyncOwnerForTest = ConvertToJson(dResult)
End Function


Public Sub TestAsyncOwnerCompletes()
    Dim dResult As Dictionary
    Dim strDriver As String
    On Error GoTo ErrHandler
    strDriver = FSO.BuildPath(GetInstallSettings.strInstallFolder, ADDIN_BASENAME)
    Set dResult = ParseJson(Application.Run(strDriver & ".ReadAsyncOwnerForTest", True))
    TestAssert dResult("operation_type") = eotTestRun, "the first request retained its RunFilteredTests method"
    TestAssert Len(dResult("root_token")) > 0, "the owner holds its running lease"
    If dResult("expect_pending_policy") Then
        TestAssert dResult("policy") = edpBlock, "the first request retained its block argument"
        TestAssert dResult("mode") = eimNonInteractive, "the first request's policy took effect"
    End If
    Exit Sub
ErrHandler:
    TestAssert False, "owner test: " & Err.Description
End Sub


Public Sub TestRefusedLaunchPreservesRunningOwner()
    Dim dResult As Dictionary
    Dim dBefore As Dictionary
    Dim dAfter As Dictionary
    Dim varKey As Variant
    Dim strDriver As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    strDriver = FSO.BuildPath(GetInstallSettings.strInstallFolder, ADDIN_BASENAME)
    Set dResult = ParseJson(Application.Run(strDriver & ".ProbeRunningAsyncCaller", True))
    If dResult.Exists("probe_error") Then
        TestAssert False, "running probe: " & dResult("probe_error")
        GoTo CleanUp
    End If
    TestAssert True, "running probe completed"
    Set dBefore = dResult("before")
    Set dAfter = dResult("after")
    TestAssert Len(dBefore("root_token")) > 0, "probe reached the running driver's root"
    For Each varKey In dBefore.Keys
        TestAssert dBefore(varKey) = dAfter(varKey), "running owner retained " & CStr(varKey)
    Next varKey
    If dResult.Exists("deferred_dispatch") Then
        TestAssert Not dResult("timer_pending_after"), "the admitted timer was consumed during the running API"
    Else
        TestAssert dResult("launch").Exists("error_pattern"), "launch refused before dispatch"
        If dResult("launch").Exists("error_pattern") Then
            TestAssert dResult("launch")("error_pattern") = ERR_OPERATION_ALREADY_RUNNING, "stable busy pattern"
        End If
    End If
    TestAssert Not dResult("cancelled"), "cancelling the refused caller cannot cancel the owner"

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "running caller test " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub


' External idle-host probe only. Never call this from a running test harness.
' Admit the timer first, then hold a real root lease until its dispatch is refused.
Public Function ProbeDeferredBusyRoot(ByVal strOwnerCallback As String, _
    ByVal strRefusedCallback As String, ByVal strLogPath As String) As String

    Dim dResult As New Dictionary
    Dim cRoot As clsRootOperationLease
    Dim cOutcome As clsPromptOutcome
    Dim lngErr As Long
    Dim strErr As String
    On Error GoTo ErrHandler
    If Operation.IsActive Or TimerIsPending Then Err.Raise 5, , "Probe requires an idle host"
    dResult.Add "pending_start", APIAsync(strRefusedCallback, "BuildAs", _
        "a33-must-not-build", "a33-must-not-build.accdb", strSessionEnvelope:=CurrentCompatibilityEnvelope())
    MCP.RegisterCallback strOwnerCallback
    Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
    If cRoot Is Nothing Then Err.Raise 5, , "Probe root was refused"
    Set cOutcome = Operation.ResolvePrompt(vbOKOnly, "A33 owner journal", "Owner acknowledgement")
    Log.Add "A33 deferred timer owner", False
    Log.SaveFile strLogPath
    dResult.Add "before", RunningCallerSnapshot()
    LeaveCompatibilitySession
    If TimerIsPending Then WinAPITimerCallback
    dResult.Add "timer_pending_after", TimerIsPending
    dResult.Add "after", RunningCallerSnapshot()
    MCP.SkipCancelThrottleForTest
    dResult.Add "cancelled", MCP.CheckCancelled
    cRoot.Complete eorSuccess
    Set cRoot = Nothing
    dResult.Add "clean", RunningCallerSnapshot()
CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If lngErr <> 0 Then dResult.Add "probe_error", CStr(lngErr) & ": " & strErr
    ProbeDeferredBusyRoot = ConvertToJson(dResult)
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function

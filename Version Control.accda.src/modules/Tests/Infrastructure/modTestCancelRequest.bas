Attribute VB_Name = "modTestCancelRequest"
'---------------------------------------------------------------------------------------
' Module    : modTestCancelRequest
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Tests for the MCP cancel checkpoint (clsOperation.CheckCancelRequest),
'           : the poll behind it (clsMCP.CheckCancelled), and the request owner every
'           : route shares (clsOperation.RequestCancel).
'           : Like modTestOperationLifecycle, these run inside a test-run root, so roots
'           : are driven on private clsOperation instances. A private instance never
'           : asks MCP and routes a request only to its own error level, so nothing here
'           : can cancel the run hosting it. A runner stopped by a request is checked
'           : live instead. See docs/noninteractive-dialogs.md.
'           : Run: ?VCS.RunTests("modTestCancelRequest")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


' Fails at once. XMLHTTP rejects the malformed host before it tries to connect.
Private Const UNREACHABLE_CALLBACK As String = "http://invalid..host/callback"

' A last poll 100 ms before midnight. Timer has since restarted near zero.
Private Const BEFORE_MIDNIGHT_MS As Double = 86399900#

' The private MCP that CancelPollProbe keeps between its steps.
Private m_cPollProbe As clsMCP


'---------------------------------------------------------------------------------------
' Procedure : NewOperation
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : A private, unattended operation instance, isolated from the session one.
'---------------------------------------------------------------------------------------
'
Private Function NewOperation() As clsOperation
    Dim cOp As clsOperation
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set NewOperation = cOp
End Function


'---------------------------------------------------------------------------------------
' Procedure : McpCounter
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : One counter from an MCP instance's diagnostic data.
'---------------------------------------------------------------------------------------
'
Private Function McpCounter(cMCP As clsMCP, strKey As String) As Long
    Dim dData As Dictionary
    Set dData = cMCP.GetDiagnosticData
    McpCounter = dData(strKey)
End Function


Public Sub TestCancelPollFakeFiresOnTheNthCheck()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotExport)
    cOp.SetCancelPollForTest 3
    TestAssert Not cOp.CheckCancelRequest, "the first check sees no request"
    TestAssert Not cOp.CheckCancelRequest, "the second check sees no request"
    TestAssert cOp.CheckCancelRequest, "the third check sees the request"
    TestAssert cOp.CancelRequested, "the root records the request"
    TestAssert cOp.CheckCancelRequest, "later checks keep reporting it"
    cRoot.Complete eorFailed

    ' The fake answers once, and a new root starts without the old request.
    Set cRoot = cOp.TryBeginRoot(eotExport)
    TestAssert Not cOp.CancelRequested, "a new root starts without a request"
    TestAssert Not cOp.CheckCancelRequest, "the fake answered only once"

    ' Zero disarms it.
    cOp.SetCancelPollForTest 1
    cOp.SetCancelPollForTest 0
    TestAssert Not cOp.CheckCancelRequest, "a disarmed fake reports nothing"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "a root with no request finishes as asked"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CancelRun
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Run a root of intType until its loop sees a cancel request, then complete
'           : it the way that operation does after a critical error. An export goes
'           : through Finish, and a build or merge goes through the build continuation.
'           : One acknowledged prompt puts a decision on the terminal callback.
'---------------------------------------------------------------------------------------
'
Private Function CancelRun(ByVal intType As eOperationType) As clsOperation

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cResumed As clsRootOperationLease
    Dim strLabel As String

    strLabel = " (type " & intType & ")"
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(intType, edpBlock)
    If intType = eotExport Then
        Set cResumed = cRoot
    Else
        Set cResumed = cOp.ResumeRoot(cRoot.DetachForContinuation())
    End If
    cOp.ResolvePrompt vbOKOnly, "Notice", "Informational"
    cOp.SetCancelPollForTest 1
    TestAssert cOp.CheckCancelRequest, "the loop sees the request" & strLabel
    TestAssert cOp.ErrorLevel = eelCritical, "the request is the critical level loops stop on" & strLabel
    If intType = eotExport Then
        cResumed.Complete
    Else
        ' What Build's CleanUp records after a critical error.
        cOp.Result = eorFailed
        cResumed.Complete BuildContinuationResult(cOp)
    End If
    Set CancelRun = cOp

End Function


Public Sub TestCancelRequestEndsExportBuildAndMergeCancelled()

    Dim varType As Variant
    Dim cOp As clsOperation
    Dim dDone As Dictionary
    Dim strLabel As String

    For Each varType In Array(eotExport, eotBuild, eotMerge)
        strLabel = " (type " & varType & ")"
        Set cOp = CancelRun(varType)
        Set dDone = cOp.LastCompletion
        TestAssert cOp.Result = eorCanceled, "the root finishes cancelled" & strLabel
        TestAssert dDone("type") = "cancelled", "the terminal callback is cancelled" & strLabel
        TestAssert dDone.Exists("log_path"), "the callback carries log_path" & strLabel
        TestAssert dDone("decisions").Count = 1, "the callback keeps the decisions" & strLabel
    Next varType

End Sub


Public Sub TestCancelRequestYieldsToDecisionsAndRuntimeErrors()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dDone As Dictionary

    ' An unanswered prompt still wins, because the caller has to answer it.
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotMerge, edpBlock)
    cOp.ResolvePrompt vbYesNo, "Confirm", "Proceed?"
    cOp.SetCancelPollForTest 1
    cOp.CheckCancelRequest
    cRoot.Complete eorFailed
    TestAssert cOp.Result = eorDecisionRequired, "decision_required wins over the cancel"

    ' A runtime error is a failure, whatever else stopped the loop.
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotExport)
    cOp.SetCancelPollForTest 1
    cOp.CheckCancelRequest
    cOp.RecordRootRuntimeError cOp.CurrentRootToken, "Type mismatch", 13
    cRoot.Complete eorFailed
    Set dDone = cOp.LastCompletion
    TestAssert cOp.Result = eorFailed, "a runtime error still fails the root"
    TestAssert dDone("type") = "error", "the callback reports the error"

    ' A test run's request goes to its runner, which saves its results and reports
    ' the run itself, so the error level and the result are left alone. (A private
    ' instance routes nothing, so the run hosting this test carries on.)
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotTestRun)
    cOp.SetCancelPollForTest 1
    TestAssert cOp.CheckCancelRequest, "a test run sees the request"
    TestAssert cOp.ErrorLevel < eelCritical, "a test run is not stopped through the error level"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "a test run keeps the result its runner reports"

End Sub


Public Sub TestCancelCheckPollsOnlyARunningRoot()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cResumed As clsRootOperationLease
    Dim cPause As clsOperationPause
    Dim strToken As String

    ' With no root there is nothing to cancel, and the armed fake is not spent.
    Set cOp = NewOperation
    cOp.SetCancelPollForTest 1
    TestAssert Not cOp.CheckCancelRequest, "no root, no request"

    ' A paused root is inside foreign code, such as a user hook.
    Set cRoot = cOp.TryBeginRoot(eotExport)
    Set cPause = cOp.TryPause()
    TestAssert Not cOp.CheckCancelRequest, "a paused root does not poll"
    cPause.ResumePause
    TestAssert cOp.CheckCancelRequest, "the fake answers the first check of the running root"
    cRoot.Complete
    TestAssert cOp.Result = eorCanceled, "the export finishes cancelled"

    ' A root staged for a timer continuation polls again once resumed.
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotMerge)
    strToken = cRoot.DetachForContinuation()
    cOp.SetCancelPollForTest 1
    TestAssert Not cOp.CheckCancelRequest, "a staged root does not poll"
    Set cResumed = cOp.ResumeRoot(strToken)
    TestAssert cOp.CheckCancelRequest, "the resumed root polls"
    cResumed.Complete eorFailed
    TestAssert cOp.Result = eorCanceled, "the resumed merge finishes cancelled"

End Sub


Public Sub TestRequestCancelOnlyTakesARunningRoot()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cResumed As clsRootOperationLease
    Dim cPause As clsOperationPause
    Dim strToken As String
    Dim eelBefore As eErrorLevel

    ' With no root the request is refused and nothing changes. Without a root the
    ' level read is the shared scope's, so compare it with itself, not a fixed value.
    Set cOp = NewOperation
    eelBefore = cOp.ErrorLevel
    TestAssert Not cOp.RequestCancel, "no root takes no request"
    TestAssert Not cOp.CancelRequested, "a refused request is not recorded"
    TestAssert cOp.ErrorLevel = eelBefore, "a refused request leaves the level alone"

    ' A running root records it once and routes it to the level its loops stop on.
    Set cRoot = cOp.TryBeginRoot(eotExport)
    TestAssert cOp.RequestCancel, "a running root takes the request"
    TestAssert cOp.CancelRequested, "the request is recorded"
    TestAssert cOp.ErrorLevel = eelCritical, "the request is the critical level loops stop on"
    cOp.ErrorLevel = eelNoError
    TestAssert cOp.RequestCancel, "a second request reports the recorded one"
    TestAssert cOp.ErrorLevel = eelNoError, "a second request is not routed again"
    TestAssert cOp.Status = eosRunning, "a request does not finish the root"
    cRoot.Complete eorFailed
    TestAssert cOp.Result = eorCanceled, "the root finishes cancelled"

    ' A finished root takes none, and the next root starts without it.
    TestAssert Not cOp.RequestCancel, "a finished root takes no request"
    Set cRoot = cOp.TryBeginRoot(eotExport)
    TestAssert Not cOp.CancelRequested, "a new root starts without the old request"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "the new root finishes as asked"

    ' A paused root is in foreign code, so it refuses until it resumes.
    Set cRoot = cOp.TryBeginRoot(eotExport)
    Set cPause = cOp.TryPause()
    TestAssert Not cOp.RequestCancel, "a paused root takes no request"
    TestAssert Not cOp.CancelRequested, "the paused root records nothing"
    cPause.ResumePause
    TestAssert cOp.RequestCancel, "the resumed root takes it"
    cRoot.Complete eorFailed
    TestAssert cOp.Result = eorCanceled, "the export finishes cancelled"

    ' A staged root waits for its continuation, so it refuses until resumed.
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotMerge)
    strToken = cRoot.DetachForContinuation()
    TestAssert Not cOp.RequestCancel, "a staged root takes no request"
    TestAssert cOp.ErrorLevel = eelNoError, "the staged root's level is left alone"
    Set cResumed = cOp.ResumeRoot(strToken)
    TestAssert cOp.RequestCancel, "the resumed root takes it"
    cResumed.Complete eorFailed
    TestAssert cOp.Result = eorCanceled, "the merge finishes cancelled"

End Sub


Public Sub TestRequestCancelYieldsToDecisionsAndRuntimeErrors()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotMerge, edpBlock)
    cOp.ResolvePrompt vbYesNo, "Confirm", "Proceed?"
    TestAssert cOp.RequestCancel, "the request is taken"
    cRoot.Complete eorFailed
    TestAssert cOp.Result = eorDecisionRequired, "decision_required wins over the request"

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotExport)
    TestAssert cOp.RequestCancel, "the request is taken"
    cOp.RecordRootRuntimeError cOp.CurrentRootToken, "Type mismatch", 13
    TestAssert cOp.HasRootRuntimeError, "the runtime error is kept"
    cRoot.Complete eorFailed
    TestAssert cOp.Result = eorFailed, "a runtime error still fails the root"

    ' A private test-run root routes nothing. Its runner reports the run.
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotTestRun)
    TestAssert cOp.RequestCancel, "a test run takes the request"
    TestAssert cOp.ErrorLevel < eelCritical, "a test run is not stopped through the error level"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "a test run keeps the result its runner reports"

End Sub


Public Sub TestInactiveMCPMakesNoCancelPoll()

    Dim cMCP As clsMCP
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim lngSessionPolls As Long

    ' No callback registered.
    Set cMCP = New clsMCP
    TestAssert Not cMCP.CheckCancelled, "an inactive MCP reports no request"
    TestAssert McpCounter(cMCP, "cancelPolls") = 0, "an inactive MCP sends no request"

    ' A callback without an operation id is not active either.
    cMCP.RegisterCallback "{""callback_url"":""" & UNREACHABLE_CALLBACK & """}"
    TestAssert Not cMCP.IsActive, "a callback needs an operation id"
    TestAssert Not cMCP.CheckCancelled, "an incomplete callback reports no request"
    TestAssert McpCounter(cMCP, "cancelPolls") = 0, "an incomplete callback sends no request"

    ' Only the session operation asks MCP. With the throttle out of the way, a poll
    ' from a private root would show on the session's counter.
    MCP.SkipCancelThrottleForTest
    lngSessionPolls = McpCounter(MCP, "cancelPolls")
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotExport)
    TestAssert Not cOp.CheckCancelRequest, "a private root sees no request"
    cRoot.Complete eorSuccess
    TestAssert McpCounter(MCP, "cancelPolls") = lngSessionPolls, "a private root sends no poll"

End Sub


Public Sub TestCancelPollErrorIsLoggedOnceAndIgnored()

    Dim cMCP As clsMCP
    Dim blnFirst As Boolean
    Dim blnSecond As Boolean
    Dim dblLogBefore As Double
    Dim dblLogAfterFirst As Double
    Dim dblLogAfterSecond As Double
    Dim lngErrorsBefore As Long
    Dim eelBefore As eErrorLevel

    Set cMCP = New clsMCP
    cMCP.RegisterCallback "{""callback_url"":""" & UNREACHABLE_CALLBACK & """,""operation_id"":""cancel-poll-unit""}"
    TestAssert cMCP.IsActive, "the callback is registered"

    ' Nothing between the reads may write to the log, so the assertions wait.
    lngErrorsBefore = Log.ErrorCount
    eelBefore = Operation.ErrorLevel
    dblLogBefore = Log.BufferChars
    blnFirst = cMCP.CheckCancelled
    dblLogAfterFirst = Log.BufferChars
    cMCP.SkipCancelThrottleForTest
    blnSecond = cMCP.CheckCancelled
    dblLogAfterSecond = Log.BufferChars

    TestAssert Not blnFirst, "a failed poll reads as no request"
    TestAssert Not blnSecond, "so does the next one"
    TestAssert McpCounter(cMCP, "cancelPolls") = 2, "both polls were sent"
    TestAssert McpCounter(cMCP, "cancelPollErrors") = 2, "both polls failed"
    TestAssert dblLogAfterFirst > dblLogBefore, "the first failure is logged"
    TestAssert dblLogAfterSecond = dblLogAfterFirst, "the second failure is not"
    TestAssert Log.ErrorCount = lngErrorsBefore, "a failed poll is not an error of the run"
    TestAssert Operation.ErrorLevel = eelBefore, "a failed poll leaves the run's outcome alone"

End Sub


Public Sub TestCancelPollThrottleHoldsAndSurvivesMidnight()

    Dim cMCP As clsMCP
    Dim blnFirst As Boolean
    Dim blnThrottled As Boolean
    Dim blnAfterMidnight As Boolean
    Dim lngPollsFirst As Long
    Dim lngPollsThrottled As Long
    Dim lngPollsAfterMidnight As Long
    Dim lngErrorsBefore As Long
    Dim eelBefore As eErrorLevel

    Set cMCP = New clsMCP
    cMCP.RegisterCallback "{""callback_url"":""" & UNREACHABLE_CALLBACK & """,""operation_id"":""cancel-poll-throttle""}"
    TestAssert cMCP.IsActive, "the callback is registered"

    ' Nothing between the reads may write to the log, so the assertions wait.
    lngErrorsBefore = Log.ErrorCount
    eelBefore = Operation.ErrorLevel
    blnFirst = cMCP.CheckCancelled
    lngPollsFirst = McpCounter(cMCP, "cancelPolls")
    blnThrottled = cMCP.CheckCancelled
    lngPollsThrottled = McpCounter(cMCP, "cancelPolls")
    cMCP.SetLastCancelCheckForTest BEFORE_MIDNIGHT_MS
    blnAfterMidnight = cMCP.CheckCancelled
    lngPollsAfterMidnight = McpCounter(cMCP, "cancelPolls")

    TestAssert Not blnFirst, "a failed poll reads as no request"
    TestAssert lngPollsFirst = 1, "the first check polls"
    TestAssert Not blnThrottled, "the throttled check reads as no request"
    TestAssert lngPollsThrottled = 1, "an immediate second check does not poll"
    TestAssert Not blnAfterMidnight, "the check after midnight reads as no request"
    TestAssert lngPollsAfterMidnight = 2, "a last poll before midnight does not hold the next one"
    TestAssert McpCounter(cMCP, "cancelPollErrors") = 2, "only the two sent polls failed"
    TestAssert Log.ErrorCount = lngErrorsBefore, "a failed poll is not an error of the run"
    TestAssert Operation.ErrorLevel = eelBefore, "a failed poll leaves the run's outcome alone"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PollStep
' Author    : Josh
' Date      : 10/21/2026
' Purpose   : One check of the probe's MCP, with its result and poll counters.
'---------------------------------------------------------------------------------------
'
Private Function PollStep() As Dictionary
    Dim dStep As New Dictionary
    dStep.Add "result", m_cPollProbe.CheckCancelled
    dStep.Add "cancelPolls", McpCounter(m_cPollProbe, "cancelPolls")
    dStep.Add "cancelPollErrors", McpCounter(m_cPollProbe, "cancelPollErrors")
    Set PollStep = dStep
End Function


'---------------------------------------------------------------------------------------
' Procedure : CancelPollProbe
' Author    : Josh
' Date      : 10/21/2026
' Purpose   : Parameterized probe for the live cancel_poll case in
'           : Testing/check_noninteractive_policies.py. It is not discovered as a test.
'           : It polls the harness endpoint through a private MCP, never the session
'           : one, so no running operation can see the answer. The steps run in order,
'           : one call each, so the harness can switch its answer between them.
'           :   first    : register strCallbackInfo, then check twice at once. The
'           :              second check falls inside the throttle.
'           :   uncached : skip the throttle and check again.
'           :   midnight : set the last check just before midnight and check again.
'           :              This step releases the probe's MCP.
'           : Returns JSON with the step, one entry per check, and any probe error.
'---------------------------------------------------------------------------------------
'
Public Function CancelPollProbe(ByVal strCallbackInfo As String, ByVal strStep As String) As String

    Dim dResult As New Dictionary
    Dim colChecks As New Collection
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    dResult.Add "step", strStep
    dResult.Add "checks", colChecks
    If strStep = "first" Then
        Set m_cPollProbe = New clsMCP
        m_cPollProbe.RegisterCallback strCallbackInfo
        dResult.Add "active", m_cPollProbe.IsActive
        colChecks.Add PollStep()
        colChecks.Add PollStep()
    ElseIf m_cPollProbe Is Nothing Then
        Err.Raise vbObjectError + 1, , "Run the first step before " & strStep
    ElseIf strStep = "uncached" Then
        m_cPollProbe.SkipCancelThrottleForTest
        colChecks.Add PollStep()
    ElseIf strStep = "midnight" Then
        m_cPollProbe.SetLastCancelCheckForTest BEFORE_MIDNIGHT_MS
        colChecks.Add PollStep()
        Set m_cPollProbe = Nothing
    Else
        Err.Raise vbObjectError + 2, , "Unknown step " & strStep
    End If

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then dResult.Add "probe_error", CStr(lngErr) & ": " & strErr
    CancelPollProbe = ConvertToJson(dResult)
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function

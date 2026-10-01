Attribute VB_Name = "modTestCancelRequest"
'---------------------------------------------------------------------------------------
' Module    : modTestCancelRequest
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Tests for the MCP cancel checkpoint (clsOperation.CheckCancelRequest) and
'           : the poll behind it (clsMCP.CheckCancelled).
'           : Like modTestOperationLifecycle, these run inside a test-run root, so roots
'           : are driven on private clsOperation instances. A private instance never
'           : asks MCP and routes a request only to its own error level, so nothing here
'           : can cancel the run hosting it. A runner stopped by a request is checked
'           : live instead; see docs/noninteractive-dialogs.md.
'           : Run: ?VCS.RunTests("modTestCancelRequest")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


' Fails at once: the malformed host is rejected before any connection is attempted.
Private Const UNREACHABLE_CALLBACK As String = "http://invalid..host/callback"


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
'           : it the way that operation does after a critical error: an export through
'           : Finish, a build or merge through the build continuation. One acknowledged
'           : prompt puts a decision on the terminal callback.
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

    ' An unanswered prompt still wins: the caller has to answer it.
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

    ' No root: nothing to cancel, and the armed fake is not spent.
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
    cMCP.RegisterCallback "{""callback_url"":""" & UNREACHABLE_CALLBACK & """,""operation_id"":""a28-unit""}"
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

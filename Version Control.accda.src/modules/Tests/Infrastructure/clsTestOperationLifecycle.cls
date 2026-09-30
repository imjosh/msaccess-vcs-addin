VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestOperationLifecycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestOperationLifecycle
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Wiring tests for policy requests against a running operation: a policy request
'           : made while an operation runs is refused without changing that operation's
'           : interaction state. A session policy is caller-owned: it outlives the
'           : operations that run under it and closes only when the caller clears it.
'           : Every test that needs a running operation uses a private
'           : instance, so no result depends on where the host run started. The session
'           : operation is only read, by the clearing test. Start, restore and exit paths
'           : of MergeBuild and RunFilteredTests need a free operation and are covered by
'           : the scripted pass in X01.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")


' Live state captured before each test and checked again after it.
Private m_Mode As eInteractionMode
Private m_Policy As eDecisionPolicy
Private m_Blocked As Boolean
Private m_ErrorLevel As eErrorLevel


Private Sub Class_Initialize()
    m_Mode = Operation.InteractionMode
    m_Policy = Operation.DecisionPolicy
    m_Blocked = Operation.DecisionBlocked
    m_ErrorLevel = Operation.ErrorLevel
End Sub


Private Function LiveStateUnchanged() As Boolean
    LiveStateUnchanged = (Operation.InteractionMode = m_Mode) _
        And (Operation.DecisionPolicy = m_Policy) _
        And (Operation.DecisionBlocked = m_Blocked) _
        And (Operation.ErrorLevel = m_ErrorLevel)
End Function


' A private operation, so a session policy set on it never reaches the live operation.
' The policy scope goes with the instance, even when a test stops before clearing it.
Private Function NewUnattendedOperation() As clsOperation
    Set NewUnattendedOperation = New clsOperation
    NewUnattendedOperation.ForceUnattended = True
End Function


Public Sub TestBeginWithPolicyRefusedWhileRunning()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim intPolicy As eDecisionPolicy
    Dim intPolicyBefore As eDecisionPolicy
    Dim intMode As eInteractionMode
    Dim blnBlocked As Boolean
    Dim intLevel As eErrorLevel

    ' A private instance, so the result does not depend on whether the session operation
    ' happens to hold the host run's root. Through the MCP the runner lives in the
    ' installed add-in and this project's session operation only adopts that root from
    ' the registry when it is created after the root begins.
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "the first root is granted"
    intMode = cOp.InteractionMode
    intPolicyBefore = cOp.DecisionPolicy
    blnBlocked = cOp.DecisionBlocked
    intLevel = cOp.ErrorLevel

    For intPolicy = edpBlock To edpDecline
        TestAssert Not cOp.Begin(eotOther, intPolicy), _
            "a policy request is refused while an operation is running (policy " & intPolicy & ")"
        TestAssert (cOp.InteractionMode = intMode) And (cOp.DecisionBlocked = blnBlocked) _
            And (cOp.ErrorLevel = intLevel) And (cOp.DecisionPolicy = intPolicyBefore), _
            "a refused policy request leaves mode, policy, blocked flag and error level alone (policy " & intPolicy & ")"
    Next intPolicy
    TestAssert cOp.Status = eosRunning, "the running operation is not disturbed"
    TestAssert cRoot.IsValid, "the running operation keeps its lease"

    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "the first root still completes"

End Sub


Public Sub TestPolicyRequestRefusedWhileRunning()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strJson As String
    Dim intModeBefore As eInteractionMode
    Dim intPolicyBefore As eDecisionPolicy

    ' SetOperationPolicy asks PolicyRequestRefusal about the session operation. A private
    ' instance stands in for it, so the running state is set up here and never depends on
    ' where the host run started.
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "the root is granted"
    intModeBefore = cOp.InteractionMode
    intPolicyBefore = cOp.DecisionPolicy

    strJson = PolicyRequestRefusal(cOp)
    TestAssert InStr(strJson, """success"":false") > 0, "setting a policy while an operation runs fails"
    TestAssert InStr(strJson, "operation_already_running") > 0, "the refusal names operation_already_running"
    TestAssert (cOp.InteractionMode = intModeBefore) And (cOp.DecisionPolicy = intPolicyBefore), _
        "the refused request did not touch the running operation"

    cRoot.Complete eorSuccess

End Sub


Public Sub TestPolicyRequestAcceptedWhenReady()

    Dim cOp As clsOperation

    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    TestAssert cOp.Status = eosReady, "a new operation is ready"
    TestAssert Len(PolicyRequestRefusal(cOp)) = 0, "a ready operation is not refused"

End Sub


Public Sub TestClearOperationPolicyIsHarmlessWhileRunning()

    TestAssert InStr(VCS.ClearOperationPolicy, """success"":true") > 0, "clearing a policy that was never set succeeds"
    TestAssert InStr(VCS.ClearOperationPolicy, """success"":true") > 0, "clearing twice succeeds"
    TestAssert LiveStateUnchanged, "clearing an unset policy leaves the running operation alone"

End Sub


Public Sub TestSessionPolicyStaysInForceAfterFinish()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewUnattendedOperation
    TestAssert InStr(SetSessionPolicy(cOp, "decline"), """success"":true") > 0, "the policy is set"
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert cOp.DecisionPolicy = edpDecline, "the operation runs under the session policy"
    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "the operation finished"
    TestAssert cOp.InteractionMode = eimNonInteractive, "Finish leaves the session policy noninteractive"
    TestAssert cOp.DecisionPolicy = edpDecline, "Finish leaves the session policy in force"
    TestAssert Not cOp.ResolvePrompt(vbYesNo, "Confirm", "Proceed?").Blocked, "the policy still answers prompts"
    ClearSessionPolicy cOp

End Sub


Public Sub TestClearSessionPolicyClosesItAndASecondClearIsANoOp()

    Dim cOp As clsOperation
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy
    Dim lngOther As Long

    ' The caller clears a policy set for an operation that never started.
    Set cOp = NewUnattendedOperation
    intMode = cOp.InteractionMode
    intPolicy = cOp.DecisionPolicy
    SetSessionPolicy cOp, "block"
    TestAssert InStr(ClearSessionPolicy(cOp), """success"":true") > 0, "clearing succeeds"
    TestAssert (cOp.InteractionMode = intMode) And (cOp.DecisionPolicy = intPolicy), _
        "clearing restores the mode and policy from before the set"

    ' A second clear must not close a scope someone else opened since.
    lngOther = cOp.PushInteractionScope(eimNonInteractive, edpDecline, False)
    TestAssert InStr(ClearSessionPolicy(cOp), """success"":true") > 0, "a second clear succeeds"
    TestAssert cOp.DecisionPolicy = edpDecline, "a second clear changes nothing"
    cOp.CloseInteractionScope lngOther

End Sub


Public Sub TestSetSessionPolicyReplacesRatherThanNests()

    Dim cOp As clsOperation
    Dim intPolicy As eDecisionPolicy

    Set cOp = NewUnattendedOperation
    intPolicy = cOp.DecisionPolicy
    SetSessionPolicy cOp, "block"
    SetSessionPolicy cOp, "decline"
    TestAssert cOp.DecisionPolicy = edpDecline, "the second policy is in force"
    ClearSessionPolicy cOp
    TestAssert cOp.DecisionPolicy = intPolicy, "one clear removes it, with nothing left underneath"

End Sub


Public Sub TestEachOperationUnderASessionPolicyStartsUnblocked()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    ' A prompt one operation left unanswered belongs to that operation's result, not
    ' to every later operation that runs under the same session policy.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cOp.ResolvePrompt vbYesNo, "Confirm", "Proceed?"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorDecisionRequired, "the first operation stopped for a decision"

    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cOp.DecisionBlocked, "the next operation starts unblocked"
    TestAssert cOp.Decisions.Count = 0, "the next operation starts with no decisions"
    TestAssert cOp.DecisionPolicy = edpBlock, "the next operation runs under the session policy"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "the next operation can succeed"
    ClearSessionPolicy cOp

End Sub


Public Sub TestOperationPolicyInsideSessionPolicyEndsAtFinish()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    ' MergeBuild and RunFilteredTests pass their own policy to the root. It applies for
    ' that operation only, and the session policy is back in force afterwards.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "decline"
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    TestAssert cOp.DecisionPolicy = edpBlock, "the operation's own policy applies while it runs"
    cRoot.Complete eorSuccess
    TestAssert (cOp.InteractionMode = eimNonInteractive) And (cOp.DecisionPolicy = edpDecline), _
        "the session policy is back after Finish"
    ClearSessionPolicy cOp

End Sub


Public Sub TestRefusedSingleObjectImportFailsUnderSessionPolicy()

    Dim cOp As clsOperation
    Dim colErrors As Collection
    Dim dEntry As Dictionary
    Dim strJson As String

    ' ImportObject as the MCP calls it: a session policy, then a merge that clsDbForm
    ' refuses with Log.Error. The policy acknowledges the error's message box, and the
    ' result still has to report the failure after Finish restores the error level.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    TestAssert cOp.Begin(eotMerge), "the import begins under the session policy"
    TestAssert Not cOp.ResolvePrompt(vbExclamation, "Error", "Merging not supported").Blocked, _
        "the logged error's message box is acknowledged without a dialog"
    cOp.ErrorLevel = eelError
    Set dEntry = New Dictionary
    dEntry.Add "level", "error"
    dEntry.Add "message", "ERROR: Merging not supported for add-in forms. Use full build instead."
    Set colErrors = New Collection
    colErrors.Add dEntry

    strJson = FinishSingleObjectImport(cOp, "C:\src\logs\Merge_1.log", colErrors)
    TestAssert cOp.Status = eosReady, "the import's operation is finished"
    TestAssert InStr(strJson, """success"":false") > 0, "a refused merge is not a success"
    TestAssert InStr(strJson, "Merging not supported for add-in forms") > 0, "the logged error is in the result"
    TestAssert InStr(strJson, "Merge_1.log") > 0, "the result names the log"
    ClearSessionPolicy cOp

End Sub


Public Sub TestSingleObjectErrorFinishesItsOperation()

    Dim cOp As clsOperation
    Dim strJson As String

    ' An error raised inside ImportObject or ExportObject after Begin. The next call
    ' could not begin while the root stays open, and the session policy set around this
    ' one must be back in force, not the operation's scope.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    TestAssert cOp.Begin(eotMerge), "the import begins"

    strJson = FailSingleObjectOperation(cOp, "C:\src\logs\Merge_1.log", 91, "Object variable not set")
    TestAssert cOp.Status = eosReady, "the raised error finishes the operation"
    TestAssert cOp.Result = eorFailed, "the operation finishes as failed"
    TestAssert cOp.DecisionPolicy = edpBlock, "the session policy is back in force"
    TestAssert InStr(strJson, """success"":false") > 0, "the result is not a success"
    TestAssert InStr(strJson, "Object variable not set") > 0, "the result carries the error"
    TestAssert InStr(strJson, """errorNumber"":91") > 0, "the result carries the error number"
    TestAssert InStr(strJson, "Merge_1.log") > 0, "the result names the log"
    TestAssert cOp.Begin(eotMerge), "the next call can begin"
    cOp.Finish eorSuccess
    ClearSessionPolicy cOp

End Sub


Public Sub TestBlockedPromptInSingleObjectImportIsDecisionRequired()

    Dim cOp As clsOperation
    Dim strJson As String

    ' ImportObject as the MCP calls it, with a confirmation the block policy does not
    ' answer. The caller has to see the decision, not a plain success.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    TestAssert cOp.Begin(eotMerge), "the import begins under the session policy"
    TestAssert cOp.ResolvePrompt(vbYesNo, "Confirm", "Overwrite?").Blocked, "the policy blocks the prompt"

    strJson = FinishSingleObjectImport(cOp, "C:\src\logs\Merge_1.log", New Collection)
    TestAssert cOp.Status = eosReady, "the import's operation is finished"
    TestAssert InStr(strJson, """success"":false") > 0, "a blocked prompt is not a success"
    TestAssert InStr(strJson, """error_pattern"":""decision_required""") > 0, "the pattern is decision_required"
    TestAssert InStr(strJson, """decision_required"":true") > 0, "the result flags the decision"
    TestAssert InStr(strJson, "Overwrite?") > 0, "the decisions list the blocked prompt"
    TestAssert InStr(strJson, "Merge_1.log") > 0, "the result names the log"
    ClearSessionPolicy cOp

End Sub


Public Sub TestBlockedPromptInSingleObjectExportIsDecisionRequired()

    Dim cOp As clsOperation
    Dim strJson As String

    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    TestAssert cOp.Begin(eotExport), "the export begins under the session policy"
    TestAssert cOp.ResolvePrompt(vbYesNo, "Confirm", "Overwrite?").Blocked, "the policy blocks the prompt"

    strJson = FinishSingleObjectExport(cOp, "C:\src\logs\Export_1.log", New Dictionary)
    TestAssert cOp.Status = eosReady, "the export's operation is finished"
    TestAssert InStr(strJson, """success"":false") > 0, "a blocked prompt is not a success"
    TestAssert InStr(strJson, """error_pattern"":""decision_required""") > 0, "the pattern is decision_required"
    TestAssert InStr(strJson, "Overwrite?") > 0, "the decisions list the blocked prompt"
    TestAssert InStr(strJson, "Export_1.log") > 0, "the result names the log"
    ClearSessionPolicy cOp

End Sub


Public Sub TestSingleObjectExportWithoutPromptSucceeds()

    Dim cOp As clsOperation
    Dim strJson As String

    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    TestAssert cOp.Begin(eotExport), "the export begins under the session policy"

    strJson = FinishSingleObjectExport(cOp, "C:\src\logs\Export_1.log", New Dictionary)
    TestAssert InStr(strJson, """success"":true") > 0, "an export with no prompt succeeds"
    TestAssert InStr(strJson, "decision_required") = 0, "no decision is reported"
    TestAssert InStr(strJson, "Export_1.log") > 0, "the result names the log"
    ClearSessionPolicy cOp

End Sub


Public Sub TestSingleObjectErrorAfterBlockedPromptIsDecisionRequired()

    Dim cOp As clsOperation
    Dim strJson As String

    ' A raised error after a blocked prompt: the decision is primary, the error rides along.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    TestAssert cOp.Begin(eotMerge), "the import begins"
    cOp.ResolvePrompt vbYesNo, "Confirm", "Overwrite?"

    strJson = FailSingleObjectOperation(cOp, "C:\src\logs\Merge_1.log", 91, "Object variable not set")
    TestAssert cOp.Status = eosReady, "the raised error finishes the operation"
    TestAssert InStr(strJson, """error_pattern"":""decision_required""") > 0, "the pattern is decision_required"
    TestAssert InStr(strJson, "Overwrite?") > 0, "the decisions list the blocked prompt"
    TestAssert InStr(strJson, """runtime_error"":""Object variable not set""") > 0, "the error is kept as runtime_error"
    TestAssert InStr(strJson, """errorNumber"":91") > 0, "the result carries the error number"
    ClearSessionPolicy cOp

End Sub


Public Sub TestSingleObjectErrorBeforeBeginLeavesRunningRootAlone()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strJson As String

    ' An error before the call's own Begin: the running root belongs to someone else.
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    strJson = FailSingleObjectOperation(Nothing, vbNullString, 5, "Invalid procedure call")
    TestAssert cOp.Status = eosRunning, "the other root is still running"
    TestAssert InStr(strJson, """success"":false") > 0, "the result is not a success"
    TestAssert InStr(strJson, "logPath") = 0, "no log is named when none was written"
    cRoot.Complete eorSuccess

End Sub


Public Sub TestSessionPolicySetAndClearedInSeparateCalls()

    Dim cOp As clsOperation
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy

    ' The MCP sets and clears the policy in two API calls, each on a new
    ' clsVersionControl, so nothing but the operation lives from one to the next.
    Set cOp = NewUnattendedOperation
    intMode = cOp.InteractionMode
    intPolicy = cOp.DecisionPolicy
    SetPolicyInOneCall cOp, "block"
    SetPolicyInOneCall cOp, "decline"
    TestAssert cOp.DecisionPolicy = edpDecline, "the second call's policy is in force"
    ClearPolicyInOneCall cOp
    TestAssert (cOp.InteractionMode = intMode) And (cOp.DecisionPolicy = intPolicy), _
        "a clear in a later call restores the mode and policy from before the first set"

End Sub


Public Sub TestSessionPolicyBelongsToItsOperation()

    Dim cFirst As clsOperation
    Dim cSecond As clsOperation
    Dim intMode As eInteractionMode

    ' Scope tokens are numbered per operation, so each operation must hold its own.
    Set cFirst = NewUnattendedOperation
    Set cSecond = NewUnattendedOperation
    intMode = cFirst.InteractionMode
    SetSessionPolicy cFirst, "block"
    SetSessionPolicy cSecond, "decline"
    ClearSessionPolicy cFirst
    TestAssert cFirst.InteractionMode = intMode, "clearing one operation's policy closes it"
    TestAssert (cSecond.InteractionMode = eimNonInteractive) And (cSecond.DecisionPolicy = edpDecline), _
        "the other operation's policy stays in force"
    ClearSessionPolicy cSecond

End Sub


Public Sub TestRuntimeErrorRidesOnTerminalCallback()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dDone As Dictionary

    ' An async caller reads only the callback, so the error text has to be on it.
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cRoot.RecordRuntimeError "Division by zero", 11
    cRoot.Complete eorFailed

    Set dDone = cOp.LastCompletion
    TestAssert dDone("type") = "error", "a runtime error finishes as an error callback"
    TestAssert dDone("message") = "Division by zero", "the message is the error text"
    TestAssert dDone("runtime_error") = "Division by zero", "the callback carries runtime_error"
    TestAssert dDone("errorNumber") = 11, "the callback carries the error number"
    TestAssert Not dDone.Exists("decision_required"), "no prompt was blocked"

End Sub


Public Sub TestRuntimeErrorAfterBlockedPromptRidesOnDecisionCallback()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dDone As Dictionary

    ' The decision is primary; the runtime error rides along, as in the sync result.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cOp.ResolvePrompt vbYesNo, "Confirm", "Overwrite?"
    cRoot.RecordRuntimeError "Object variable not set", 91
    cRoot.Complete eorFailed

    Set dDone = cOp.LastCompletion
    TestAssert dDone("decision_required") = True, "the callback is decision_required"
    TestAssert dDone("error_pattern") = ERR_DECISION_REQUIRED, "the pattern is decision_required"
    TestAssert dDone.Exists("decisions"), "the callback lists the blocked prompt"
    TestAssert dDone("runtime_error") = "Object variable not set", "the error is kept as runtime_error"
    TestAssert dDone("errorNumber") = 91, "the callback carries the error number"
    ClearSessionPolicy cOp

End Sub


Public Sub TestRuntimeErrorBelongsToItsRoot()

    Dim cOp As clsOperation
    Dim cFirst As clsRootOperationLease
    Dim cSecond As clsRootOperationLease

    ' Neither the next root nor a lease that no longer holds the root can carry it over.
    Set cOp = NewUnattendedOperation
    Set cFirst = cOp.TryBeginRoot(eotOther)
    cFirst.RecordRuntimeError "Division by zero", 11
    cFirst.Complete eorFailed

    Set cSecond = cOp.TryBeginRoot(eotOther)
    cFirst.RecordRuntimeError "Stale lease", 5
    cSecond.Complete eorSuccess
    TestAssert cOp.LastCompletion("type") = "complete", "the next root succeeds"
    TestAssert Not cOp.LastCompletion.Exists("runtime_error"), "no runtime_error on the next root"
    TestAssert Not cOp.LastCompletion.Exists("errorNumber"), "no error number on the next root"

End Sub


' One SetOperationPolicy or ClearOperationPolicy request: no state survives the call
' except what the operation itself keeps.
Private Sub SetPolicyInOneCall(ByVal cOp As clsOperation, ByVal strPolicy As String)
    TestAssert InStr(SetSessionPolicy(cOp, strPolicy), """success"":true") > 0, "the " & strPolicy & " policy is set"
End Sub

Private Sub ClearPolicyInOneCall(ByVal cOp As clsOperation)
    TestAssert InStr(ClearSessionPolicy(cOp), """success"":true") > 0, "clearing succeeds"
End Sub

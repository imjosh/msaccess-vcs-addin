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
'           : operation is only read, by the clearing test and the invalid test-run
'           : policy test. Start, restore and exit paths of MergeBuild and
'           : RunFilteredTests need a free operation and are covered by scripted passes
'           : (X01; A18 for a test run's options loading).
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


Public Sub TestRefusedNestedBeginPreservesSynchronousRoot()

    Dim cOp As clsOperation
    Dim strToken As String
    Dim intMode As eInteractionMode
    Dim intModeBefore As eInteractionMode
    Dim dDone As Dictionary

    ' Begin keeps the lease of its synchronous root on the instance. A refused Begin
    ' must not replace it: releasing that lease abandons the root while its owner is
    ' still executing.
    Set cOp = NewUnattendedOperation
    intModeBefore = cOp.InteractionMode
    TestAssert cOp.Begin(eotExport, edpBlock), "the first Begin is granted"
    strToken = cOp.CurrentRootToken
    intMode = cOp.InteractionMode
    TestAssert Len(strToken) > 0, "the root has a token"
    TestAssert cOp.DecisionPolicy = edpBlock, "the root runs under the block policy"

    TestAssert Not cOp.Begin(eotMerge, edpBlock), "a nested Begin is refused"
    TestAssert cOp.Status = eosRunning, "the refusal leaves the root running"
    TestAssert cOp.CurrentRootToken = strToken, "the root keeps its identity"
    TestAssert cOp.OperationType = eotExport, "the root keeps its type"
    TestAssert cOp.DecisionPolicy = edpBlock, "the root keeps its policy"
    TestAssert cOp.InteractionMode = intMode, "the root keeps its interaction scope"
    TestAssert cOp.LastCompletion Is Nothing, "the refusal emits no terminal callback"

    ' A second refusal is as harmless as the first.
    TestAssert Not cOp.Begin(eotOther), "an unscoped nested Begin is refused"
    TestAssert cOp.Status = eosRunning, "the root is still running"
    TestAssert cOp.LastCompletion Is Nothing, "still no terminal callback"

    cOp.Finish eorSuccess
    Set dDone = cOp.LastCompletion
    TestAssert Not dDone Is Nothing, "the owner's Finish completes the root"
    TestAssert cOp.Status = eosReady, "the operation is ready again"
    TestAssert cOp.Result = eorSuccess, "the root keeps the owner's result"
    TestAssert cOp.InteractionMode = intModeBefore, "the interaction scope is restored"
    TestAssert cOp.Begin(eotOther), "a later Begin succeeds"
    cOp.Finish eorSuccess
    TestAssert Not cOp.LastCompletion Is dDone, "the later root completes on its own"

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


Public Sub TestInvalidSessionPolicyLeavesThePolicyInForce()

    Dim cOp As clsOperation
    Dim intPolicy As eDecisionPolicy

    Set cOp = NewUnattendedOperation
    intPolicy = cOp.DecisionPolicy
    SetSessionPolicy cOp, "block"
    TestAssert InStr(SetSessionPolicy(cOp, "no-such-policy"), ERR_INVALID_DECISION_POLICY) > 0, _
        "an unknown name is refused"
    TestAssert cOp.DecisionPolicy = edpBlock, "the refusal leaves the earlier policy in force"
    TestAssert cOp.InteractionMode = eimNonInteractive, "the refusal leaves the operation noninteractive"
    ClearSessionPolicy cOp
    TestAssert cOp.DecisionPolicy = intPolicy, "one clear still removes it"

End Sub


Public Sub TestInvalidTestRunPolicyRefusedBeforeLoadingOptions()

    Dim cSaved As clsOptions
    Dim intStatus As eOperationState
    Dim blnLoaded As Boolean
    Dim strJson As String

    ' Loading the options can log an error that opens a message box, so an unknown
    ' policy is refused before RunFilteredTests reads its default filter. Nothing
    ' starts, so the session operation is only read.
    If OptionsLoaded Then Set cSaved = Options
    intStatus = Operation.Status
    Set Options = Nothing
    strJson = VCS.RunFilteredTests("no-such-policy")
    blnLoaded = OptionsLoaded
    Set Options = cSaved

    TestAssert InStr(strJson, ERR_INVALID_DECISION_POLICY) > 0, "an unknown name is refused"
    TestAssert Not blnLoaded, "the refusal does not load the options"
    TestAssert Operation.Status = intStatus, "the refusal leaves the operation's status alone"
    TestAssert LiveStateUnchanged(), "the refusal leaves the interaction state alone"

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


Public Sub TestApiErrorAfterOwnFinishDoesNotFinishAgain()

    Dim cOp As clsOperation
    Dim dDone As Dictionary
    Dim strJson As String

    ' An error raised in ImportByType or ExportByType after FinishScopedSync already
    ' finished the call's own root. The root completes once, with its own result.
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    TestAssert cOp.Begin(eotExport), "the export begins"
    cOp.Finish eorSuccess
    Set dDone = cOp.LastCompletion

    strJson = FailSingleObjectOperation(cOp, "C:\src\logs\Export_1.log", 6, "Overflow")
    TestAssert cOp.LastCompletion Is dDone, "no second completion is emitted"
    TestAssert cOp.Result = eorSuccess, "the finished root keeps its result"
    TestAssert cOp.DecisionPolicy = edpBlock, "the session policy is still in force"
    TestAssert InStr(strJson, """success"":false") > 0, "the result is not a success"
    TestAssert InStr(strJson, """error"":""Overflow""") > 0, "the result carries the error"
    TestAssert InStr(strJson, """errorNumber"":6") > 0, "the result carries the error number"
    ClearSessionPolicy cOp

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


Public Sub TestInteractionModeSelectionReportsEffectiveMode()

    Dim lngErr As Long
    Dim strErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary
    Dim intMode As Long

    On Error GoTo ErrHandler
    Set cOp = NewUnattendedOperation
    For intMode = eimNonInteractive To eimNormal Step -1
        Set dResult = ParseJson(SelectInteractionMode(cOp, intMode))
        TestAssert dResult("success") = True, "unscoped mode selection succeeds"
        TestAssert dResult("requested_mode") = intMode, "requested mode is reported"
        TestAssert dResult("effective_mode") = intMode, "effective mode matches the request"
        TestAssert cOp.InteractionMode = intMode, "the mode actually took effect"
        TestAssert Not dResult.Exists("error_pattern"), "acceptance has no refusal pattern"
        TestAssert Not dResult.Exists("started"), "selecting a mode starts no operation"
    Next intMode

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "TestInteractionModeSelectionReportsEffectiveMode " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestInteractiveSelectionPreservesCallerScope()

    Dim lngErr As Long
    Dim strErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary
    Dim lngToken As Long

    On Error GoTo ErrHandler
    Set cOp = NewUnattendedOperation
    lngToken = cOp.PushInteractionScope(eimNonInteractive, edpDecline)
    cOp.RecordDecision NewDecision(DECISION_DECLINED, "Confirm", "Overwrite?", DECISION_DECLINED)
    cOp.DecisionBlocked = True
    cOp.ErrorLevel = eelCritical
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNormal))
    TestAssert dResult("success") = False, "interactive selection is refused"
    TestAssert dResult("error_pattern") = "interaction_mode_refused", "stable refusal pattern"
    TestAssert Len(dResult("error")) > 0, "refusal explains how to recover"
    TestAssert dResult("requested_mode") = 0, "interactive mode was requested"
    TestAssert dResult("effective_mode") = 2, "noninteractive mode is still effective"
    TestAssert cOp.InteractionMode = eimNonInteractive, "the enclosing mode stays in force"
    TestAssert cOp.DecisionPolicy = edpDecline, "the enclosing policy stays in force"
    TestAssert cOp.DecisionBlocked, "the blocked flag is preserved"
    TestAssert cOp.ErrorLevel = eelCritical, "the error level is preserved"
    TestAssert cOp.Decisions.Count = 1, "the decision journal is preserved"
    cOp.CloseInteractionScope lngToken
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNormal))
    TestAssert dResult("success") = True, "the owner's close allows interactive selection"
    TestAssert cOp.InteractionMode = eimNormal, "interactive mode is effective after the close"

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "TestInteractiveSelectionPreservesCallerScope " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestInteractiveSelectionAfterFailedPolicyCleanup()

    Dim lngErr As Long
    Dim strErr As String
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dResult As Dictionary

    On Error GoTo ErrHandler
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "block"
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cRoot.RecordRuntimeError "Failed operation", 5
    cRoot.Complete eorFailed
    ' Simulate a caller whose cleanup never reached ClearOperationPolicy.
    ' Finishing the failed root releases only operation-owned scopes.
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNormal))
    TestAssert cOp.Status = eosReady, "the failed operation has finished"
    TestAssert dResult("success") = False, "leftover session policy refuses interactive mode"
    TestAssert dResult("error_pattern") = "interaction_mode_refused", "cleanup failure is observable"
    TestAssert dResult("effective_mode") = 2, "leftover noninteractive mode is reported"
    TestAssert cOp.DecisionPolicy = edpBlock, "the caller's policy remains intact"
    TestAssert cOp.Result = eorFailed, "selecting a mode preserves the failed outcome"
    ClearPolicyInOneCall cOp
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNormal))
    TestAssert dResult("success") = True, "a later owner cleanup allows selection"
    TestAssert dResult("effective_mode") = 0, "interactive mode is effective after cleanup"

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If lngErr <> 0 Then TestAssert False, "TestInteractiveSelectionAfterFailedPolicyCleanup " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestSilentSelectionInsidePolicyScopeIsRefused()

    Dim lngErr As Long
    Dim strErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary

    On Error GoTo ErrHandler
    Set cOp = NewUnattendedOperation
    SetSessionPolicy cOp, "decline"
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimSilent))
    TestAssert dResult("success") = False, "silent mode cannot weaken a noninteractive scope"
    TestAssert dResult("requested_mode") = 1, "silent request is reported"
    TestAssert dResult("effective_mode") = 2, "effective mode remains noninteractive"
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNonInteractive))
    TestAssert dResult("success") = True, "an already effective mode is accepted"
    TestAssert cOp.DecisionPolicy = edpDecline, "selecting the same mode preserves policy"
    ClearPolicyInOneCall cOp

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "TestSilentSelectionInsidePolicyScopeIsRefused " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestInteractiveSelectionCannotRelaxRunningRoot()

    Dim lngErr As Long
    Dim strErr As String
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dResult As Dictionary

    On Error GoTo ErrHandler
    Set cOp = NewUnattendedOperation
    cOp.InteractionMode = eimSilent
    Set cRoot = cOp.TryBeginRoot(eotOther)
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNormal))
    TestAssert dResult("success") = False, "a running silent root cannot be relaxed"
    TestAssert dResult("error_pattern") = "interaction_mode_refused", "active-root refusal is structured"
    TestAssert dResult("effective_mode") = 1, "silent mode is still effective"
    TestAssert cRoot.IsValid, "the root lease is left intact"
    cRoot.Complete eorSuccess
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNormal))
    TestAssert dResult("success") = True, "interactive selection succeeds after the root finishes"

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If lngErr <> 0 Then TestAssert False, "TestInteractiveSelectionCannotRelaxRunningRoot " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestInvalidInteractionModeLeavesStateIntact()

    Dim lngErr As Long
    Dim strErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary
    Dim varMode As Variant

    On Error GoTo ErrHandler
    Set cOp = NewUnattendedOperation
    cOp.InteractionMode = eimSilent
    For Each varMode In Array(-1, 3, 32768)
        Set dResult = ParseJson(SelectInteractionMode(cOp, CLng(varMode)))
        TestAssert dResult("success") = False, "unknown mode is refused"
        TestAssert dResult("error_pattern") = "invalid_interaction_mode", "invalid-mode pattern is stable"
        TestAssert dResult("requested_mode") = varMode, "invalid request is reported"
        TestAssert dResult("effective_mode") = 1, "the previous mode is reported"
        TestAssert cOp.InteractionMode = eimSilent, "invalid input does not mutate the mode"
    Next varMode

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "TestInvalidInteractionModeLeavesStateIntact " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


' One SetOperationPolicy or ClearOperationPolicy request: no state survives the call
' except what the operation itself keeps.
Private Sub SetPolicyInOneCall(ByVal cOp As clsOperation, ByVal strPolicy As String)
    TestAssert InStr(SetSessionPolicy(cOp, strPolicy), """success"":true") > 0, "the " & strPolicy & " policy is set"
End Sub

Private Sub ClearPolicyInOneCall(ByVal cOp As clsOperation)
    TestAssert InStr(ClearSessionPolicy(cOp), """success"":true") > 0, "clearing succeeds"
End Sub

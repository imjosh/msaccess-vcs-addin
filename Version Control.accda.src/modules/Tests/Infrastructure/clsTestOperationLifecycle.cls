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
'           : interaction state. Explicit policies belong to individual requests and
'           : completion releases them without requiring a client cleanup call.
'           : Every test that needs a running operation uses a private
'           : instance, so no result depends on where the host run started. The session
'           : operation is only read, by the clearing test and the invalid test-run
'           : policy test. Start, restore and exit paths of MergeBuild and
'           : RunFilteredTests need a free operation and are covered by scripted passes
'           : (for a test run's options loading).
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
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
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
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
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

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestBeginWithPolicyRefusedWhileRunning cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestBeginWithPolicyRefusedWhileRunning " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRefusedNestedBeginPreservesSynchronousRoot()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
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
    TestAssert Not cOp.Begin(eotOther, edpBlock), "an unscoped nested Begin is refused"
    TestAssert cOp.Status = eosRunning, "the root is still running"
    TestAssert cOp.LastCompletion Is Nothing, "still no terminal callback"

    cOp.Finish eorSuccess
    Set dDone = cOp.LastCompletion
    TestAssert Not dDone Is Nothing, "the owner's Finish completes the root"
    TestAssert cOp.Status = eosReady, "the operation is ready again"
    TestAssert cOp.Result = eorSuccess, "the root keeps the owner's result"
    TestAssert cOp.InteractionMode = intModeBefore, "the interaction scope is restored"
    TestAssert cOp.Begin(eotOther, edpBlock), "a later Begin succeeds"
    cOp.Finish eorSuccess
    TestAssert Not cOp.LastCompletion Is dDone, "the later root completes on its own"

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRefusedNestedBeginPreservesSynchronousRoot cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRefusedNestedBeginPreservesSynchronousRoot " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestPolicyRequestRefusedWhileRunning()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strJson As String
    Dim intModeBefore As eInteractionMode
    Dim intPolicyBefore As eDecisionPolicy

    ' PolicyRequestRefusal is asked about the session operation. A private
    ' instance stands in for it, so the running state is set up here and never depends on
    ' where the host run started.
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    TestAssert Not cRoot Is Nothing, "the root is granted"
    intModeBefore = cOp.InteractionMode
    intPolicyBefore = cOp.DecisionPolicy

    strJson = PolicyRequestRefusal(cOp)
    TestAssert InStr(strJson, """success"":false") > 0, "setting a policy while an operation runs fails"
    TestAssert InStr(strJson, "operation_already_running") > 0, "the refusal names operation_already_running"
    TestAssert (cOp.InteractionMode = intModeBefore) And (cOp.DecisionPolicy = intPolicyBefore), _
        "the refused request did not touch the running operation"

    cRoot.Complete eorSuccess

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestPolicyRequestRefusedWhileRunning cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestPolicyRequestRefusedWhileRunning " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestPolicyRequestAcceptedWhenReady()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation

    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    TestAssert cOp.Status = eosReady, "a new operation is ready"
    TestAssert Len(PolicyRequestRefusal(cOp)) = 0, "a ready operation is not refused"

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestPolicyRequestAcceptedWhenReady cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestPolicyRequestAcceptedWhenReady " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestClearOperationPolicyIsHarmlessWhileRunning()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    TestAssert InStr(VCS.ClearOperationPolicy, """success"":true") > 0, "clearing a policy that was never set succeeds"
    TestAssert InStr(VCS.ClearOperationPolicy, """success"":true") > 0, "clearing twice succeeds"
    TestAssert LiveStateUnchanged, "clearing an unset policy leaves the running operation alone"

CleanUp:
    On Error Resume Next
    If lngTestErr <> 0 Then TestAssert False, "TestClearOperationPolicyIsHarmlessWhileRunning " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRequestPolicyEndsAtFinish()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpDecline)
    TestAssert cOp.DecisionPolicy = edpDecline, "this request owns its policy"
    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "the request completed"
    TestAssert cOp.InteractionMode = eimNormal, "completion releases noninteractive mode"
    TestAssert cOp.DecisionPolicy = edpAsk, "completion releases the request policy"
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert cOp.DecisionPolicy = edpAsk, "a later request has no remembered policy"
    cRoot.Complete eorSuccess

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRequestPolicyEndsAtFinish cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRequestPolicyEndsAtFinish " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestClearSessionPolicyClosesItAndASecondClearIsANoOp()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy
    Dim lngOther As Long

    ' The caller clears a policy set for an operation that never started.
    Set cOp = NewUnattendedOperation
    intMode = cOp.InteractionMode
    intPolicy = cOp.DecisionPolicy
    TestAssert InStr(ClearSessionPolicy(cOp), """success"":true") > 0, "clearing succeeds"
    TestAssert (cOp.InteractionMode = intMode) And (cOp.DecisionPolicy = intPolicy), _
        "clearing restores the mode and policy from before the set"

    ' A second clear must not close a scope someone else opened since.
    lngOther = cOp.PushInteractionScope(eimNonInteractive, edpDecline, False)
    TestAssert InStr(ClearSessionPolicy(cOp), """success"":true") > 0, "a second clear succeeds"
    TestAssert cOp.DecisionPolicy = edpDecline, "a second clear changes nothing"
    cOp.CloseInteractionScope lngOther

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestClearSessionPolicyClosesItAndASecondClearIsANoOp cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestClearSessionPolicyClosesItAndASecondClearIsANoOp " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestPersistentPolicyActivationIsRetired()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Set cOp = NewUnattendedOperation
    TestAssert InStr(SetSessionPolicy(cOp, "block"), "policy_activation_retired") > 0, "activation is explicitly retired"
    TestAssert InStr(SetSessionPolicy(cOp, "decline"), "policy_activation_retired") > 0, "a second activation is also refused"
    TestAssert cOp.InteractionMode = eimNormal, "retired activation changes no mode"
    TestAssert cOp.DecisionPolicy = edpAsk, "retired activation remembers no policy"

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestPersistentPolicyActivationIsRetired cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestPersistentPolicyActivationIsRetired " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRetiredActivationPreservesActivePolicy()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    TestAssert InStr(SetSessionPolicy(cOp, "no-such-policy"), "policy_activation_retired") > 0, "all persistent activation is retired"
    TestAssert cOp.DecisionPolicy = edpBlock, "the active request keeps its policy"
    TestAssert cOp.Decisions.Count = 0, "a refusal does not append to its journal"
    cRoot.Complete eorSuccess

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRetiredActivationPreservesActivePolicy cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRetiredActivationPreservesActivePolicy " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestInvalidTestRunPolicyRefusedBeforeLoadingOptions()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
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

CleanUp:
    On Error Resume Next
    If lngTestErr <> 0 Then TestAssert False, "TestInvalidTestRunPolicyRefusedBeforeLoadingOptions " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestEachExplicitRequestStartsUnblocked()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    ' A prompt one operation left unanswered belongs to that operation's result, not
    ' to every later operation that runs under the same session policy.
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    cOp.ResolvePrompt vbYesNo, "Confirm", "Proceed?"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorDecisionRequired, "the first operation stopped for a decision"

    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    TestAssert Not cOp.DecisionBlocked, "the next operation starts unblocked"
    TestAssert cOp.Decisions.Count = 0, "the next operation starts with no decisions"
    TestAssert cOp.DecisionPolicy = edpBlock, "the next operation runs under the session policy"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "the next operation can succeed"
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestEachExplicitRequestStartsUnblocked cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestEachExplicitRequestStartsUnblocked " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRequestPolicyHasNoSessionFallback()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    TestAssert cOp.DecisionPolicy = edpBlock, "the request policy applies"
    cRoot.Complete eorSuccess
    TestAssert cOp.InteractionMode = eimNormal And cOp.DecisionPolicy = edpAsk, "completion releases owned policy state"

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRequestPolicyHasNoSessionFallback cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRequestPolicyHasNoSessionFallback " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRefusedSingleObjectImportFailsUnderRequestPolicy()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim colErrors As Collection
    Dim dEntry As Dictionary
    Dim strJson As String

    ' ImportObject as the MCP calls it: a session policy, then a merge that clsDbForm
    ' refuses with Log.Error. The policy acknowledges the error's message box, and the
    ' result still has to report the failure after Finish restores the error level.
    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotMerge, edpBlock), "the import begins under the session policy"
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

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRefusedSingleObjectImportFailsUnderRequestPolicy cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRefusedSingleObjectImportFailsUnderRequestPolicy " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestSingleObjectErrorFinishesItsOperation()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim strJson As String

    ' An error raised inside ImportObject or ExportObject after Begin. The next call
    ' could not begin while the root stays open, and the session policy set around this
    ' one must be back in force, not the operation's scope.
    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotMerge, edpBlock), "the import begins"

    strJson = FailSingleObjectOperation(cOp, "C:\src\logs\Merge_1.log", 91, "Object variable not set")
    TestAssert cOp.Status = eosReady, "the raised error finishes the operation"
    TestAssert cOp.Result = eorFailed, "the operation finishes as failed"
    TestAssert cOp.DecisionPolicy = edpAsk, "the completed request releases its policy"
    TestAssert InStr(strJson, """success"":false") > 0, "the result is not a success"
    TestAssert InStr(strJson, "Object variable not set") > 0, "the result carries the error"
    TestAssert InStr(strJson, """errorNumber"":91") > 0, "the result carries the error number"
    TestAssert InStr(strJson, "Merge_1.log") > 0, "the result names the log"
    TestAssert cOp.Begin(eotMerge, edpBlock), "the next call can begin"
    cOp.Finish eorSuccess
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestSingleObjectErrorFinishesItsOperation cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestSingleObjectErrorFinishesItsOperation " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestBlockedPromptInSingleObjectImportIsDecisionRequired()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim strJson As String

    ' ImportObject as the MCP calls it, with a confirmation the block policy does not
    ' answer. The caller has to see the decision, not a plain success.
    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotMerge, edpBlock), "the import begins under the session policy"
    TestAssert cOp.ResolvePrompt(vbYesNo, "Confirm", "Overwrite?").Blocked, "the policy blocks the prompt"

    strJson = FinishSingleObjectImport(cOp, "C:\src\logs\Merge_1.log", New Collection)
    TestAssert cOp.Status = eosReady, "the import's operation is finished"
    TestAssert InStr(strJson, """success"":false") > 0, "a blocked prompt is not a success"
    TestAssert InStr(strJson, """error_pattern"":""decision_required""") > 0, "the pattern is decision_required"
    TestAssert InStr(strJson, """decision_required"":true") > 0, "the result flags the decision"
    TestAssert InStr(strJson, "Overwrite?") > 0, "the decisions list the blocked prompt"
    TestAssert InStr(strJson, "Merge_1.log") > 0, "the result names the log"
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestBlockedPromptInSingleObjectImportIsDecisionRequired cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestBlockedPromptInSingleObjectImportIsDecisionRequired " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestBlockedPromptInSingleObjectExportIsDecisionRequired()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim strJson As String

    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotExport, edpBlock), "the export begins under the session policy"
    TestAssert cOp.ResolvePrompt(vbYesNo, "Confirm", "Overwrite?").Blocked, "the policy blocks the prompt"

    strJson = FinishSingleObjectExport(cOp, "C:\src\logs\Export_1.log", New Dictionary)
    TestAssert cOp.Status = eosReady, "the export's operation is finished"
    TestAssert InStr(strJson, """success"":false") > 0, "a blocked prompt is not a success"
    TestAssert InStr(strJson, """error_pattern"":""decision_required""") > 0, "the pattern is decision_required"
    TestAssert InStr(strJson, "Overwrite?") > 0, "the decisions list the blocked prompt"
    TestAssert InStr(strJson, "Export_1.log") > 0, "the result names the log"
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestBlockedPromptInSingleObjectExportIsDecisionRequired cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestBlockedPromptInSingleObjectExportIsDecisionRequired " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestSingleObjectExportWithoutPromptSucceeds()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim strJson As String

    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotExport, edpBlock), "the export begins under the session policy"

    strJson = FinishSingleObjectExport(cOp, "C:\src\logs\Export_1.log", New Dictionary)
    TestAssert InStr(strJson, """success"":true") > 0, "an export with no prompt succeeds"
    TestAssert InStr(strJson, "decision_required") = 0, "no decision is reported"
    TestAssert InStr(strJson, "Export_1.log") > 0, "the result names the log"
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestSingleObjectExportWithoutPromptSucceeds cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestSingleObjectExportWithoutPromptSucceeds " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestSingleObjectErrorAfterBlockedPromptIsDecisionRequired()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim strJson As String

    ' A raised error after a blocked prompt: the decision is primary, the error rides along.
    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotMerge, edpBlock), "the import begins"
    cOp.ResolvePrompt vbYesNo, "Confirm", "Overwrite?"

    strJson = FailSingleObjectOperation(cOp, "C:\src\logs\Merge_1.log", 91, "Object variable not set")
    TestAssert cOp.Status = eosReady, "the raised error finishes the operation"
    TestAssert InStr(strJson, """error_pattern"":""decision_required""") > 0, "the pattern is decision_required"
    TestAssert InStr(strJson, "Overwrite?") > 0, "the decisions list the blocked prompt"
    TestAssert InStr(strJson, """runtime_error"":""Object variable not set""") > 0, "the error is kept as runtime_error"
    TestAssert InStr(strJson, """errorNumber"":91") > 0, "the result carries the error number"
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestSingleObjectErrorAfterBlockedPromptIsDecisionRequired cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestSingleObjectErrorAfterBlockedPromptIsDecisionRequired " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


' The prompts a policy answers without stopping the operation: an OK-only box
' acknowledged under block, a conflict applied by prefer_source, and a confirmation
' declined by decline.
Private Function AnsweredPromptScenarios() As Variant
    AnsweredPromptScenarios = Array("acknowledged", "applied", "declined")
End Function


' A private operation begun under the policy that answers the scenario's prompt, with
' that answer in its journal. strResolution returns the resolution the journal holds.
Private Function BeginWithAnsweredPrompt(ByVal strScenario As String, _
    ByVal intType As eOperationType, ByRef strResolution As String) As clsOperation

    Dim cOp As clsOperation
    Dim strPolicy As String

    Select Case strScenario
        Case "acknowledged": strPolicy = "block"
        Case "applied":      strPolicy = "prefer_source"
        Case Else:           strPolicy = "decline"
    End Select
    Set cOp = NewUnattendedOperation
    Dim intPolicy As eDecisionPolicy
    ParseDecisionPolicy strPolicy, intPolicy
    TestAssert cOp.Begin(intType, intPolicy), strScenario & ": the operation begins under " & strPolicy

    Select Case strScenario
        Case "acknowledged"
            cOp.ResolvePrompt vbOKOnly, "Notice", "Finished."
            strResolution = DECISION_ACKNOWLEDGED
        Case "applied"
            cOp.RecordDecision NewDecision(DECISION_APPLIED, "Merge conflicts", "2 conflict(s)", _
                ConflictResolutionName(edpPreferSource), "policy=" & CStr(edpPreferSource))
            strResolution = ConflictResolutionName(edpPreferSource)
        Case Else
            cOp.ResolvePrompt vbYesNo, "Confirm", "Overwrite?"
            strResolution = DECISION_DECLINED
    End Select
    Set BeginWithAnsweredPrompt = cOp

End Function


' The result reports the one answered prompt, and still reports the outcome itself.
Private Sub AssertJournalCarried(ByVal strJson As String, ByVal strResolution As String, _
    ByVal strLabel As String)

    Dim dResult As Dictionary

    Set dResult = ParseJson(strJson)
    TestAssert dResult.Exists("decisions"), strLabel & ": the result carries decisions"
    If dResult.Exists("decisions") Then
        TestAssert dResult("decisions").Count = 1, strLabel & ": one prompt was answered"
        TestAssert dResult("decisions")(1)("resolution") = strResolution, _
            strLabel & ": the journal holds the answer (" & strResolution & ")"
    End If
    TestAssert Not dResult.Exists("decision_required"), strLabel & ": an answered prompt is not decision_required"

End Sub


Public Sub TestAnsweredPromptsAreInSingleObjectImportResult()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim colErrors As Collection
    Dim dEntry As Dictionary
    Dim dResult As Dictionary
    Dim varScenario As Variant
    Dim strResolution As String
    Dim strJson As String

    Set dEntry = New Dictionary
    dEntry.Add "level", "error"
    dEntry.Add "message", "ERROR: Merging not supported for add-in forms."
    Set colErrors = New Collection
    colErrors.Add dEntry

    For Each varScenario In AnsweredPromptScenarios
        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotMerge, strResolution)
        strJson = FinishSingleObjectImport(cOp, "C:\src\logs\Merge_1.log", New Collection)
        AssertJournalCarried strJson, strResolution, "import success, " & varScenario
        TestAssert ParseJson(strJson)("success") = True, varScenario & ": the import still succeeds"
        TestAssert InStr(strJson, "Merge_1.log") > 0, varScenario & ": the result names the log"
        ClearSessionPolicy cOp

        ' A plain failure after the answered prompt keeps the journal too.
        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotMerge, strResolution)
        cOp.ErrorLevel = eelError
        strJson = FinishSingleObjectImport(cOp, "C:\src\logs\Merge_1.log", colErrors)
        AssertJournalCarried strJson, strResolution, "import failure, " & varScenario
        Set dResult = ParseJson(strJson)
        TestAssert dResult("success") = False, varScenario & ": the failure is still a failure"
        TestAssert InStr(dResult("error"), "Merging not supported") > 0, varScenario & ": the error is kept"
        ClearSessionPolicy cOp

        ' So does a cancel.
    Next varScenario

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestAnsweredPromptsAreInSingleObjectImportResult cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestAnsweredPromptsAreInSingleObjectImportResult " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestAnsweredPromptsAreInSingleObjectExportResult()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim dFailed As Dictionary
    Dim varScenario As Variant
    Dim strResolution As String
    Dim strJson As String

    For Each varScenario In AnsweredPromptScenarios
        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotExport, strResolution)
        strJson = FinishSingleObjectExport(cOp, "C:\src\logs\Export_1.log", New Dictionary)
        AssertJournalCarried strJson, strResolution, "export success, " & varScenario
        TestAssert ParseJson(strJson)("success") = True, varScenario & ": the export still succeeds"
        TestAssert InStr(strJson, "Export_1.log") > 0, varScenario & ": the result names the log"
        ClearSessionPolicy cOp

        ' An export that stopped early arrives with its own failure.
        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotExport, strResolution)
        cOp.ErrorLevel = eelCritical
        Set dFailed = New Dictionary
        dFailed.Add "success", False
        dFailed.Add "error", "Export fixture failure"
        dFailed.Add "cancelled", True
        strJson = FinishSingleObjectExport(cOp, "C:\src\logs\Export_1.log", dFailed)
        AssertJournalCarried strJson, strResolution, "export cancel, " & varScenario
        Set dFailed = ParseJson(strJson)
        TestAssert dFailed("success") = False, varScenario & ": the stopped export is still a failure"
        TestAssert dFailed("cancelled") = True, varScenario & ": the cancel is kept"
        ClearSessionPolicy cOp
    Next varScenario

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestAnsweredPromptsAreInSingleObjectExportResult cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestAnsweredPromptsAreInSingleObjectExportResult " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestAnsweredPromptsAreInScopedSyncResult()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary
    Dim varScenario As Variant
    Dim strResolution As String
    Dim strJson As String

    For Each varScenario In AnsweredPromptScenarios
        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotMerge, strResolution)
        Set dResult = New Dictionary
        dResult.Add "success", True
        dResult.Add "logPath", "C:\src\logs\Merge_1.log"
        strJson = FinishScopedOperation(cOp, dResult, eorSuccess, "C:\src\logs\Merge_1.log")
        AssertJournalCarried strJson, strResolution, "scoped success, " & varScenario
        TestAssert ParseJson(strJson)("success") = True, varScenario & ": the sync still succeeds"
        TestAssert InStr(strJson, "Merge_1.log") > 0, varScenario & ": the result names the log"
        ClearSessionPolicy cOp

        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotExport, strResolution)
        Set dResult = New Dictionary
        dResult.Add "success", False
        dResult.Add "error", "Operation completed with errors. Check the log for details."
        strJson = FinishScopedOperation(cOp, dResult, eorFailed, "C:\src\logs\Export_1.log")
        AssertJournalCarried strJson, strResolution, "scoped failure, " & varScenario
        TestAssert ParseJson(strJson)("success") = False, varScenario & ": the failure is still a failure"
        ClearSessionPolicy cOp

        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotExport, strResolution)
        Set dResult = New Dictionary
        dResult.Add "success", False
        dResult.Add "error", "Operation was canceled."
        strJson = FinishScopedOperation(cOp, dResult, eorCanceled, "C:\src\logs\Export_1.log")
        AssertJournalCarried strJson, strResolution, "scoped cancel, " & varScenario
        ClearSessionPolicy cOp
    Next varScenario

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestAnsweredPromptsAreInScopedSyncResult cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestAnsweredPromptsAreInScopedSyncResult " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestScopedDecisionRequiredKeepsLogPath()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary

    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotMerge, edpBlock), "the import begins under the session policy"
    cOp.ResolvePrompt vbOKOnly, "Notice", "Finished."
    TestAssert cOp.ResolvePrompt(vbYesNo, "Confirm", "Overwrite?").Blocked, "the policy blocks the prompt"

    Set dResult = ParseJson(FinishScopedOperation(cOp, New Dictionary, eorDecisionRequired, _
        "C:\src\logs\Merge_1.log"))
    TestAssert dResult("decision_required") = True, "the blocked prompt is primary"
    TestAssert dResult("logPath") = "C:\src\logs\Merge_1.log", "the decision-required result names the log"
    TestAssert dResult("decisions").Count = 2, "the journal holds the acknowledged and the blocked prompt"
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestScopedDecisionRequiredKeepsLogPath cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestScopedDecisionRequiredKeepsLogPath " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestSingleObjectErrorAfterAnsweredPromptKeepsJournal()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary
    Dim varScenario As Variant
    Dim strResolution As String
    Dim strJson As String

    For Each varScenario In AnsweredPromptScenarios
        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotMerge, strResolution)
        strJson = FailSingleObjectOperation(cOp, "C:\src\logs\Merge_1.log", 91, "Object variable not set")
        AssertJournalCarried strJson, strResolution, "runtime error, " & varScenario
        Set dResult = ParseJson(strJson)
        TestAssert dResult("success") = False, varScenario & ": the runtime error is a failure"
        TestAssert dResult("error") = "Object variable not set", varScenario & ": the error is kept"
        TestAssert dResult("errorNumber") = 91, varScenario & ": the error number is kept"
        ClearSessionPolicy cOp
    Next varScenario

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestSingleObjectErrorAfterAnsweredPromptKeepsJournal cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestSingleObjectErrorAfterAnsweredPromptKeepsJournal " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestTestRunResultCarriesAnsweredPrompts()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim varScenario As Variant
    Dim strResolution As String
    Dim strRun As String

    ' The synchronous test JSON as ExecuteTests builds it, a passed run and a failed one.
    For Each varScenario In AnsweredPromptScenarios
        Set cOp = BeginWithAnsweredPrompt(CStr(varScenario), eotTestRun, strResolution)
        strRun = "{""success"":true,""allPassed"":true,""passed"":3}"
        AssertJournalCarried OverlayDecisionJournal(strRun, cOp.Decisions), strResolution, _
            "test run passed, " & varScenario
        strRun = "{""success"":false,""allPassed"":false,""error"":""A test failed.""}"
        AssertJournalCarried OverlayDecisionJournal(strRun, cOp.Decisions), strResolution, _
            "test run failed, " & varScenario
        TestAssert InStr(OverlayDecisionJournal(strRun, cOp.Decisions), "A test failed.") > 0, _
            varScenario & ": the failure text is kept"
        cOp.Finish eorSuccess
        ClearSessionPolicy cOp
    Next varScenario

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestTestRunResultCarriesAnsweredPrompts cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestTestRunResultCarriesAnsweredPrompts " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestEmptyJournalLeavesResultShapeUnchanged()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim dResult As Dictionary
    Dim strRun As String

    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotMerge, edpBlock), "the import begins under the session policy"
    TestAssert InStr(FinishSingleObjectImport(cOp, "C:\src\logs\Merge_1.log", New Collection), "decisions") = 0, _
        "an import that met no prompt has no decisions key"

    TestAssert cOp.Begin(eotExport, edpBlock), "the export begins under the session policy"
    TestAssert InStr(FinishSingleObjectExport(cOp, "C:\src\logs\Export_1.log", New Dictionary), "decisions") = 0, _
        "an export that met no prompt has no decisions key"

    TestAssert cOp.Begin(eotExport, edpBlock), "the scoped export begins under the session policy"
    Set dResult = New Dictionary
    dResult.Add "success", True
    TestAssert InStr(FinishScopedOperation(cOp, dResult, eorSuccess, "C:\src\logs\Export_1.log"), "decisions") = 0, _
        "a scoped sync that met no prompt has no decisions key"

    TestAssert cOp.Begin(eotMerge, edpBlock), "the failing import begins under the session policy"
    TestAssert InStr(FailSingleObjectOperation(cOp, "C:\src\logs\Merge_1.log", 91, "Object variable not set"), "decisions") = 0, _
        "a runtime error that met no prompt has no decisions key"

    TestAssert cOp.Begin(eotTestRun, edpBlock), "the test run begins under the session policy"
    strRun = "{""success"":true,""allPassed"":true}"
    TestAssert OverlayDecisionJournal(strRun, cOp.Decisions) = strRun, "an empty journal leaves the test JSON as it was"
    TestAssert OverlayDecisionJournal(strRun, Nothing) = strRun, "no journal leaves the test JSON as it was"
    cOp.Finish eorSuccess
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestEmptyJournalLeavesResultShapeUnchanged cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestEmptyJournalLeavesResultShapeUnchanged " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestSingleObjectErrorBeforeBeginLeavesRunningRootAlone()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strJson As String

    ' An error before the call's own Begin: the running root belongs to someone else.
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    strJson = FailSingleObjectOperation(Nothing, vbNullString, 5, "Invalid procedure call")
    TestAssert cOp.Status = eosRunning, "the other root is still running"
    TestAssert InStr(strJson, """success"":false") > 0, "the result is not a success"
    TestAssert InStr(strJson, "logPath") = 0, "no log is named when none was written"
    cRoot.Complete eorSuccess

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestSingleObjectErrorBeforeBeginLeavesRunningRootAlone cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestSingleObjectErrorBeforeBeginLeavesRunningRootAlone " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestApiErrorAfterOwnFinishDoesNotFinishAgain()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim dDone As Dictionary
    Dim strJson As String

    ' An error raised in ImportByType or ExportByType after FinishScopedSync already
    ' finished the call's own root. The root completes once, with its own result.
    Set cOp = NewUnattendedOperation
    TestAssert cOp.Begin(eotExport, edpBlock), "the export begins"
    cOp.Finish eorSuccess
    Set dDone = cOp.LastCompletion

    strJson = FailSingleObjectOperation(cOp, "C:\src\logs\Export_1.log", 6, "Overflow")
    TestAssert cOp.LastCompletion Is dDone, "no second completion is emitted"
    TestAssert cOp.Result = eorSuccess, "the finished root keeps its result"
    TestAssert cOp.DecisionPolicy = edpAsk, "the completed request leaves no policy"
    TestAssert InStr(strJson, """success"":false") > 0, "the result is not a success"
    TestAssert InStr(strJson, """error"":""Overflow""") > 0, "the result carries the error"
    TestAssert InStr(strJson, """errorNumber"":6") > 0, "the result carries the error number"
    ClearSessionPolicy cOp

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestApiErrorAfterOwnFinishDoesNotFinishAgain cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestApiErrorAfterOwnFinishDoesNotFinishAgain " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRetiredSetterAndHarmlessCleanup()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Set cOp = NewUnattendedOperation
    TestAssert InStr(SetSessionPolicy(cOp, "block"), "policy_activation_retired") > 0, "setter calls must migrate"
    TestAssert InStr(ClearSessionPolicy(cOp), "true") > 0, "legacy cleanup succeeds harmlessly"
    TestAssert cOp.DecisionPolicy = edpAsk, "neither call leaves a policy"

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRetiredSetterAndHarmlessCleanup cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRetiredSetterAndHarmlessCleanup " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRequestPoliciesBelongToTheirOperations()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cFirst As clsOperation
    Dim cSecond As clsOperation
    Dim cFirstRoot As clsRootOperationLease
    Dim cSecondRoot As clsRootOperationLease
    Set cFirst = NewUnattendedOperation
    Set cSecond = NewUnattendedOperation
    Set cFirstRoot = cFirst.TryBeginRoot(eotOther, edpBlock)
    Set cSecondRoot = cSecond.TryBeginRoot(eotOther, edpDecline)
    cFirstRoot.Complete eorSuccess
    TestAssert cSecond.DecisionPolicy = edpDecline, "another owner retains its policy"
    TestAssert cFirst.PublicId <> cSecond.PublicId, "public IDs are unique across roots"
    cSecondRoot.Complete eorSuccess

CleanUp:
    On Error Resume Next
    If Not cFirstRoot Is Nothing Then cFirstRoot.Complete eorFailed
    If Not cSecondRoot Is Nothing Then cSecondRoot.Complete eorFailed
    If Not cFirst Is Nothing Then
        If cFirst.IsActive Then cFirst.Finish eorFailed
    End If
    If Not cSecond Is Nothing Then
        If cSecond.IsActive Then cSecond.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRequestPoliciesBelongToTheirOperations cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRequestPoliciesBelongToTheirOperations " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRuntimeErrorRidesOnTerminalCallback()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dDone As Dictionary

    ' An async caller reads only the callback, so the error text has to be on it.
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    cRoot.RecordRuntimeError "Division by zero", 11
    cRoot.Complete eorFailed

    Set dDone = cOp.LastCompletion
    TestAssert dDone("type") = "error", "a runtime error finishes as an error callback"
    TestAssert dDone("message") = "Division by zero", "the message is the error text"
    TestAssert dDone("runtime_error") = "Division by zero", "the callback carries runtime_error"
    TestAssert dDone("errorNumber") = 11, "the callback carries the error number"
    TestAssert Not dDone.Exists("decision_required"), "no prompt was blocked"

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRuntimeErrorRidesOnTerminalCallback cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRuntimeErrorRidesOnTerminalCallback " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRuntimeErrorAfterBlockedPromptRidesOnDecisionCallback()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dDone As Dictionary

    ' The decision is primary; the runtime error rides along, as in the sync result.
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
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

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRuntimeErrorAfterBlockedPromptRidesOnDecisionCallback cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRuntimeErrorAfterBlockedPromptRidesOnDecisionCallback " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestRuntimeErrorBelongsToItsRoot()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim cOp As clsOperation
    Dim cFirst As clsRootOperationLease
    Dim cSecond As clsRootOperationLease

    ' Neither the next root nor a lease that no longer holds the root can carry it over.
    Set cOp = NewUnattendedOperation
    Set cFirst = cOp.TryBeginRoot(eotOther, edpBlock)
    cFirst.RecordRuntimeError "Division by zero", 11
    cFirst.Complete eorFailed

    Set cSecond = cOp.TryBeginRoot(eotOther, edpBlock)
    cFirst.RecordRuntimeError "Stale lease", 5
    cSecond.Complete eorSuccess
    TestAssert cOp.LastCompletion("type") = "complete", "the next root succeeds"
    TestAssert Not cOp.LastCompletion.Exists("runtime_error"), "no runtime_error on the next root"
    TestAssert Not cOp.LastCompletion.Exists("errorNumber"), "no error number on the next root"

CleanUp:
    On Error Resume Next
    If Not cFirst Is Nothing Then cFirst.Complete eorFailed
    If Not cSecond Is Nothing Then cSecond.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestRuntimeErrorBelongsToItsRoot cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestRuntimeErrorBelongsToItsRoot " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestInteractionModeSelectionReportsEffectiveMode()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Dim dResult As Dictionary
    Dim intMode As Long
    Set cOp = NewUnattendedOperation
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNonInteractive))
    TestAssert Not dResult("success"), "standalone mode 2 is retired"
    TestAssert dResult("requested_mode") = 2 And dResult("effective_mode") = 0, "the response preserves requested and effective modes"
    TestAssert dResult("error_pattern") = ERR_INTERACTION_MODE_REFUSED, "mode retirement is explicit"
    For intMode = eimSilent To eimNormal Step -1
        Set dResult = ParseJson(SelectInteractionMode(cOp, intMode))
        TestAssert dResult("success"), "legacy idle mode remains available"
        TestAssert dResult("effective_mode") = intMode, "effective mode is truthful"
    Next intMode

CleanUp:
    On Error Resume Next
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestInteractionModeSelectionReportsEffectiveMode cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestInteractionModeSelectionReportsEffectiveMode " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
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
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dResult As Dictionary
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    cRoot.RecordRuntimeError "Failed operation", 5
    cRoot.Complete eorFailed
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNormal))
    TestAssert dResult("success"), "failure requires no client cleanup call"
    TestAssert cOp.DecisionPolicy = edpAsk, "failure releases the request policy"
    TestAssert cOp.Result = eorFailed, "mode selection preserves the failed result"

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestInteractiveSelectionAfterFailedPolicyCleanup cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestInteractiveSelectionAfterFailedPolicyCleanup " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestSilentSelectionInsidePolicyScopeIsRefused()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim dResult As Dictionary
    Set cOp = NewUnattendedOperation
    Set cRoot = cOp.TryBeginRoot(eotOther, edpDecline)
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimSilent))
    TestAssert Not dResult("success"), "another request cannot change the owner mode"
    TestAssert dResult("effective_mode") = 2, "the request remains noninteractive"
    Set dResult = ParseJson(SelectInteractionMode(cOp, eimNonInteractive))
    TestAssert Not dResult("success"), "mode-only activation is retired even when a root is noninteractive"
    TestAssert cOp.DecisionPolicy = edpDecline, "the owner retains its policy"
    cRoot.Complete eorSuccess

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If Not cOp Is Nothing Then
        If cOp.IsActive Then cOp.Finish eorFailed
    End If
    If Err.Number <> 0 Then TestAssert False, "TestSilentSelectionInsidePolicyScopeIsRefused cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestSilentSelectionInsidePolicyScopeIsRefused " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
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
Private Sub ClearPolicyInOneCall(ByVal cOp As clsOperation)
    TestAssert InStr(ClearSessionPolicy(cOp), """success"":true") > 0, "clearing succeeds"
End Sub

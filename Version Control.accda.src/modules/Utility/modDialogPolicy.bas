Attribute VB_Name = "modDialogPolicy"
'---------------------------------------------------------------------------------------
' Module    : modDialogPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Noninteractive decision policy for add-in prompts, and the result JSON
'           : for refused, started and blocked operations.
'           : Informational OK dialogs are acknowledged. A Yes/No or conflict
'           : prompt is applied only when the active policy names that choice.
'           : Anything else is recorded as decision_required and is not approved.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Utility")


' Error patterns returned when a request is refused or stops for a decision.
Public Const ERR_INVALID_DECISION_POLICY As String = "invalid_decision_policy"
Public Const ERR_INTERACTION_MODE_REFUSED As String = "interaction_mode_refused"
Public Const ERR_INVALID_INTERACTION_MODE As String = "invalid_interaction_mode"
Public Const ERR_MERGE_NOT_AVAILABLE As String = "merge_not_available"
Public Const ERR_OPERATION_ALREADY_RUNNING As String = "operation_already_running"
Public Const ERR_DECISION_REQUIRED As String = DECISION_REQUIRED
Public Const ERR_PROJECT_NOT_COMPILED As String = "project_not_compiled"


' The low bits of a MsgBox style select the button set (vbOKOnly to vbRetryCancel).
Public Const MB_BUTTON_STYLE_MASK As Long = 7


'---------------------------------------------------------------------------------------
' Procedure : PolicyTable
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : The one list of policy names a caller can pass, mapped to eDecisionPolicy.
'           : Parsing, naming and the invalid-policy message all read it, so a new
'           : policy is registered here (and in eDecisionPolicy). Keys are lower case
'           : and in the order the invalid-policy message lists them.
'---------------------------------------------------------------------------------------
'
Private Function PolicyTable() As Dictionary

    Static dTable As Dictionary

    If dTable Is Nothing Then
        Set dTable = New Dictionary
        dTable.Add "block", edpBlock
        dTable.Add "prefer_source", edpPreferSource
        dTable.Add "prefer_database", edpPreferDatabase
        dTable.Add "skip", edpSkip
        dTable.Add "decline", edpDecline
    End If
    Set PolicyTable = dTable

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseDecisionPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Map a policy name to eDecisionPolicy. Returns False for unknown names.
'           : Accepted names are the keys of PolicyTable.
'---------------------------------------------------------------------------------------
'
Public Function ParseDecisionPolicy(ByVal strPolicy As String, ByRef intPolicy As eDecisionPolicy) As Boolean

    Dim dTable As Dictionary
    Dim strName As String

    Set dTable = PolicyTable()
    strName = LCase$(Trim$(strPolicy))
    If dTable.Exists(strName) Then
        intPolicy = dTable.Item(strName)
        ParseDecisionPolicy = True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : DecisionPolicyName
' Date      : 09/29/2026
' Purpose   : The name a caller passes for a policy. Empty for edpAsk, which is not a
'           : noninteractive policy.
'---------------------------------------------------------------------------------------
'
Public Function DecisionPolicyName(ByVal intPolicy As eDecisionPolicy) As String

    Dim dTable As Dictionary
    Dim varName As Variant

    Set dTable = PolicyTable()
    For Each varName In dTable.Keys
        If dTable.Item(varName) = intPolicy Then
            DecisionPolicyName = varName
            Exit Function
        End If
    Next varName

End Function


'---------------------------------------------------------------------------------------
' Procedure : ConflictResolutionName
' Date      : 09/29/2026
' Purpose   : The resolution reported when a policy answered a conflict: the policy
'           : outcome identifies the transferred or preserved side by direction.
'---------------------------------------------------------------------------------------
'
Public Function ConflictResolutionName(ByVal intPolicy As eDecisionPolicy, _
    Optional ByVal intDirection As eIndexOperationType = eatImport) As String

    If intDirection = eatImport Then
        If intPolicy = edpPreferSource Then
            ConflictResolutionName = "import_source"
        Else
            ConflictResolutionName = DECISION_KEEP_DATABASE
        End If
    ElseIf intPolicy = edpPreferDatabase Then
        ConflictResolutionName = "export_database"
    Else
        ConflictResolutionName = "keep_source"
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetSessionPolicy
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Open a caller-owned noninteractive scope on this operation and return the
'           : JSON result for SetOperationPolicy. The operation holds the scope
'           : between calls: one it already holds is replaced, not nested. Finish
'           : leaves the scope open; only ClearSessionPolicy closes it. The
'           : operation is passed in so the rule can be tested without touching the
'           : session operation.
'---------------------------------------------------------------------------------------
'
Public Function SetSessionPolicy(ByVal cOp As clsOperation, ByVal strPolicy As String) As String
    SetSessionPolicy = RefusalJson("policy_activation_retired", T("Persistent policy activation is no longer supported. Supply the decision policy on the operation request."), False)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearSessionPolicy
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Close the scope SetSessionPolicy opened and return the JSON result for
'           : ClearOperationPolicy. Safe when no policy is set, and safe to call twice.
'---------------------------------------------------------------------------------------
'
Public Function ClearSessionPolicy(ByVal cOp As clsOperation) As String
    ClearSessionPolicy = "{""success"":true}"
End Function


'---------------------------------------------------------------------------------------
' Procedure : SelectInteractionMode
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Report whether the requested mode took effect. The setter preserves
'           : enclosing scopes and cannot relax a running root. Read back the mode
'           : rather than treating an ignored selection as acceptance. No scope is
'           : closed here, including one left open after failed caller cleanup.
'---------------------------------------------------------------------------------------
'
Public Function SelectInteractionMode(ByVal cOp As clsOperation, ByVal intMode As Long) As String

    Dim dResult As Dictionary
    Dim strPattern As String
    Dim strError As String
    Dim intEffective As eInteractionMode

    If intMode < eimNormal Or intMode > eimNonInteractive Then
        strPattern = ERR_INVALID_INTERACTION_MODE
        strError = T("Unknown interaction mode. Use 0 (normal), 1 (silent), or 2 (noninteractive).")
    ElseIf intMode = eimNonInteractive Then
        strPattern = ERR_INTERACTION_MODE_REFUSED
        strError = T("Standalone noninteractive mode is no longer supported. Supply a decision policy on the operation request.")
    ElseIf cOp.IsActive Then
        strPattern = ERR_INTERACTION_MODE_REFUSED
        strError = T("An active operation owns its interaction mode.")
    Else
        cOp.InteractionMode = intMode
        If cOp.InteractionMode <> intMode Then
            strPattern = ERR_INTERACTION_MODE_REFUSED
            strError = T("The requested interaction mode could not take effect. An enclosing noninteractive scope or an active operation must be released by its owner first.")
        End If
    End If

    intEffective = cOp.InteractionMode
    Set dResult = New Dictionary
    dResult.Add "success", (Len(strPattern) = 0)
    dResult.Add "requested_mode", intMode
    dResult.Add "effective_mode", intEffective
    If Len(strPattern) > 0 Then
        dResult.Add "error_pattern", strPattern
        dResult.Add "error", strError
    End If
    SelectInteractionMode = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : InteractionIsNonInteractive
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True when the current operation must not open add-in UI.
'---------------------------------------------------------------------------------------
'
Public Function InteractionIsNonInteractive() As Boolean
    InteractionIsNonInteractive = (Operation.InteractionMode = eimNonInteractive)
End Function


'---------------------------------------------------------------------------------------
' Procedure : MCPCallBeforeRoot
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : True when an MCP call has registered its callback but no root is running.
'           : The call's policy scope opens with its root, so until then the mode reads
'           : normal even for a noninteractive call. Nobody is at the screen either
'           : way, so a logged error stays in the log instead of opening a message box.
'           : The MCP registration is released when the root finishes.
'           : Pass MCP.IsActive and Operation.IsActive.
'---------------------------------------------------------------------------------------
'
Public Function MCPCallBeforeRoot(ByVal blnMCPActive As Boolean, ByVal blnRootActive As Boolean) As Boolean
    MCPCallBeforeRoot = blnMCPActive And Not blnRootActive
End Function


'---------------------------------------------------------------------------------------
' Procedure : MainFormWaitsForRoot
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : True when closing frmVCSMain must ask to cancel the running root first.
'           : A test run only counts while this project's runner is running it: closing
'           : the form cannot stop a run that nothing here is driving, so asking would
'           : refuse every close until the heartbeat timed out.
'           : Pass Operation.Status, Operation.OperationType, and whether
'           : TestRunner.State is etrsRunning.
'---------------------------------------------------------------------------------------
'
Public Function MainFormWaitsForRoot(ByVal intStatus As eOperationState, ByVal intType As eOperationType, _
    ByVal blnRunnerRunning As Boolean) As Boolean

    If intStatus <> eosRunning Then Exit Function
    If intType = eotTestRun Then
        MainFormWaitsForRoot = blnRunnerRunning
    Else
        MainFormWaitsForRoot = True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : InteractionIsNormal
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True when the user is present and dialogs may be shown.
'---------------------------------------------------------------------------------------
'
Public Function InteractionIsNormal() As Boolean
    InteractionIsNormal = (Operation.InteractionMode = eimNormal)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ShowMainForm
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Open the main form. Every site that shows it goes through here so a
'           : noninteractive operation keeps it hidden.
'---------------------------------------------------------------------------------------
'
Public Sub ShowMainForm()
    If InteractionIsNonInteractive() Then
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    Else
        DoCmd.OpenForm "frmVCSMain"
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetMainFormVisible
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Show the open main form for an attended run, and keep it hidden for a
'           : noninteractive one, including a form that was already visible.
'---------------------------------------------------------------------------------------
'
Public Sub SetMainFormVisible(ByVal frmMain As Form_frmVCSMain)
    frmMain.Visible = Not InteractionIsNonInteractive()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveNonInteractivePrompt
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Choose a MsgBox result without showing UI.
'           : OK-only prompts are acknowledged. Decline returns the non-destructive
'           : button (No, Cancel, or Abort). Every other policy leaves a required
'           : confirmation blocked and still returns that non-destructive button so
'           : callers that proceed only on Yes/OK do not apply the change.
'---------------------------------------------------------------------------------------
'
Public Function ResolveNonInteractivePrompt(ByVal intButtons As Long, _
    ByVal intPolicy As eDecisionPolicy) As clsPromptOutcome

    Dim cOutcome As clsPromptOutcome
    Dim intStyle As Long

    intStyle = intButtons And MB_BUTTON_STYLE_MASK

    ' The defaults answer an OK-only prompt: acknowledged, nothing blocked.
    Set cOutcome = New clsPromptOutcome
    cOutcome.Blocked = False
    cOutcome.Resolution = DECISION_ACKNOWLEDGED
    cOutcome.Result = vbOK
    Set ResolveNonInteractivePrompt = cOutcome
    If intStyle = vbOKOnly Then Exit Function

    cOutcome.Result = NonDestructiveResult(intStyle)

    If intPolicy = edpDecline Then
        cOutcome.Resolution = DECISION_DECLINED
        Exit Function
    End If

    ' prefer_source / prefer_database / skip apply to merge conflicts, not to a
    ' generic confirmation. block, ask, and those conflict policies all stop here.
    cOutcome.Resolution = DECISION_REQUIRED
    cOutcome.Blocked = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : NewDecision
' Date      : 09/30/2026
' Purpose   : Build one decision journal entry. Kind, resolution and detail are wire
'           : values; title and message are the text of the prompt or conflict.
'---------------------------------------------------------------------------------------
'
Public Function NewDecision(ByVal strKind As String, ByVal strTitle As String, _
    ByVal strMessage As String, ByVal strResolution As String, _
    Optional ByVal strDetail As String) As clsDecisionRecord

    Dim cRecord As clsDecisionRecord

    Set cRecord = New clsDecisionRecord
    cRecord.Kind = strKind
    cRecord.Title = strTitle
    cRecord.Message = strMessage
    cRecord.Resolution = strResolution
    cRecord.Detail = strDetail
    Set NewDecision = cRecord

End Function


'---------------------------------------------------------------------------------------
' Procedure : ConflictActionForPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Map a decision policy to a conflict resolution.
'           : intRequested is the action the conflict itself asks for.
'           : Preferences name the side preserved, independently of direction.
'           : No policy authorizes a conflicting deletion. ercNone requires review.
'---------------------------------------------------------------------------------------
'
Public Function ConflictActionForPolicy(ByVal intPolicy As eDecisionPolicy, _
    Optional ByVal intRequested As eResolveConflict = ercNone, _
    Optional ByVal intDirection As eIndexOperationType = eatImport, _
    Optional ByVal strClassification As String = "overwrite") As eResolveConflict

    ' A new source file never authorizes deletion and never blocks unrelated work.
    If strClassification = "untracked_orphan" Then
        ConflictActionForPolicy = ercSkip
        Exit Function
    End If

    Select Case intPolicy
        Case edpSkip, edpDecline
            ConflictActionForPolicy = ercSkip
        Case edpPreferSource
            If intRequested = ercSkip Then
                ConflictActionForPolicy = ercSkip
            ElseIf intDirection <> eatImport Then
                ConflictActionForPolicy = ercSkip
            ElseIf intRequested <> ercDelete Then
                ConflictActionForPolicy = ercOverwrite
            End If
        Case edpPreferDatabase
            If intDirection = eatImport Then
                ConflictActionForPolicy = ercSkip
            ElseIf intRequested <> ercDelete Then
                ConflictActionForPolicy = ercOverwrite
            End If
        Case Else
            ConflictActionForPolicy = ercNone
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : DecisionRequiredJson
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Structured failure for an operation that stopped on an unresolved prompt.
'---------------------------------------------------------------------------------------
'
Public Function DecisionRequiredJson(Optional ByVal colDecisions As Collection) As String
    DecisionRequiredJson = ConvertToJson(DecisionRequiredResult(colDecisions))
End Function


'---------------------------------------------------------------------------------------
' Procedure : DecisionRequiredResult
' Date      : 09/29/2026
' Purpose   : The decision_required result as a dictionary, so callers can add fields
'           : before encoding. Uses the running operation's decisions unless the
'           : decisions of a finished operation are passed in.
'---------------------------------------------------------------------------------------
'
Public Function DecisionRequiredResult(Optional ByVal colDecisions As Collection) As Dictionary

    Dim dResult As Dictionary

    If colDecisions Is Nothing Then Set colDecisions = Operation.Decisions
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error_pattern", ERR_DECISION_REQUIRED
    dResult.Add "decision_required", True
    dResult.Add "error", DecisionRequiredMessage()
    dResult.Add "decisions", colDecisions
    Set DecisionRequiredResult = dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddDecisionJournal
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : Attach the decision journal to a synchronous result, as the terminal
'           : callback does: only when the journal is not empty, so a run that met no
'           : prompt keeps its result shape. A decision_required result already carries
'           : its journal and is left alone.
'---------------------------------------------------------------------------------------
'
Public Sub AddDecisionJournal(dResult As Object, ByVal colDecisions As Collection)

    If colDecisions Is Nothing Then Exit Sub
    If colDecisions.Count = 0 Then Exit Sub
    If dResult.Exists("decisions") Then Exit Sub
    dResult.Add "decisions", colDecisions

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OverlayDecisionJournal
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : Add a decision journal to a JSON result. The result is returned as it
'           : came when the journal is empty, so the shape only changes for a run that
'           : met a prompt. If the payload cannot be parsed it is returned unchanged
'           : rather than lose it.
'---------------------------------------------------------------------------------------
'
Public Function OverlayDecisionJournal(ByVal strJson As String, _
    ByVal colDecisions As Collection) As String

    Dim dParsed As Object

    OverlayDecisionJournal = strJson
    If Len(strJson) = 0 Then Exit Function
    If colDecisions Is Nothing Then Exit Function
    If colDecisions.Count = 0 Then Exit Function

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dParsed = ParseJson(strJson)
    AddDecisionJournal dParsed, colDecisions
    OverlayDecisionJournal = ConvertToJson(dParsed)

ErrHandler:
    ' Nothing to add when the payload was malformed; the caller keeps the original.

End Function


'---------------------------------------------------------------------------------------
' Procedure : OverlayDecisionRequired
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Force a JSON result to report decision_required. Closing or bypassing
'           : a prompt must not leave success true.
'---------------------------------------------------------------------------------------
'

'---------------------------------------------------------------------------------------
' Procedure : SetDecisionRequiredOutcome
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Normalize blocked decisions into the authoritative failed result.
'---------------------------------------------------------------------------------------
'
Public Sub SetDecisionRequiredOutcome(ByVal dResult As Object, ByVal colDecisions As Collection)
    ' Preserve the runner failure before assigning the primary decision fields.
    ' Runtime exceptions already have runtime_error/errorNumber and keep that shape.
    If dResult.Exists("error_pattern") Then
        If dResult("error_pattern") = ERR_PROJECT_NOT_COMPILED Then
            dResult("run_error_pattern") = dResult("error_pattern")
            If dResult.Exists("error") Then dResult("run_error") = dResult("error")
        End If
    End If
    dResult("success") = False
    dResult("error_pattern") = ERR_DECISION_REQUIRED
    dResult("decision_required") = True
    dResult("error") = DecisionRequiredMessage()
    Set dResult("decisions") = colDecisions
End Sub



'---------------------------------------------------------------------------------------
' Procedure : OverlayDecisionRequired
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Preserve an existing result while applying blocked-decision precedence.
'---------------------------------------------------------------------------------------
'
Public Function OverlayDecisionRequired(ByVal strJson As String) As String

    Dim dParsed As Object

    LogUnhandledErrors
    On Error GoTo Fallback

    If Len(strJson) = 0 Then
        OverlayDecisionRequired = DecisionRequiredJson()
        Exit Function
    End If

    Set dParsed = ParseJson(strJson)
    SetDecisionRequiredOutcome dParsed, Operation.Decisions
    OverlayDecisionRequired = ConvertToJson(dParsed)
    Exit Function

Fallback:
    OverlayDecisionRequired = DecisionRequiredJson()

End Function


'---------------------------------------------------------------------------------------
' Procedure : InvalidPolicyMessage
' Date      : 09/29/2026
' Purpose   : The message for an unknown decision policy, listing the valid names.
'---------------------------------------------------------------------------------------
'
Public Function InvalidPolicyMessage() As String

    Dim varNames As Variant
    Dim strList As String
    Dim lngIdx As Long

    ' The policy names are wire values, so they stay out of the translated sentence.
    varNames = PolicyTable().Keys
    For lngIdx = LBound(varNames) To UBound(varNames)
        If lngIdx > LBound(varNames) Then strList = strList & ", "
        If lngIdx = UBound(varNames) And lngIdx > LBound(varNames) Then strList = strList & "or "
        strList = strList & varNames(lngIdx)
    Next lngIdx
    InvalidPolicyMessage = T("Unknown decision policy. Use {0}.", var0:=strList)

End Function


'---------------------------------------------------------------------------------------
' Procedure : OperationRunningMessage
' Date      : 09/29/2026
' Purpose   : The message for a request refused because another operation is running.
'---------------------------------------------------------------------------------------
'
Public Function OperationRunningMessage() As String
    OperationRunningMessage = T("Another operation is already running.")
End Function


'---------------------------------------------------------------------------------------
' Procedure : PolicyRequestRefusal
' Date      : 09/30/2026
' Purpose   : The refusal JSON for a request to set a decision policy on this operation,
'           : or an empty string when the operation is ready to take one. A running or
'           : staged operation keeps the policy it began with. The operation is passed
'           : in so the rule can be tested without touching the session operation.
'---------------------------------------------------------------------------------------
'
Public Function PolicyRequestRefusal(ByVal cOp As clsOperation) As String
    If cOp.Status <> eosReady Then
        PolicyRequestRefusal = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : DecisionRequiredMessage
' Date      : 09/29/2026
' Purpose   : The message for an operation that stopped on an unanswered prompt.
'---------------------------------------------------------------------------------------
'
Private Function DecisionRequiredMessage() As String
    DecisionRequiredMessage = T("A required decision was not covered by the decision policy.")
End Function


'---------------------------------------------------------------------------------------
' Procedure : RefusalJson
' Date      : 09/29/2026
' Purpose   : Result for a request refused before it started. When blnPostCallback is
'           : set and MCP is active, the same payload is also posted as the completion
'           : callback, since an async caller only reads the callback.
'---------------------------------------------------------------------------------------
'
Public Function RefusalJson(ByVal strPattern As String, ByVal strMessage As String, _
    Optional ByVal blnPostCallback As Boolean = True) As String

    Dim dResult As Dictionary

    Set dResult = New Dictionary
    dResult.Add "success", False
    If Len(strPattern) > 0 Then dResult.Add "error_pattern", strPattern
    dResult.Add "error", strMessage
    ' Refusals belong to the incoming request. Its dispatcher owns any delivery.
    RefusalJson = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : StartedJson
' Date      : 09/29/2026
' Purpose   : Start result: the operation began. This is never a final result; the
'           : outcome arrives through the completion callback. operation_id is the id
'           : the MCP server registered, and is omitted when MCP is not active.
'---------------------------------------------------------------------------------------
'
Public Function StartedJson() As String

    Dim dResult As Dictionary

    Set dResult = New Dictionary
    dResult.Add "success", True
    dResult.Add "started", True
    AddOperationIdentity dResult, Operation
    StartedJson = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : RuntimeErrorJson
' Date      : 09/29/2026
' Purpose   : Result for a run that hit a runtime error. If a prompt was blocked,
'           : decision_required is the primary result and carries the decisions, with
'           : the runtime error alongside. Otherwise it is a plain failure.
'---------------------------------------------------------------------------------------
'
Public Function RuntimeErrorJson(ByVal strDescription As String, ByVal lngNumber As Long) As String

    Dim dResult As Dictionary

    If Operation.DecisionBlocked Then
        Set dResult = DecisionRequiredResult()
        dResult.Add "runtime_error", strDescription
    Else
        Set dResult = New Dictionary
        dResult.Add "success", False
        dResult.Add "error", strDescription
    End If
    dResult.Add "errorNumber", lngNumber
    RuntimeErrorJson = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : PostRefusal
' Date      : 09/29/2026
' Purpose   : Post a pre-root refusal without using an active root's callback.
'---------------------------------------------------------------------------------------
'
Private Sub PostRefusal(ByVal dResult As Dictionary, ByVal strMessage As String)
    ' An active callback belongs to its root, including during direct VBA calls.
    If Operation.IsActive Then Exit Sub
    If MCP.IsActive Then MCP.PostCallback "error", -1, -1, strMessage, dResult
End Sub











'---------------------------------------------------------------------------------------
' Procedure : NonDestructiveResult
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Button that refuses the prompt: No, Cancel, or Abort.
'---------------------------------------------------------------------------------------
'
Private Function NonDestructiveResult(ByVal intStyle As Long) As VbMsgBoxResult

    ' Prompt callers must require an affirmative response before performing the
    ' action being offered. No and Cancel paths must avoid that action.
    ' For Yes/No/Cancel, return Cancel because No may select another action.
    Select Case intStyle
        Case vbOKCancel, vbRetryCancel, vbYesNoCancel
            NonDestructiveResult = vbCancel
        Case vbYesNo
            NonDestructiveResult = vbNo
        Case vbAbortRetryIgnore
            NonDestructiveResult = vbAbort
        Case Else
            NonDestructiveResult = vbCancel
    End Select

End Function

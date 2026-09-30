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
Public Const ERR_MERGE_NOT_AVAILABLE As String = "merge_not_available"
Public Const ERR_OPERATION_ALREADY_RUNNING As String = "operation_already_running"
Public Const ERR_DECISION_REQUIRED As String = "decision_required"


' The low bits of a MsgBox style select the button set (vbOKOnly to vbRetryCancel).
Public Const MB_BUTTON_STYLE_MASK As Long = 7


'---------------------------------------------------------------------------------------
' Procedure : ParseDecisionPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Map a policy name to eDecisionPolicy. Returns False for unknown names.
'           : Accepted names: block, prefer_source, prefer_database, skip, decline.
'---------------------------------------------------------------------------------------
'
Public Function ParseDecisionPolicy(ByVal strPolicy As String, ByRef intPolicy As eDecisionPolicy) As Boolean

    Select Case LCase$(Trim$(strPolicy))
        Case "block"
            intPolicy = edpBlock
        Case "prefer_source"
            intPolicy = edpPreferSource
        Case "prefer_database"
            intPolicy = edpPreferDatabase
        Case "skip"
            intPolicy = edpSkip
        Case "decline"
            intPolicy = edpDecline
        Case Else
            ParseDecisionPolicy = False
            Exit Function
    End Select
    ParseDecisionPolicy = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : BeginNonInteractive
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Open a noninteractive scope for one operation. Finish closes it. If the
'           : operation never starts, pass the returned token to
'           : Operation.CloseInteractionScope.
'---------------------------------------------------------------------------------------
'
Public Function BeginNonInteractive(ByVal strPolicy As String, Optional ByRef lngToken As Long) As Boolean

    Dim intPolicy As eDecisionPolicy

    lngToken = 0
    If Not ParseDecisionPolicy(strPolicy, intPolicy) Then
        BeginNonInteractive = False
        Exit Function
    End If
    lngToken = Operation.PushInteractionScope(eimNonInteractive, intPolicy, True)
    BeginNonInteractive = True

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
Public Sub ResolveNonInteractivePrompt(ByVal intButtons As Long, _
    ByVal intPolicy As eDecisionPolicy, _
    ByRef intResult As VbMsgBoxResult, _
    ByRef strResolution As String, _
    ByRef blnBlocked As Boolean)

    Dim intStyle As Long

    intStyle = intButtons And MB_BUTTON_STYLE_MASK
    blnBlocked = False
    strResolution = "acknowledged"
    intResult = vbOK

    If intStyle = vbOKOnly Then
        intResult = vbOK
        strResolution = "acknowledged"
        Exit Sub
    End If

    intResult = NonDestructiveResult(intStyle)

    If intPolicy = edpDecline Then
        strResolution = "declined"
        blnBlocked = False
        Exit Sub
    End If

    ' prefer_source / prefer_database / skip apply to merge conflicts, not to a
    ' generic confirmation. block, ask, and those conflict policies all stop here.
    strResolution = "decision_required"
    blnBlocked = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConflictActionForPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Map a decision policy to a conflict resolution.
'           : intRequested is the action the conflict itself asks for.
'           : prefer_source applies it, falling back to overwrite when none is set.
'           : ercNone means the policy does not resolve the conflict (block).
'           : prefer_database, skip, and decline keep the database object.
'---------------------------------------------------------------------------------------
'
Public Function ConflictActionForPolicy(ByVal intPolicy As eDecisionPolicy, _
    Optional ByVal intRequested As eResolveConflict = ercNone) As eResolveConflict

    Select Case intPolicy
        Case edpPreferSource
            If intRequested = ercNone Then
                ConflictActionForPolicy = ercOverwrite
            Else
                ConflictActionForPolicy = intRequested
            End If
        Case edpPreferDatabase, edpSkip, edpDecline
            ConflictActionForPolicy = ercSkip
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
Private Function DecisionRequiredResult(Optional ByVal colDecisions As Collection) As Dictionary

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
' Procedure : OverlayDecisionRequired
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Force a JSON result to report decision_required. Closing or bypassing
'           : a prompt must not leave success true.
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
    dParsed("success") = False
    dParsed("error_pattern") = ERR_DECISION_REQUIRED
    dParsed("decision_required") = True
    dParsed("error") = DecisionRequiredMessage()
    Set dParsed("decisions") = Operation.Decisions
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
    ' The policy names are wire values, so they stay out of the translated sentence.
    InvalidPolicyMessage = T("Unknown decision policy. Use {0}.", _
        var0:="block, prefer_source, prefer_database, skip, or decline")
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
    If blnPostCallback Then PostRefusal dResult, strMessage
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
    If Len(MCP.OperationId) > 0 Then dResult.Add "operation_id", MCP.OperationId
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
' Purpose   : Post a refusal to the MCP completion callback when MCP is active.
'---------------------------------------------------------------------------------------
'
Private Sub PostRefusal(ByVal dResult As Dictionary, ByVal strMessage As String)
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

    Select Case intStyle
        Case vbOKCancel, vbRetryCancel
            NonDestructiveResult = vbCancel
        Case vbYesNo, vbYesNoCancel
            NonDestructiveResult = vbNo
        Case vbAbortRetryIgnore
            NonDestructiveResult = vbAbort
        Case Else
            NonDestructiveResult = vbCancel
    End Select

End Function

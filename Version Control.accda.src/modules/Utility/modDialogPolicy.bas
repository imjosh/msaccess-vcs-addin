Attribute VB_Name = "modDialogPolicy"
'---------------------------------------------------------------------------------------
' Module    : modDialogPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Noninteractive decision policy for add-in prompts, plus listing and
'           : dismissal of add-in forms that are already open.
'           : Informational OK dialogs are acknowledged. A Yes/No or conflict
'           : prompt is applied only when the active policy names that choice.
'           : Anything else is recorded as decision_required and is not approved.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Utility")


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

    intStyle = intButtons And 7
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
Public Function DecisionRequiredJson() As String

    DecisionRequiredJson = "{""success"":false," & _
        """error_pattern"":""decision_required""," & _
        """decision_required"":true," & _
        """error"":""A required decision was not covered by the decision policy.""," & _
        """decisions"":" & Operation.DecisionsJson() & "}"

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

    On Error GoTo Fallback

    If Len(strJson) = 0 Then
        OverlayDecisionRequired = DecisionRequiredJson()
        Exit Function
    End If

    Set dParsed = ParseJson(strJson)
    dParsed("success") = False
    dParsed("error_pattern") = "decision_required"
    dParsed("decision_required") = True
    dParsed("error") = "A required decision was not covered by the decision policy."
    Set dParsed("decisions") = ParseJson(Operation.DecisionsJson())
    OverlayDecisionRequired = ConvertToJson(dParsed)
    Exit Function

Fallback:
    OverlayDecisionRequired = DecisionRequiredJson()

End Function


'---------------------------------------------------------------------------------------
' Procedure : JsonEscape
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Escape a string for inclusion in a hand-built JSON value.
'---------------------------------------------------------------------------------------
'
Public Function JsonEscape(ByVal strValue As String) As String

    Dim strOut As String

    strOut = strValue
    strOut = Replace(strOut, "\", "\\")
    strOut = Replace(strOut, """", "\""")
    strOut = Replace(strOut, vbCrLf, "\n")
    strOut = Replace(strOut, vbCr, "\n")
    strOut = Replace(strOut, vbLf, "\n")
    strOut = Replace(strOut, vbTab, "\t")
    JsonEscape = strOut

End Function


'---------------------------------------------------------------------------------------
' Procedure : ListAddinDialogs
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Enumerate open add-in forms. This requires Access/VBA to be responsive.
'           : A modal MsgBox or the conflict dialog blocks this call; use the MCP
'           : dialog inspector for those.
'---------------------------------------------------------------------------------------
'
Public Function ListAddinDialogs() As String

    Dim frm As Form
    Dim dRoot As Dictionary
    Dim colItems As Collection
    Dim dItem As Dictionary

    Set dRoot = New Dictionary
    Set colItems = New Collection

    On Error Resume Next
    For Each frm In Forms
        If IsAddinFormName(frm.Name) Then
            Set dItem = New Dictionary
            dItem.Add "name", frm.Name
            dItem.Add "caption", frm.Caption
            dItem.Add "hwnd", CStr(frm.hwnd)
            dItem.Add "visible", CBool(frm.Visible)
            dItem.Add "modal", CBool(frm.Modal)
            dItem.Add "kind", "addin_form"
            colItems.Add dItem
        End If
    Next frm
    On Error GoTo 0

    dRoot.Add "success", True
    dRoot.Add "dialogs", colItems
    dRoot.Add "operation_status", OperationStatusName()
    dRoot.Add "interaction", IIf(InteractionIsNonInteractive(), "noninteractive", "interactive")
    ListAddinDialogs = ConvertToJson(dRoot)

End Function


'---------------------------------------------------------------------------------------
' Procedure : DismissAddinDialog
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Close a completed add-in form, or request cancellation of a running
'           : operation. Closing a completed window does not cancel anything.
'           : action "cancel" sets the cancel/critical flag and does not call Finish
'           : itself; the running operation stops at its next check.
'---------------------------------------------------------------------------------------
'
Public Function DismissAddinDialog(ByVal strName As String, Optional ByVal strAction As String = "close") As String

    Dim strAct As String
    Dim blnRunning As Boolean
    Dim blnInterrupted As Boolean
    Dim blnClosed As Boolean
    Dim dRoot As Dictionary

    strAct = LCase$(Trim$(strAction))
    If Len(strAct) = 0 Then strAct = "close"

    Set dRoot = New Dictionary
    dRoot.Add "name", strName
    dRoot.Add "action", strAct

    If Not IsAddinFormName(strName) Then
        dRoot.Add "success", False
        dRoot.Add "error_pattern", "not_addin_form"
        dRoot.Add "error", "Only add-in forms (frmVCS*) can be dismissed here."
        dRoot.Add "interrupted", False
        DismissAddinDialog = ConvertToJson(dRoot)
        Exit Function
    End If

    If strAct <> "close" And strAct <> "cancel" Then
        dRoot.Add "success", False
        dRoot.Add "error_pattern", "unsupported_action"
        dRoot.Add "error", "Supported actions are close and cancel. Button clicks on VBA or Access dialogs use the MCP dialog inspector."
        dRoot.Add "interrupted", False
        DismissAddinDialog = ConvertToJson(dRoot)
        Exit Function
    End If

    blnRunning = OperationIsActive()

    If strAct = "close" And blnRunning Then
        dRoot.Add "success", False
        dRoot.Add "error_pattern", "operation_in_progress"
        dRoot.Add "error", "This window belongs to a running operation. Close is only for a finished results window. Pass action=cancel to interrupt the operation."
        dRoot.Add "interrupted", False
        dRoot.Add "closed", False
        DismissAddinDialog = ConvertToJson(dRoot)
        Exit Function
    End If

    If strAct = "cancel" And blnRunning Then
        blnInterrupted = True
        If TestRunner.State = etrsRunning Then
            TestRunner.Cancel
        Else
            Operation.ErrorLevel = eelCritical
            Operation.Result = eorCanceled
        End If
        dRoot.Add "success", True
        dRoot.Add "interrupted", True
        dRoot.Add "closed", False
        dRoot.Add "message", "Cancellation requested. The operation stops at its next check and Finish reports cancelled."
        DismissAddinDialog = ConvertToJson(dRoot)
        Exit Function
    End If

    If Not AddinFormIsLoaded(strName) Then
        dRoot.Add "success", False
        dRoot.Add "error_pattern", "not_open"
        dRoot.Add "error", "That add-in form is not open."
        dRoot.Add "interrupted", False
        dRoot.Add "closed", False
        DismissAddinDialog = ConvertToJson(dRoot)
        Exit Function
    End If

    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    blnClosed = (Err.Number = 0) And Not AddinFormIsLoaded(strName)
    If Err.Number <> 0 Then dRoot.Add "error", Err.Description
    Err.Clear
    On Error GoTo 0

    dRoot.Add "success", blnClosed
    dRoot.Add "closed", blnClosed
    dRoot.Add "interrupted", blnInterrupted
    If blnClosed Then dRoot.Add "message", "Closed a finished add-in window."
    DismissAddinDialog = ConvertToJson(dRoot)

End Function


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


'---------------------------------------------------------------------------------------
' Procedure : IsAddinFormName
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Add-in dialogs use the frmVCS name prefix.
'---------------------------------------------------------------------------------------
'
Private Function IsAddinFormName(ByVal strName As String) As Boolean
    IsAddinFormName = (StrComp(Left$(strName, 6), "frmVCS", vbTextCompare) = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddinFormIsLoaded
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True when Forms contains this add-in form.
'---------------------------------------------------------------------------------------
'
Private Function AddinFormIsLoaded(ByVal strName As String) As Boolean

    Dim frm As Form

    On Error Resume Next
    For Each frm In Forms
        If StrComp(frm.Name, strName, vbTextCompare) = 0 Then
            AddinFormIsLoaded = True
            Exit Function
        End If
    Next frm

End Function


'---------------------------------------------------------------------------------------
' Procedure : OperationIsActive
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True while an export, build, merge, or test run has not finished.
'---------------------------------------------------------------------------------------
'
Private Function OperationIsActive() As Boolean
    OperationIsActive = (Operation.Status = eosRunning Or Operation.Status = eosStaged)
End Function


'---------------------------------------------------------------------------------------
' Procedure : OperationStatusName
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Status string for dialog-list JSON.
'---------------------------------------------------------------------------------------
'
Private Function OperationStatusName() As String

    Select Case Operation.Status
        Case eosRunning: OperationStatusName = "running"
        Case eosStaged: OperationStatusName = "staged"
        Case Else: OperationStatusName = "ready"
    End Select

End Function

Attribute VB_Name = "modDialogInspect"
'---------------------------------------------------------------------------------------
' Module    : modDialogInspect
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : List and dismiss add-in forms that are already open. The public API is
'           : exposed through VCS.ListAddinDialogs and VCS.DismissAddinDialog.
'           : These calls need Access to be responsive. A modal MsgBox or the conflict
'           : dialog blocks them; the MCP dialog inspector covers those.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Utility")


Private Const ADDIN_FORM_PREFIX As String = "frmVCS"

Private Const ACTION_CLOSE As String = "close"
Private Const ACTION_CANCEL As String = "cancel"

Private Const ERR_NOT_ADDIN_FORM As String = "not_addin_form"
Private Const ERR_UNSUPPORTED_ACTION As String = "unsupported_action"
Private Const ERR_OPERATION_IN_PROGRESS As String = "operation_in_progress"
Private Const ERR_NOT_OPEN As String = "not_open"
Private Const ERR_NO_RUNNING_OPERATION As String = "no_running_operation"
Private Const ERR_OPERATION_ID_REQUIRED As String = "operation_id_required"
Private Const ERR_OPERATION_ID_MISMATCH As String = "operation_id_mismatch"
Private Const ERR_OPERATION_NOT_RUNNING As String = "operation_not_running"


'---------------------------------------------------------------------------------------
' Procedure : ListAddinDialogs
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Enumerate open add-in forms as JSON.
'---------------------------------------------------------------------------------------
'
Public Function ListAddinDialogs() As String

    Dim frm As Form
    Dim dRoot As Dictionary
    Dim colItems As Collection
    Dim dItem As Dictionary

    Set dRoot = New Dictionary
    Set colItems = New Collection

    ' Some properties are unavailable on a form that is closing.
    InitializeRequestErrors "ListAddinDialogs"
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
    Err.Clear
    On Error GoTo 0

    dRoot.Add "success", True
    dRoot.Add "dialogs", colItems
    dRoot.Add "operation_status", Operation.StatusName()
    dRoot.Add "interaction", IIf(InteractionIsNonInteractive(), "noninteractive", "interactive")
    AddOperationIdentity dRoot, Operation
    dRoot("error_level") = Operation.ErrorLevel
    ListAddinDialogs = ConvertToJson(dRoot)

End Function


'---------------------------------------------------------------------------------------
' Procedure : DismissAddinDialog
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Close a finished add-in form, or request cancellation of a running
'           : operation. Cancel goes through Operation.RequestCancel and does not
'           : call Finish; the running operation stops at its next check, and its own
'           : result reports whether it was cancelled. A call that came through the
'           : API dispatcher must name the running operation by its public ID (the
'           : MCP operation ID); trusted VBA in this session may omit it. A refused
'           : cancel changes nothing.
'---------------------------------------------------------------------------------------
'
Public Function DismissAddinDialog(ByVal strName As String, Optional ByVal strAction As String = ACTION_CLOSE, _
    Optional ByVal strOperationId As String = vbNullString) As String

    Dim strAct As String
    Dim blnRunning As Boolean
    Dim blnClosed As Boolean
    Dim dRoot As Dictionary

    strAct = LCase$(Trim$(strAction))
    If Len(strAct) = 0 Then strAct = ACTION_CLOSE

    Set dRoot = New Dictionary
    dRoot.Add "name", strName
    dRoot.Add "action", strAct

    If Not IsAddinFormName(strName) Then
        DismissAddinDialog = DismissRefusal(dRoot, ERR_NOT_ADDIN_FORM, _
            T("Only add-in forms (frmVCS*) can be dismissed here."))
        Exit Function
    End If

    If strAct <> ACTION_CLOSE And strAct <> ACTION_CANCEL Then
        DismissAddinDialog = DismissRefusal(dRoot, ERR_UNSUPPORTED_ACTION, _
            T("Supported actions are close and cancel. Button clicks on VBA or Access dialogs use the MCP dialog inspector."))
        Exit Function
    End If

    blnRunning = Operation.IsActive()

    ' A pending timer request is work about to start: close waits for it too.
    If strAct = ACTION_CLOSE And (blnRunning Or TimerIsPending) Then
        DismissAddinDialog = DismissRefusal(dRoot, ERR_OPERATION_IN_PROGRESS, _
            T("This window belongs to a running operation. Close is only for a finished results window. Pass action=cancel to interrupt the operation."), True)
        Exit Function
    End If

    If strAct = ACTION_CANCEL Then
        DismissAddinDialog = RequestOperationCancel(dRoot, strOperationId)
        Exit Function
    End If

    If Not AddinFormIsLoaded(strName) Then
        DismissAddinDialog = DismissRefusal(dRoot, ERR_NOT_OPEN, T("That add-in form is not open."), True)
        Exit Function
    End If

    LogUnhandledErrors
    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    blnClosed = (Err.Number = 0) And Not AddinFormIsLoaded(strName)
    If Err.Number <> 0 Then dRoot.Add "error", Err.Description
    CatchAny eelNoError, vbNullString, "modDialogInspect.DismissAddinDialog"
    On Error GoTo 0

    dRoot.Add "success", blnClosed
    dRoot.Add "closed", blnClosed
    dRoot.Add "interrupted", False
    If blnClosed Then dRoot.Add "message", T("Closed a finished add-in window.")
    DismissAddinDialog = ConvertToJson(dRoot)

End Function


'---------------------------------------------------------------------------------------
' Procedure : DismissRefusal
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Failure result for a dismiss request that changed nothing.
'---------------------------------------------------------------------------------------
'
Private Function DismissRefusal(ByVal dRoot As Dictionary, ByVal strPattern As String, _
    ByVal strMessage As String, Optional ByVal blnIncludeClosed As Boolean = False) As String

    dRoot.Add "success", False
    dRoot.Add "error_pattern", strPattern
    dRoot.Add "error", strMessage
    dRoot.Add "interrupted", False
    If blnIncludeClosed Then dRoot.Add "closed", False
    DismissRefusal = ConvertToJson(dRoot)

End Function


'---------------------------------------------------------------------------------------
' Procedure : RequestOperationCancel
' Author    : Josh
' Date      : 10/19/2026
' Purpose   : The cancel action. Validates the request against the running operation,
'           : then hands it to Operation.RequestCancel. Acceptance is only a request:
'           : nothing is finished or closed here.
'---------------------------------------------------------------------------------------
'
Private Function RequestOperationCancel(ByVal dRoot As Dictionary, ByVal strOperationId As String) As String

    Dim strActiveId As String
    Dim strPattern As String

    strActiveId = ActiveOperationPublicId()
    strPattern = CancelRequestRefusal(strOperationId, strActiveId, Operation.IsActive(), APIDispatchBusy())
    If Len(strPattern) = 0 Then
        ' A staged or paused root cannot take a request now.
        If Not Operation.RequestCancel("modDialogInspect.DismissAddinDialog") Then strPattern = ERR_OPERATION_NOT_RUNNING
    End If

    If Len(strPattern) > 0 Then
        dRoot.Add "requested", False
        RequestOperationCancel = DismissRefusal(dRoot, strPattern, CancelRefusalMessage(strPattern), True)
        Exit Function
    End If

    dRoot.Add "success", True
    dRoot.Add "requested", True
    dRoot.Add "interrupted", True
    dRoot.Add "closed", False
    If Len(strActiveId) > 0 Then dRoot.Add "operation_id", strActiveId
    dRoot.Add "message", T("Cancellation requested. This is only a request: the operation stops at its next check, and its own result reports whether it was cancelled.")
    RequestOperationCancel = ConvertToJson(dRoot)

End Function


'---------------------------------------------------------------------------------------
' Procedure : CancelRequestRefusal
' Author    : Josh
' Date      : 10/19/2026
' Purpose   : The refusal pattern for a cancel request, or an empty string when it may
'           : go ahead. With no active operation there is nothing to cancel. An ID,
'           : when given, must be the active operation's public ID; an operation with
'           : no public ID matches none, so a stale or foreign ID is refused. The ID
'           : may be omitted only when blnIdRequired is False (trusted VBA). Pure, so
'           : the table can be tested without a running operation.
'---------------------------------------------------------------------------------------
'
Public Function CancelRequestRefusal(ByVal strOperationId As String, ByVal strActiveId As String, _
    ByVal blnActive As Boolean, ByVal blnIdRequired As Boolean) As String

    If Not blnActive Then
        CancelRequestRefusal = ERR_NO_RUNNING_OPERATION
    ElseIf Len(strOperationId) = 0 Then
        If blnIdRequired Then CancelRequestRefusal = ERR_OPERATION_ID_REQUIRED
    ElseIf Len(strActiveId) = 0 Then
        CancelRequestRefusal = ERR_OPERATION_ID_MISMATCH
    ElseIf StrComp(strOperationId, strActiveId, vbBinaryCompare) <> 0 Then
        CancelRequestRefusal = ERR_OPERATION_ID_MISMATCH
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ActiveOperationPublicId
' Author    : Josh
' Date      : 10/19/2026
' Purpose   : The public ID of the session's active operation, as stored on the
'           : operation by its request. An MCP operation that has already posted its
'           : outcome has no active ID, so a late cancel cannot name it.
'---------------------------------------------------------------------------------------
'
Private Function ActiveOperationPublicId() As String
    If Not Operation.IsActive() Then Exit Function
    If Operation.Source = eosMCPTool And MCP.IsActive Then
        If MCP.TerminalPosted Then Exit Function
    End If
    ActiveOperationPublicId = Operation.PublicId
End Function


'---------------------------------------------------------------------------------------
' Procedure : CancelRefusalMessage
' Author    : Josh
' Date      : 10/19/2026
' Purpose   : The message for a refused cancel request.
'---------------------------------------------------------------------------------------
'
Private Function CancelRefusalMessage(ByVal strPattern As String) As String
    Select Case strPattern
        Case ERR_NO_RUNNING_OPERATION
            CancelRefusalMessage = T("No operation is running, so there is nothing to cancel. Nothing was changed.")
        Case ERR_OPERATION_ID_REQUIRED
            CancelRefusalMessage = T("Pass the running operation's ID to cancel it through the API. Nothing was changed.")
        Case ERR_OPERATION_ID_MISMATCH
            CancelRefusalMessage = T("That ID does not name the running operation. Nothing was changed.")
        Case ERR_OPERATION_NOT_RUNNING
            CancelRefusalMessage = T("The operation is paused or between steps and cannot take a request now. Nothing was changed.")
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
    IsAddinFormName = (StrComp(Left$(strName, Len(ADDIN_FORM_PREFIX)), ADDIN_FORM_PREFIX, vbTextCompare) = 0)
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

    InitializeRequestErrors "AddinFormIsLoaded"
    On Error Resume Next
    For Each frm In Forms
        If StrComp(frm.Name, strName, vbTextCompare) = 0 Then
            AddinFormIsLoaded = True
            Exit For
        End If
    Next frm
    CatchAny eelNoError, vbNullString, "modDialogInspect.AddinFormIsLoaded"
    On Error GoTo 0

End Function

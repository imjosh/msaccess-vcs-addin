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
    LogUnhandledErrors
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
    CatchAny eelNoError, vbNullString, "modDialogInspect.ListAddinDialogs"
    On Error GoTo 0

    dRoot.Add "success", True
    dRoot.Add "dialogs", colItems
    dRoot.Add "operation_status", Operation.StatusName()
    dRoot.Add "interaction", IIf(InteractionIsNonInteractive(), "noninteractive", "interactive")
    ListAddinDialogs = ConvertToJson(dRoot)

End Function


'---------------------------------------------------------------------------------------
' Procedure : DismissAddinDialog
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Close a finished add-in form, or request cancellation of a running
'           : operation. Cancel sets the cancel flag or critical error level and does
'           : not call Finish; the running operation stops at its next check.
'---------------------------------------------------------------------------------------
'
Public Function DismissAddinDialog(ByVal strName As String, Optional ByVal strAction As String = ACTION_CLOSE) As String

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

    If strAct = ACTION_CLOSE And blnRunning Then
        DismissAddinDialog = DismissRefusal(dRoot, ERR_OPERATION_IN_PROGRESS, _
            T("This window belongs to a running operation. Close is only for a finished results window. Pass action=cancel to interrupt the operation."), True)
        Exit Function
    End If

    If strAct = ACTION_CANCEL And blnRunning Then
        RequestCancel
        dRoot.Add "success", True
        dRoot.Add "interrupted", True
        dRoot.Add "closed", False
        dRoot.Add "message", T("Cancellation requested. The operation stops at its next check and Finish reports cancelled.")
        DismissAddinDialog = ConvertToJson(dRoot)
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
' Procedure : RequestCancel
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Ask the running operation to stop at its next check.
'---------------------------------------------------------------------------------------
'
Private Sub RequestCancel()

    If TestRunner.State = etrsRunning Then
        TestRunner.Cancel
    Else
        Operation.ErrorLevel = eelCritical
        Operation.Result = eorCanceled
    End If

End Sub


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

    LogUnhandledErrors
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

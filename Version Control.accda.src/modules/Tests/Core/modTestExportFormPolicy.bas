Attribute VB_Name = "modTestExportFormPolicy"
'---------------------------------------------------------------------------------------
' Module    : modTestExportFormPolicy
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Regression checks that Export, FullExport and ExportVBA keep the main form
'           : hidden under a session decision policy, and still show it interactively.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


'---------------------------------------------------------------------------------------
' Procedure : ExportFormPolicyCheck
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Needs a free session operation, so it cannot run inside the runner's root.
'           : Run it on the development copy, against a database that is not the add-in
'           : (the export entry points then stop at the running-database refusal, after
'           : the form has been shown or hidden):
'           : vcs_run_vba(<dev copy>, "MCP_TempFunction = ExportFormPolicyCheck()").
'           : Returns "OK", or the failed checks one per line. A Function, so the
'           : runner never lists it as a test.
'---------------------------------------------------------------------------------------
'
Public Function ExportFormPolicyCheck() As String

    Dim strFailed As String
    Dim intEntry As Long
    Dim intStart As Long
    Dim blnStartVisible As Boolean
    Dim strCase As String
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy

    If Operation.Status = eosRunning Then
        ExportFormPolicyCheck = "Not run: an operation is already running."
        Exit Function
    End If

    VCS.SetOperationPolicy "block"
    intMode = Operation.InteractionMode
    intPolicy = Operation.DecisionPolicy
    Check strFailed, intMode = eimNonInteractive, "the session policy is noninteractive"

    For intEntry = 1 To 3
        For intStart = 0 To 1
            blnStartVisible = (intStart = 1)
            strCase = EntryName(intEntry) & IIf(blnStartVisible, " (was visible)", " (was hidden)")
            CloseMainForm
            DoCmd.OpenForm "frmVCSMain", , , , , IIf(blnStartVisible, acNormal, acHidden)
            Check strFailed, Form_frmVCSMain.Visible = blnStartVisible, strCase & ": the form starts as expected"
            RunEntry intEntry
            Check strFailed, IsLoaded(acForm, "frmVCSMain", False), strCase & ": the form is still loaded"
            If IsLoaded(acForm, "frmVCSMain", False) Then
                Check strFailed, Not Form_frmVCSMain.Visible, strCase & ": the form stays hidden"
            End If
            Check strFailed, Operation.Status <> eosRunning, strCase & ": the operation finished"
            Check strFailed, Operation.InteractionMode = intMode, strCase & ": the session mode is kept"
            Check strFailed, Operation.DecisionPolicy = intPolicy, strCase & ": the session policy is kept"
        Next intStart
    Next intEntry
    CloseMainForm
    VCS.ClearOperationPolicy

    ' Interactive control: the same helper shows the form.
    Check strFailed, Operation.InteractionMode = eimNormal, "the policy is cleared"
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    SetMainFormVisible Form_frmVCSMain
    Check strFailed, Form_frmVCSMain.Visible, "an interactive export shows the form"
    CloseMainForm

    If Len(strFailed) = 0 Then strFailed = "OK"
    ExportFormPolicyCheck = strFailed

End Function


Private Sub RunEntry(intEntry As Long)
    Select Case intEntry
        Case 1: VCS.Export
        Case 2: VCS.FullExport
        Case 3: VCS.ExportVBA
    End Select
End Sub


Private Function EntryName(intEntry As Long) As String
    EntryName = Choose(intEntry, "Export", "FullExport", "ExportVBA")
End Function


Private Sub CloseMainForm()
    If IsLoaded(acForm, "frmVCSMain", False) Then DoCmd.Close acForm, "frmVCSMain"
End Sub


Private Sub Check(ByRef strFailed As String, blnPassed As Boolean, strCheck As String)
    If Not blnPassed Then strFailed = strFailed & "FAIL: " & strCheck & vbCrLf
End Sub

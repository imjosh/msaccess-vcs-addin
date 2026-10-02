Attribute VB_Name = "modTestModeOnlyScope"
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")

' Parameterized public probes run in the installed driver, outside the runner's
' root. Mutation callers must use a disposable database and export folder.
Public Function ModeOnlyScopeProbe(ByVal strAction As String) As String
    Dim dResult As Dictionary
    Dim cVcs As clsVersionControl
    Dim cRoot As clsRootOperationLease
    Dim cRefused As clsRootOperationLease
    Dim lngErr As Long
    Dim strErr As String
    On Error GoTo ErrHandler
    Set dResult = New Dictionary
    Select Case strAction
        Case "snapshot"
            Set dResult = ModeSnapshot()
        Case "fault"
            Set cVcs = New clsVersionControl
            cVcs.FaultAfterBegin = 5
            dResult.Add "outcome", ParseJson(cVcs.ExportByType("queries"))
            dResult.Add "after", ModeSnapshot()
        Case "refused"
            Set cRoot = Operation.TryBeginRoot(eotExport)
            Operation.ResolvePrompt vbYesNo, "A34 parent", "Unanswered parent prompt"
            dResult.Add "before", ModeSnapshot()
            Set cRefused = Operation.TryBeginRoot(eotMerge, edpDecline)
            dResult.Add "acquired", Not cRefused Is Nothing
            dResult.Add "invalid_mode", ParseJson(VCS.SetInteractionMode(99))
            dResult.Add "weaker_mode", ParseJson(VCS.SetInteractionMode(eimNormal))
            dResult.Add "invalid_policy", ParseJson(VCS.SetOperationPolicy("invalid-a34"))
            dResult.Add "after", ModeSnapshot()
            cRoot.Complete eorSuccess
            dResult.Add "completed", ModeSnapshot()
        Case Else
            dResult.Add "error", "Unknown probe action"
    End Select
CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If lngErr <> 0 Then
        dResult.Add "probe_error", strErr
        dResult.Add "probe_error_number", lngErr
    End If
    ModeOnlyScopeProbe = ConvertToJson(dResult)
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function

Private Function ModeSnapshot() As Dictionary
    Dim dResult As Dictionary
    Set dResult = New Dictionary
    dResult.Add "mode", Operation.InteractionMode
    dResult.Add "policy", Operation.DecisionPolicy
    dResult.Add "blocked", Operation.DecisionBlocked
    dResult.Add "decisions", Operation.Decisions
    dResult.Add "error_level", Operation.ErrorLevel
    dResult.Add "status", Operation.StatusName
    dResult.Add "token", Operation.CurrentRootToken
    dResult.Add "result", Operation.Result
    dResult.Add "force_unattended", Operation.ForceUnattended
    dResult.Add "error_trapping", Application.GetOption("Error Trapping")
    dResult.Add "form_loaded", IsLoaded(acForm, "frmVCSMain", False)
    If IsLoaded(acForm, "frmVCSMain", False) Then
        dResult.Add "form_visible", Form_frmVCSMain.Visible
    End If
    Set ModeSnapshot = dResult
End Function

Public Sub TestModeOnlyRootDiscardsItsJournal()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    On Error GoTo ErrHandler
    Set cOp = New clsOperation
    SelectInteractionMode cOp, eimNonInteractive
    Set cRoot = cOp.TryBeginRoot(eotExport)
    cOp.ResolvePrompt vbYesNo, "A34", "Proceed?"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorDecisionRequired, "blocked operation reports decision_required"
    TestAssert cOp.LastDecisions.Count = 1, "completed journal is retained in the result"
    TestAssert Not cOp.DecisionBlocked, "blocked state is released"
    TestAssert cOp.Decisions.Count = 0, "live journal is restored"
    TestAssert cOp.InteractionMode = eimNormal, "mode-only completion returns to normal"
    Set cRoot = cOp.TryBeginRoot(eotExport)
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "next interactive operation succeeds"
    TestAssert cOp.LastDecisions.Count = 0, "next outcome has no stale journal"
    Exit Sub
ErrHandler:
    TestAssert False, "mode-only root: " & Err.Description
End Sub

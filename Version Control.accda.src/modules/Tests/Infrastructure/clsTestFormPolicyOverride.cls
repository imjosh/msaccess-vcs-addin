VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestFormPolicyOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestFormPolicyOverride
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Tests for the guard MergeBuild uses to set the form's decision policy. The
'           : raising case matters because cmdBuild_Click can raise, and a policy left on
'           : the form would leak into the next, unrelated Build or MergeBuild call.
'           : The tests use the hidden main form but never run a build, and each one
'           : puts the form's policy back to the value it found.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Private m_frmSaved As Form_frmVCSMain
Private m_intSavedPolicy As eDecisionPolicy
Private m_blnPolicySaved As Boolean


' The form is left loaded and hidden, as Build and MergeBuild leave it. Closing it
' would be refused while the run hosting these tests holds a root.
Private Function MainForm() As Form_frmVCSMain
    If Not IsLoaded(acForm, "frmVCSMain", False) Then
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    End If
    Set MainForm = Form_frmVCSMain
    If Not m_blnPolicySaved Then
        Set m_frmSaved = Form_frmVCSMain
        m_intSavedPolicy = m_frmSaved.intDecisionPolicy
        m_blnPolicySaved = True
    End If
End Function


' Applies an override, then raises (division by zero) with the guard still in scope.
Private Sub RaiseWhileOverridden(ByVal intPolicy As eDecisionPolicy)

    Dim cOverride As clsFormPolicyOverride
    Dim lngZero As Long

    Set cOverride = New clsFormPolicyOverride
    cOverride.Init MainForm, intPolicy
    lngZero = 1 \ lngZero

End Sub


Public Sub TestOverrideAppliesThenRestoresTheFoundValue()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim frm As Form_frmVCSMain
    Dim intFound As eDecisionPolicy
    Dim cOverride As clsFormPolicyOverride

    Set frm = MainForm
    intFound = frm.intDecisionPolicy
    frm.intDecisionPolicy = edpDecline
    Set cOverride = New clsFormPolicyOverride
    cOverride.Init frm, edpBlock
    TestAssert frm.intDecisionPolicy = edpBlock, "the override is in force"
    cOverride.Restore
    TestAssert frm.intDecisionPolicy = edpDecline, "restore puts back the value it found, not a default"
    frm.intDecisionPolicy = edpPreferSource
    cOverride.Restore
    TestAssert frm.intDecisionPolicy = edpPreferSource, "a second restore does nothing"
    frm.intDecisionPolicy = intFound

CleanUp:
    On Error Resume Next
    Set cOverride = Nothing
    RestoreSavedFormPolicy
    If Err.Number <> 0 Then TestAssert False, "TestOverrideAppliesThenRestoresTheFoundValue cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestOverrideAppliesThenRestoresTheFoundValue " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestOverrideRestoredWhenTheGuardIsReleased()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    Dim frm As Form_frmVCSMain
    Dim intFound As eDecisionPolicy
    Dim cOverride As clsFormPolicyOverride

    Set frm = MainForm
    intFound = frm.intDecisionPolicy
    frm.intDecisionPolicy = edpAsk
    Set cOverride = New clsFormPolicyOverride
    cOverride.Init frm, edpSkip
    TestAssert frm.intDecisionPolicy = edpSkip, "the override is in force"
    Set cOverride = Nothing
    TestAssert frm.intDecisionPolicy = edpAsk, "releasing the guard restores the policy"
    frm.intDecisionPolicy = intFound

CleanUp:
    On Error Resume Next
    Set cOverride = Nothing
    RestoreSavedFormPolicy
    If Err.Number <> 0 Then TestAssert False, "TestOverrideRestoredWhenTheGuardIsReleased cleanup: " & Err.Number & ": " & Err.Description
    If lngTestErr <> 0 Then TestAssert False, "TestOverrideRestoredWhenTheGuardIsReleased " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestOverrideRestoredWhenTheGuardedCallRaises()
    On Error GoTo ErrHandler
    Dim frm As Form_frmVCSMain
    Dim lngExpectedErr As Long
    Dim lngTestErr As Long
    Dim strTestErr As String
    Set frm = MainForm
    frm.intDecisionPolicy = edpAsk

    On Error GoTo ExpectedError
    RaiseWhileOverridden edpPreferSource
    TestAssert False, "the guarded call should have raised"
    GoTo CleanUp
ExpectedError:
    lngExpectedErr = Err.Number
    Resume CheckExpectedError
CheckExpectedError:
    On Error GoTo ErrHandler
    TestAssert lngExpectedErr = 11, "the error reached the caller"
    TestAssert frm.intDecisionPolicy = edpAsk, "the policy is back to what it was after the error"
CleanUp:
    On Error Resume Next
    RestoreSavedFormPolicy
    If lngTestErr <> 0 Then TestAssert False, "TestOverrideRestoredWhenTheGuardedCallRaises " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub

Private Sub RestoreSavedFormPolicy()
    On Error Resume Next
    If m_blnPolicySaved Then
        m_frmSaved.intDecisionPolicy = m_intSavedPolicy
        If Err.Number <> 0 Then TestAssert False, "form policy cleanup: " & Err.Number & ": " & Err.Description
    End If
    m_blnPolicySaved = False
    Set m_frmSaved = Nothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    RestoreSavedFormPolicy
End Sub

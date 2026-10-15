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


' The form is left loaded and hidden, as Build and MergeBuild leave it. Closing it
' would be refused while the run hosting these tests holds a root.
Private Function MainForm() As Form_frmVCSMain
    If Not IsLoaded(acForm, "frmVCSMain", False) Then
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    End If
    Set MainForm = Form_frmVCSMain
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

End Sub


Public Sub TestOverrideRestoredWhenTheGuardIsReleased()

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

End Sub


Public Sub TestOverrideRestoredWhenTheGuardedCallRaises()

    Dim frm As Form_frmVCSMain
    Dim intFound As eDecisionPolicy

    Set frm = MainForm
    intFound = frm.intDecisionPolicy
    frm.intDecisionPolicy = edpAsk

    LogUnhandledErrors
    On Error GoTo ErrHandler

    RaiseWhileOverridden edpPreferSource
    TestAssert False, "the guarded call should have raised"
    frm.intDecisionPolicy = intFound
    Exit Sub

ErrHandler:
    TestAssert Err.Number = 11, "the error reached the caller"
    TestAssert frm.intDecisionPolicy = edpAsk, "the policy is back to what it was after the error"
    frm.intDecisionPolicy = intFound

End Sub

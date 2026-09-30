VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormPolicyOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFormPolicyOverride
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Sets the decision policy a form passes to its next operation (the form's
'           : intDecisionPolicy field) and puts back the value it found on an explicit
'           : Restore or in Class_Terminate, so an error raised by the guarded call
'           : cannot leave the policy behind for the next, unrelated call.
'           : The form is early bound on purpose: a late-bound read of its field, which
'           : is typed with a private-module enum, fails with error 5.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Private m_frmTarget As Form_frmVCSMain
Private m_intSaved As eDecisionPolicy
Private m_blnActive As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Init
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Save the target's current policy and replace it with intPolicy.
'---------------------------------------------------------------------------------------
'
Public Sub Init(ByVal frmTarget As Form_frmVCSMain, ByVal intPolicy As eDecisionPolicy)
    Restore
    Set m_frmTarget = frmTarget
    m_intSaved = m_frmTarget.intDecisionPolicy
    m_frmTarget.intDecisionPolicy = intPolicy
    m_blnActive = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Restore
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Put the saved policy back exactly once.
'---------------------------------------------------------------------------------------
'
Public Sub Restore()
    If Not m_blnActive Then Exit Sub
    m_blnActive = False
    m_frmTarget.intDecisionPolicy = m_intSaved
    Set m_frmTarget = Nothing
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    Restore
End Sub

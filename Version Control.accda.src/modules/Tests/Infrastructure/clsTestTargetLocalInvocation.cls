VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestTargetLocalInvocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Private Sub Class_Terminate()
    TargetLocalTerminateCount = TargetLocalTerminateCount + 1
End Sub

' Keep a target-local form type in this class without opening a form.
Private Function TargetForm() As Form_frmVCSMain
    Set TargetForm = Nothing
End Function

Public Sub TestTargetLocalFormType()
    On Error GoTo Failed
    Dim frm As Form_frmVCSMain
    Set frm = TargetForm
    TestAssert frm Is Nothing, "target-local form type resolves without opening UI"
    TestAssert TypeName(Me) = "clsTestTargetLocalInvocation", "target-local class executes"
    Exit Sub
Failed:
    TestAssert False, "target-local invocation failed: " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub

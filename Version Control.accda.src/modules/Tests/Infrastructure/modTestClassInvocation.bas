Attribute VB_Name = "modTestClassInvocation"
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Public TargetLocalTerminateCount As Long

Public Sub TestClassInvokerReturnsScalarErrorAndReleasesInstance()
    On Error GoTo Failed
    Dim obj As Object
    Dim varResult As Variant
    Dim lngBefore As Long
    Dim lngErr As Long
    Dim strErr As String
    lngBefore = TargetLocalTerminateCount
    Set obj = New clsTestTargetLocalInvocation
    ' Missing-member dispatch is caught in the helper, without a failing test body.
    varResult = Application.Run(CurrentProject.Path & PathSep & FSO.GetBaseName(CurrentProject.Name) & ".TestClassInvoke", obj, "NoSuchTestMethod")
    TestAssert VarType(varResult) = vbString, "class errors cross Application.Run as a scalar string"
    TestAssert Left$(CStr(varResult), 4) = "438" & vbTab, "missing member returns error 438 and a delimiter"
    TestAssert Len(Mid$(CStr(varResult), 5)) > 0, "caught error includes its description"
CleanUp:
    On Error Resume Next
    Set obj = Nothing
    TestAssert TargetLocalTerminateCount = lngBefore + 1, "class teardown runs after a caught invocation error"
    If lngErr <> 0 Then TestAssert False, "invoker error " & lngErr & ": " & strErr
    Err.Clear
    Exit Sub
Failed:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub

Public Sub TestClassInvokerSuccessAndReleasesInstance()
    On Error GoTo Failed
    Dim obj As Object
    Dim varResult As Variant
    Dim lngBefore As Long
    Dim lngErr As Long
    Dim strErr As String
    lngBefore = TargetLocalTerminateCount
    Set obj = New clsTestTargetLocalInvocation
    varResult = Application.Run(CurrentProject.Path & PathSep & FSO.GetBaseName(CurrentProject.Name) & ".TestClassInvoke", obj, "TestTargetLocalFormType")
    TestAssert CStr(varResult) = vbNullString, "successful invocation returns no error"
CleanUp:
    On Error Resume Next
    Set obj = Nothing
    TestAssert TargetLocalTerminateCount = lngBefore + 1, "class teardown runs after successful invocation"
    If lngErr <> 0 Then TestAssert False, "invoker error " & lngErr & ": " & strErr
    Err.Clear
    Exit Sub
Failed:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub

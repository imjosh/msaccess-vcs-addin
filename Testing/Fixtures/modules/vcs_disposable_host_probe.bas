Attribute VB_Name = "vcs_disposable_host_probe"
Option Compare Database
Option Explicit

' Import into an idle disposable current project, never a loaded library.
Public Function DisposableHostSnapshot(ByVal ignored As String) As String
    Dim state As New Dictionary
    state.Add "error_level", CLng(Operation.ErrorLevel)
    state.Add "test_run_active", TestRunActive
    state.Add "error_trapping", Application.GetOption("Error Trapping")
    state.Add "terminated_instances", TargetLocalTerminateCount
    DisposableHostSnapshot = ConvertToJson(state)
End Function

' Designate this project's database as a disposable suite host. The import tests refuse
' to run in a host that lacks this property or whose property names another database.
' Call once in the idle host, never on a primary development copy. Returns the stored
' value so the caller can confirm it matches CurrentProject.FullName.
Public Function DesignateDisposableHost(ByVal ignored As String) As String
    SetDBProperty DISPOSABLE_HOST_PROPERTY, CurrentProject.FullName
    DesignateDisposableHost = CStr(GetDBProperty(DISPOSABLE_HOST_PROPERTY))
End Function

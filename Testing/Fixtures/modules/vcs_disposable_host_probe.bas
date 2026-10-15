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

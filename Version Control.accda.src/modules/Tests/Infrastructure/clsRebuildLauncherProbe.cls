VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRebuildLauncherProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")

' Disposable test fixture. This helper only waits briefly and exits. It cannot
' rebuild, install, quit Access, or authorize a real worker to proceed.
Private Type udtProbeFileTime
    LowDateTime As Long
    HighDateTime As Long
End Type
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetProcessTimes Lib "kernel32" (ByVal hProcess As LongPtr, ByRef lpCreationTime As udtProbeFileTime, ByRef lpExitTime As udtProbeFileTime, ByRef lpKernelTime As udtProbeFileTime, ByRef lpUserTime As udtProbeFileTime) As Long
Private Const PROCESS_SYNCHRONIZE As Long = &H100000
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const WAIT_OBJECT_0 As Long = 0
Private m_Handle As LongPtr
Public LaunchCount As Long
Public ProcessId As Long
Public CreationTime As String

Public Function Launch(ByVal strStatusFile As String) As Boolean
    Dim created As udtProbeFileTime
    Dim exited As udtProbeFileTime
    Dim kernel As udtProbeFileTime
    Dim user As udtProbeFileTime
    LaunchCount = LaunchCount + 1
    ProcessId = CLng(Shell("""" & Environ$("ComSpec") & """ /d /c ping -n 2 127.0.0.1 >nul", vbHide))
    m_Handle = OpenProcess(PROCESS_SYNCHRONIZE Or PROCESS_QUERY_LIMITED_INFORMATION, 0, ProcessId)
    If m_Handle = 0 Then Err.Raise vbObjectError + 2461, "clsRebuildLauncherProbe", T("Could not hold the harmless helper process.")
    If GetProcessTimes(m_Handle, created, exited, kernel, user) = 0 Then Err.Raise vbObjectError + 2462, "clsRebuildLauncherProbe", T("Could not identify the harmless helper process.")
    CreationTime = Hex$(created.HighDateTime) & ":" & Hex$(created.LowDateTime)
    WriteRebuildStatusFile strStatusFile, "fixture-helper-ready"
    Launch = True
End Function

Public Function ConfirmExit() As Boolean
    If m_Handle = 0 Then Exit Function
    ConfirmExit = (WaitForSingleObject(m_Handle, 5000) = WAIT_OBJECT_0)
End Function

Private Sub Class_Terminate()
    If m_Handle <> 0 Then CloseHandle m_Handle
End Sub

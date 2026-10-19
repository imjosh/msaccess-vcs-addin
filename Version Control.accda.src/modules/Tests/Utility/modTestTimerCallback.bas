Attribute VB_Name = "modTestTimerCallback"
'---------------------------------------------------------------------------------------
' Module    : modTestTimerCallback
' Purpose   : A49 timer event identity. Every test here calls WinAPITimerCallback
'           : directly with crafted TIMERPROC arguments (the direct VBA seam). None
'           : waits for Windows to deliver a WM_TIMER message, and none shows how
'           : Windows itself reuses timer IDs; a reused ID is modelled by replaying the
'           : armed ID before its interval is due.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")
'@Tag("unit")

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
' Blocks without pumping messages, so Windows cannot deliver the probe timer meanwhile.
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const WM_TIMER As Long = &H113
' Tests that wait for the armed interval use the short one; the others use the long
' one, so the real timer cannot fire while a message is pumped mid-test.
Private Const ARM_SECONDS As Single = 1
Private Const IDLE_SECONDS As Single = 60
' Long enough for the short interval to be due.
Private Const DUE_MS As Long = 1150

Private m_lngRuns As Long


' Dispatch target for the armed probe request. It takes an argument so the runner does
' not discover it as a test.
Public Function TimerCallbackTarget(ByVal strMode As String) As Boolean
    m_lngRuns = m_lngRuns + 1
    If strMode = "raise" Then Err.Raise 5, "TimerCallbackTarget", "A49 injected dispatch failure"
    TimerCallbackTarget = True
End Function


Public Sub StaleTimerEventsLeaveTheRequestArmed()
    Dim lngID As LongPtr
    On Error GoTo Failed
    If Not ArmProbe("count", IDLE_SECONDS) Then Exit Sub
    lngID = ArmedID

    WinAPITimerCallback 0, 0, lngID, GetTickCount
    WinAPITimerCallback 0, WM_TIMER, 0, GetTickCount
    WinAPITimerCallback 0, WM_TIMER, lngID + 1, GetTickCount
    ' The armed ID reused: an old timer's event arriving before this request is due.
    WinAPITimerCallback 0, WM_TIMER, lngID, GetTickCount

    TestAssert m_lngRuns = 0, "a stale event dispatched"
    TestAssert TimerIsPending, "a stale event consumed the request"
    TestAssert ArmedID = lngID, "a stale event killed the armed timer"
    TestAssert GetSetting(PROJECT_NAME, "Timer", "Operation") = ProbeCommand
    TestAssert GetSetting(PROJECT_NAME, "Timer", "Param1") = "count"
    DisarmProbe
    Exit Sub
Failed:
    ReportAndDisarm
End Sub


Public Sub ArmedEventDispatchesOnce()
    Dim lngID As LongPtr
    On Error GoTo Failed
    If Not ArmProbe("count", ARM_SECONDS) Then Exit Sub
    lngID = ArmedID
    Sleep DUE_MS

    WinAPITimerCallback 0, WM_TIMER, lngID, GetTickCount
    ' A second event for the same timer, already queued before it was killed.
    WinAPITimerCallback 0, WM_TIMER, lngID, GetTickCount

    TestAssert m_lngRuns = 1, "the armed request ran " & m_lngRuns & " times"
    TestAssert Not TimerIsPending
    TestAssert GetSetting(PROJECT_NAME, "Timer", "Operation") = vbNullString
    DisarmProbe
    Exit Sub
Failed:
    ReportAndDisarm
End Sub


Public Sub EarlierEventLeavesTheLaterRequest()
    Dim lngFirst As LongPtr
    Dim lngSecond As LongPtr
    On Error GoTo Failed
    If Not ArmProbe("count", ARM_SECONDS) Then Exit Sub
    lngFirst = ArmedID
    Sleep DUE_MS
    WinAPITimerCallback 0, WM_TIMER, lngFirst, GetTickCount
    TestAssert m_lngRuns = 1, "the first request did not run"

    ' Arm the next request at once, as a continuation does, then replay the first
    ' request's event. Whether or not Windows reused the ID, it must not be consumed.
    TestAssert TrySetTimer(ProbeCommand, "count", sngSeconds:=ARM_SECONDS), "the second request was not armed"
    lngSecond = ArmedID
    WinAPITimerCallback 0, WM_TIMER, lngFirst, GetTickCount
    TestAssert m_lngRuns = 1, "the first request's event ran the second request"
    TestAssert TimerIsPending, "the first request's event consumed the second request"

    Sleep DUE_MS
    WinAPITimerCallback 0, WM_TIMER, lngSecond, GetTickCount
    TestAssert m_lngRuns = 2, "the second request ran " & (m_lngRuns - 1) & " times"
    TestAssert Not TimerIsPending
    DisarmProbe
    Exit Sub
Failed:
    ReportAndDisarm
End Sub


Public Sub DispatchFailureStaysInsideTheCallback()
    Dim blnSaved As Boolean
    Dim blnOrigBreak As Boolean
    Dim eimPrior As eInteractionMode
    Dim eelPrior As eErrorLevel
    Dim lngJournal As Long
    On Error GoTo Failed
    blnOrigBreak = Options.BreakOnError
    eimPrior = Operation.InteractionMode
    eelPrior = Operation.ErrorLevel
    blnSaved = True
    Options.BreakOnError = False
    Operation.InteractionMode = eimSilent
    Log.Add "  [DispatchFailureStaysInsideTheCallback] The timer warning logged by this test is expected.", False
    If Not ArmProbe("raise", ARM_SECONDS) Then GoTo Restore
    Sleep DUE_MS
    lngJournal = Log.ErrorJournalCount
    ' An error escaping here would reach Failed and fail this test.
    WinAPITimerCallback 0, WM_TIMER, ArmedID, GetTickCount
    TestAssert m_lngRuns = 1, "the failing request ran " & m_lngRuns & " times"
    TestAssert Not TimerIsPending
    TestAssert Log.ErrorJournalCount = lngJournal + 1, "the failure was not recorded once"
    DisarmProbe
    GoTo Restore
Failed:
    ReportAndDisarm
Restore:
    If Not blnSaved Then Exit Sub
    Options.BreakOnError = blnOrigBreak
    Operation.InteractionMode = eimPrior
    ' The recorded warning was the expected outcome, so it should not color the run.
    Operation.ErrorLevel = eelPrior
End Sub


Private Function ArmProbe(ByVal strMode As String, ByVal sngSeconds As Single) As Boolean
    m_lngRuns = 0
    TestAssert Not TimerIsPending, "another timer request is pending"
    If TimerIsPending Then Exit Function
    ArmProbe = TrySetTimer(ProbeCommand, strMode, sngSeconds:=sngSeconds)
    TestAssert ArmProbe, "the probe timer was not armed"
End Function


Private Sub DisarmProbe()
    On Error Resume Next
    If TimerIsPending Then RunPendingTimerRequest
    m_lngRuns = 0
End Sub


Private Sub ReportAndDisarm()
    TestAssert False, "unexpected error " & Err.Number & ": " & Err.Description
    DisarmProbe
End Sub


Private Function ArmedID() As LongPtr
    ArmedID = CLngPtr(Val(GetSetting(PROJECT_NAME, "Timer", "TimerID", "0")))
End Function


' Project-qualified, as the test runner builds its own Application.Run commands.
Private Function ProbeCommand() As String
    If CurrentVBProject.Name = PROJECT_NAME Then
        ProbeCommand = CurrentProject.Path & PathSep & FSO.GetBaseName(CurrentProject.Name) & ".TimerCallbackTarget"
    Else
        ProbeCommand = "[" & CurrentVBProject.Name & "].TimerCallbackTarget"
    End If
End Function

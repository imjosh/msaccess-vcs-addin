Attribute VB_Name = "modTimer"
'---------------------------------------------------------------------------------------
' Module    : modTimer
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : API timer functions for callbacks
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

' Windows API calls for Timer functionality
Private Declare PtrSafe Function ApiSetTimer Lib "user32" Alias "SetTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function ApiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Private Const WM_TIMER As Long = &H113
' Allowance for GetTickCount granularity when judging whether an event is early.
Private Const TIMER_EARLY_MS As Long = 50
Private Const REG_TIMER_OP_TOKEN As String = "OpToken"
Private Const REG_TIMER_ARMED_TICK As String = "ArmedTick"
Private Const REG_TIMER_ARMED_INTERVAL As String = "ArmedInterval"

Private m_lngTimerID As LongPtr
Private m_lngArmedTick As Long
Private m_lngArmedInterval As Long
Private m_strRunningCommand As String


'---------------------------------------------------------------------------------------
' Procedure : WinAPITimerCallback
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : TIMERPROC for the armed request (HWND, UINT, UINT_PTR, DWORD).
'           : The event is checked before anything is consumed. KillTimer does not
'           : remove a WM_TIMER that is already queued, and a thread timer's ID can be
'           : handed out again once it is killed, so an old event can carry the ID of
'           : the request armed after it. The armed timer cannot fire before its
'           : interval has passed, so an event for the armed ID that arrives sooner is
'           : the old timer's and is ignored. Ignoring a real event costs nothing: the
'           : armed timer is periodic and fires again. dwTime is not used: Windows
'           : documents it only as the tick count, not as when the event was generated.
'           : An old event that arrives a full interval after the re-arm cannot be told
'           : apart, and an interval at or below TIMER_EARLY_MS gets no such check.
'           : An event for any other ID is ignored, never killed: by the time an old
'           : event arrives, its ID may belong to another component's timer.
'           : Nothing may raise out of a Windows callback, so a failure in the request
'           : is recorded here as a warning instead: the request's own operation reports
'           : its outcome, and an error level here would count against whatever runs next.
'---------------------------------------------------------------------------------------
'
Public Sub WinAPITimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, ByVal dwTime As Long)

    Dim lngErr As Long
    Dim strErr As String
    Dim strCommand As String

    On Error GoTo ErrHandler
    m_strRunningCommand = vbNullString
    If IsArmedEvent(uMsg, idEvent) Then RunPendingTimerRequest
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strCommand = m_strRunningCommand
    m_strRunningCommand = vbNullString
    On Error Resume Next
    Debug.Print "Timer callback error " & lngErr & ": " & strErr
    Log.Error eelWarning, "Timer request " & strCommand & " failed: Error " & lngErr & " - " & strErr, _
        "modTimer.WinAPITimerCallback"
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsArmedEvent
' Purpose   : True only for a timer message for the armed ID, once its interval is due.
'---------------------------------------------------------------------------------------
'
Private Function IsArmedEvent(ByVal uMsg As Long, ByVal idEvent As LongPtr) As Boolean
    If uMsg <> WM_TIMER Or idEvent = 0 Then Exit Function
    If idEvent <> ArmedTimerID Then Exit Function
    IsArmedEvent = ArmedIntervalElapsed
End Function


'---------------------------------------------------------------------------------------
' Procedure : ArmedTimerID, ArmedIntervalElapsed
' Purpose   : The armed request's identity and due time. Module state does not survive
'           : a VBA project reset (MergeReset), so the saved copy stands in after one.
'           : A request saved without its arm tick (armed by older code) is taken as due.
'---------------------------------------------------------------------------------------
'
Private Function ArmedTimerID() As LongPtr
    If m_lngTimerID <> 0 Then
        ArmedTimerID = m_lngTimerID
    Else
        ArmedTimerID = GetSetting(PROJECT_NAME, "Timer", "TimerID", 0)
    End If
End Function

Private Function ArmedIntervalElapsed() As Boolean
    Dim strTick As String
    Dim lngTick As Long
    Dim lngInterval As Long
    If m_lngTimerID <> 0 Then
        lngTick = m_lngArmedTick
        lngInterval = m_lngArmedInterval
    Else
        strTick = GetSetting(PROJECT_NAME, "Timer", REG_TIMER_ARMED_TICK)
        If Len(strTick) = 0 Then
            ArmedIntervalElapsed = True
            Exit Function
        End If
        lngTick = CLng(Val(strTick))
        lngInterval = CLng(Val(GetSetting(PROJECT_NAME, "Timer", REG_TIMER_ARMED_INTERVAL, "0")))
    End If
    ArmedIntervalElapsed = (TickDifference(GetTickCount, lngTick) >= lngInterval - TIMER_EARLY_MS)
End Function


'---------------------------------------------------------------------------------------
' Procedure : TickDifference
' Purpose   : Signed milliseconds from one GetTickCount value to another, across the
'           : 49.7-day wrap. Both are unsigned DWORDs read into a signed Long.
'---------------------------------------------------------------------------------------
'
Private Function TickDifference(ByVal lngLater As Long, ByVal lngEarlier As Long) As Double
    Dim dblDiff As Double
    dblDiff = UnsignedTick(lngLater) - UnsignedTick(lngEarlier)
    If dblDiff >= 2147483648# Then dblDiff = dblDiff - 4294967296#
    If dblDiff < -2147483648# Then dblDiff = dblDiff + 4294967296#
    TickDifference = dblDiff
End Function

Private Function UnsignedTick(ByVal lngTick As Long) As Double
    UnsignedTick = lngTick
    If lngTick < 0 Then UnsignedTick = UnsignedTick + 4294967296#
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunPendingTimerRequest
' Purpose   : Consume and run the armed request. Called for a checked timer event, and
'           : directly by tests that dispatch without waiting for message delivery.
'---------------------------------------------------------------------------------------
'
Public Sub RunPendingTimerRequest()

    Dim strParam1 As String
    Dim strParam2 As String
    Dim strCommand As String
    Dim strOpToken As String
    Dim intMergeFilter As eContainerFilter

    ' Nothing armed, nothing to consume.
    If ArmedTimerID = 0 Then Exit Sub

    ' First, make sure we kill the timer!
    KillTimer

    ' Read in parameter values
    strCommand = GetSetting(PROJECT_NAME, "Timer", "Operation")
    m_strRunningCommand = strCommand
    strParam1 = GetSetting(PROJECT_NAME, "Timer", "Param1")
    strParam2 = GetSetting(PROJECT_NAME, "Timer", "Param2")

    ' Read callback info before clearing (needed for APIAsyncOperation)
    Dim strCallbackInfo As String
    strCallbackInfo = GetSetting(PROJECT_NAME, "Timer", "CallbackInfo")
    Dim strAdmission As String
    strAdmission = GetSetting(PROJECT_NAME, "Timer", "CompatibilityEnvelope")
    strOpToken = GetSetting(PROJECT_NAME, "Timer", REG_TIMER_OP_TOKEN)

    ' Clear values from registry (In case an operation sets another timer)
    ClearTimerRequest

    ' Now, run the desired operation. Root ownership crosses this boundary in strOpToken,
    ' never in mutable global state: a continuation resumes only the root it was armed
    ' for, and a stale or foreign token resumes nothing.
    Select Case strCommand

        Case "HandleRibbonCommand"
            HandleRibbonCommand strParam1

        Case "Build"
            ' Build from source (full or merge build)
            RunBuildFromContinuation strOpToken, strParam1, CBool(strParam2)

        Case "MergeReset"
            ' Reset the target database's VBA project between the two merge stages, on the
            ' smallest stack available: the call that prepared the database has fully
            ' unwound, and the merge has not started.
            '
            ' The next stage is armed BEFORE the reset on purpose. The reset's teardown
            ' lands asynchronously, so this stack must do nothing afterwards and simply
            ' return to the message loop. (See modBuild.ResetProjectForInPlaceMerge.)
            '
            ' The root is normally already staged by the call that armed this timer. Stage
            ' it here only if it is somehow still running, so the reset's asynchronous
            ' teardown is not mistaken for a canceled operation.
            If Operation.Status = eosRunning Then Operation.DetachRootLease strOpToken
            TraceInPlaceMerge "reset stage: merge timer armed"
            SetTimer "MergeResume", strParam1, strParam2, strOpToken
            ResetProjectForInPlaceMerge

        Case "MergeResume"
            ' Continue a merge build after the database was prepared in place and its VBA
            ' project reset. (See modBuild.PrepareMergeInPlace.)
            intMergeFilter = Val(strParam2)
            RunBuildFromContinuation strOpToken, strParam1, False, intMergeFilter, vbNullString, True

        Case "APIAsyncOperation"
            ' Handle async operation with MCP callbacks
            HandleAPIAsyncOperation strParam1, strParam2, strCallbackInfo, strAdmission

        Case "QuitForRebuild"
            ' Close this instance so the rebuild worker can replace the files it holds.
            ' Armed by clsVersionControl.RebuildAddIn before the worker is launched,
            ' because the worker cannot close an instance it could not attach to, and it
            ' cannot attach when the current database is the add-in itself.
            ' (See clsWorker.Main.)
            Application.Quit acQuitSaveAll

        Case Else
            ' Use the Run command to execute the specified operation with supplied parameters
            If strParam2 <> vbNullString Then
                Application.Run strCommand, strParam1, strParam2
            ElseIf strParam1 <> vbNullString Then
                Application.Run strCommand, strParam1
            Else
                Application.Run strCommand
            End If

    End Select
    m_strRunningCommand = vbNullString

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetTimer
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Set the API timer to trigger the desired operation.
'           : strOpToken carries root ownership to the continuation; pass the token from
'           : the lease that was detached for this handoff. It defaults to the current
'           : root so a timer armed inside an operation resumes that same operation.
'---------------------------------------------------------------------------------------
'
Public Sub SetTimer(strOperation As String, _
    Optional strParam1 As String, Optional strParam2 As String, _
    Optional strOpToken As String, Optional sngSeconds As Single = 0.5)

    If Not TrySetTimer(strOperation, strParam1, strParam2, strOpToken, sngSeconds) Then
        MsgBox2 T("Failed to Set Callback Timer"), _
            T("Multiple callback timers are not currently supported."), _
            T("Please ensure that any previous timer was completed or killed first."), vbExclamation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TimerIsPending
' Purpose   : Read admission state without changing the pending request.
'---------------------------------------------------------------------------------------
Public Function TimerIsPending() As Boolean
    TimerIsPending = (m_lngTimerID <> 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : TrySetTimer
' Purpose   : Admit the whole timer request, including its callback, in one place.
'           : A refused request writes nothing. APIAsync handles refusal without a UI.
'---------------------------------------------------------------------------------------
Public Function TrySetTimer(ByVal strOperation As String, _
    Optional ByVal strParam1 As String, Optional ByVal strParam2 As String, _
    Optional ByVal strOpToken As String, Optional ByVal sngSeconds As Single = 0.5, _
    Optional ByVal strCallbackInfo As String) As Boolean

    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String
    Dim lngInterval As Long

    On Error GoTo ErrHandler
    If TimerIsPending Then Exit Function
    If Len(strOpToken) = 0 Then strOpToken = Operation.CurrentRootToken

    SaveSetting PROJECT_NAME, "Timer", "Param1", strParam1
    SaveSetting PROJECT_NAME, "Timer", "Param2", strParam2
    SaveSetting PROJECT_NAME, "Timer", "CompatibilityEnvelope", CurrentCompatibilityEnvelope
    SaveSetting PROJECT_NAME, "Timer", "CallbackInfo", strCallbackInfo
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_OP_TOKEN, strOpToken
    SaveSetting PROJECT_NAME, "Timer", "Operation", strOperation
    ' Record when the request is due before arming, so its first event is never early.
    lngInterval = CLng(1000 * sngSeconds)
    m_lngArmedTick = GetTickCount
    m_lngArmedInterval = lngInterval
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_ARMED_TICK, m_lngArmedTick
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_ARMED_INTERVAL, lngInterval
    m_lngTimerID = ApiSetTimer(0, 0, lngInterval, AddressOf WinAPITimerCallback)
    SaveSetting PROJECT_NAME, "Timer", "TimerID", m_lngTimerID
    TrySetTimer = TimerIsPending
    If Not TrySetTimer Then ClearTimerRequest
    Exit Function

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    On Error Resume Next
    KillTimer
    ClearTimerRequest
    On Error GoTo 0
    Err.Raise lngErr, strSource, strErr
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearTimerRequest
' Purpose   : Consume or roll back the admitted request before another can be armed.
'---------------------------------------------------------------------------------------
Private Sub ClearTimerRequest()
    SaveSetting PROJECT_NAME, "Timer", "CompatibilityEnvelope", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "Operation", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "Param1", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "Param2", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "CallbackInfo", vbNullString
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_OP_TOKEN, vbNullString
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_ARMED_TICK, vbNullString
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_ARMED_INTERVAL, vbNullString
End Sub


'---------------------------------------------------------------------------------------
' Procedure : KillTimer
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Kill any existing timer
'---------------------------------------------------------------------------------------
'
Private Sub KillTimer()
    If m_lngTimerID = 0 Then m_lngTimerID = GetSetting(PROJECT_NAME, "Timer", "TimerID", 0)
    If m_lngTimerID <> 0 Then
        ApiKillTimer 0, m_lngTimerID
        Debug.Print "Killed API Timer " & m_lngTimerID
        m_lngTimerID = 0
        SaveSetting PROJECT_NAME, "Timer", "TimerID", 0
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HandleAPIAsyncOperation
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Handle async operation with MCP callbacks. Reads callback info from
'           : registry, registers with MCP, then starts the operation. Posts exactly
'           : one terminal callback for the call, even when no root was acquired.
'---------------------------------------------------------------------------------------
'
Private Sub HandleAPIAsyncOperation(strMethod As String, strArgs As String, strCallbackInfo As String, strAdmission As String)

    Dim strRefused As String
    strRefused = EnterCompatibilitySession(strAdmission)
    If Len(strRefused) > 0 Then
        PostCompatibilityRefusal strCallbackInfo, strRefused
        Exit Sub
    End If
    strRefused = CompatibilityCommandRefusal(strMethod)
    If Len(strRefused) > 0 Then
        PostCompatibilityRefusal strCallbackInfo, strRefused
        LeaveCompatibilitySession
        Exit Sub
    End If

    SuppressErrorBreaks
    LogUnhandledErrors
    On Error GoTo ErrHandler

    Dim strArg1 As String
    Dim strArg2 As String
    Dim lngPipePos As Long
    Dim varResult As Variant
    Dim intStatusBefore As eOperationState
    Dim blnRootPending As Boolean
    Dim cCallback As clsMCP
    Dim lngErr As Long
    Dim strErr As String
    Dim dError As Dictionary

    ' A timer admitted earlier can arrive while a synchronous API/root is running.
    ' Refuse before registering its callback or changing the owner's source.
    If APIDispatchBusy Or Operation.IsActive Then
        RefuseAsyncRequest strCallbackInfo, strMethod
        LeaveCompatibilitySession
        RestoreErrorBreaks
        Exit Sub
    End If

    ' Keep the admitted emitter alive across ReleaseObjects, including exception cleanup.
    Set cCallback = MCP
    cCallback.RegisterCallback strCallbackInfo
    If Len(strCallbackInfo) > 0 Then
        Operation.Source = eosMCPTool
    Else
        Operation.Source = eosExternalAPI
    End If

    ' Parse arguments (format: "arg1|arg2" or just "arg1")
    If Len(strArgs) > 0 Then
        lngPipePos = InStr(strArgs, "|")
        If lngPipePos > 0 Then
            strArg1 = Left(strArgs, lngPipePos - 1)
            strArg2 = Mid(strArgs, lngPipePos + 1)
        Else
            strArg1 = strArgs
        End If
    End If

    ' Start the operation via API
    ' Log.Add automatically routes to MCP when MCP.IsActive
    intStatusBefore = Operation.Status
    If Len(strArg2) > 0 Then
        varResult = API(strMethod, strArg1, strArg2, strSessionEnvelope:=CurrentCompatibilityEnvelope())
    ElseIf Len(strArg1) > 0 Then
        varResult = API(strMethod, strArg1, strSessionEnvelope:=CurrentCompatibilityEnvelope())
    Else
        varResult = API(strMethod, strSessionEnvelope:=CurrentCompatibilityEnvelope())
    End If

    ' Completion callback is sent from the root operation's completion, before ReleaseObjects
    MCPDebugLog "HandleAPIAsyncOperation: Operation complete, Result=" & Operation.Result

    ' A call that ended before a root posted its outcome (a pre-root error, or a Sub
    ' whose Begin was refused) is reported here, from its own return. A root this
    ' call left staged posts when its continuation finishes.
    blnRootPending = (intStatusBefore = eosReady And Operation.Status <> eosReady)
    PostUnreportedOutcome cCallback, strMethod, varResult, blnRootPending

    LeaveCompatibilitySession
    RestoreErrorBreaks
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Err.Clear
    If Not cCallback Is Nothing Then
        If cCallback.IsActive And Not cCallback.TerminalPosted Then
            Set dError = New Dictionary
            dError.Add "success", False
            dError.Add "error", strErr
            dError.Add "runtime_error", strErr
            dError.Add "errorNumber", lngErr
            cCallback.PostCallback "error", -1, -1, strMethod & " failed: " & strErr, dError
        End If
    Else
        AsyncRequestFailure strCallbackInfo, strErr, lngNumber:=lngErr
    End If
    LeaveCompatibilitySession
    RestoreErrorBreaks
    ' Never raise from a Windows timer callback after reporting its terminal outcome.
End Sub

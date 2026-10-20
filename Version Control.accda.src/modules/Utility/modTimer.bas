Attribute VB_Name = "modTimer"
'---------------------------------------------------------------------------------------
' Module    : modTimer
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : API timer functions for callbacks. Reserves callback capacity, validates
'           : explicitly owned continuations, and checks each Windows event before it
'           : consumes the armed request.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

' Windows API calls for Timer functionality
Private Declare PtrSafe Function ApiSetTimer Lib "user32" Alias "SetTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function ApiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Const WM_TIMER As Long = &H113
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const STILL_ACTIVE As Long = 259
' Allowance for GetTickCount granularity when judging whether an event is early.
Private Const TIMER_EARLY_MS As Long = 50
Private Const REG_TIMER_ARMED_TICK As String = "ArmedTick"
Private Const REG_TIMER_ARMED_INTERVAL As String = "ArmedInterval"
Private Const REG_TIMER_OWNER_PROCESS As String = "OwnerProcess"

Private m_lngTimerID As LongPtr
Private m_lngArmedTick As Long
Private m_lngArmedInterval As Long
Private m_strRunningCommand As String
Private m_blnFailNextTimer As Boolean
Private m_blnFailReservation As Boolean
Private m_strFailTimerOperation As String
Private m_lngProjectResetCalls As Long


'---------------------------------------------------------------------------------------
' Procedure : TimerIsPending
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Report installed or reserved callback capacity. The reservation is kept
'           : in the registry so it survives the VBA project reset between merge stages.
'           : A reservation left by an Access process that has since exited (a crash)
'           : is not pending: nothing can ever consume it, and it must not refuse every
'           : later request.
'---------------------------------------------------------------------------------------
'
Public Function TimerIsPending() As Boolean
    If m_lngTimerID <> 0 Then
        TimerIsPending = True
    ElseIf Len(GetSetting(PROJECT_NAME, "Timer", "Operation")) > 0 Then
        TimerIsPending = ReservationOwnerIsLive()
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReservationOwnerIsLive
' Purpose   : True when the saved reservation belongs to this process or to another
'           : Access process that is still running. A reservation without an owner was
'           : saved by older code and is treated as left over.
'---------------------------------------------------------------------------------------
'
Private Function ReservationOwnerIsLive() As Boolean
    Dim lngOwner As Long
    Dim lngExit As Long
    Dim hProcess As LongPtr
    lngOwner = CLng(Val(GetSetting(PROJECT_NAME, "Timer", REG_TIMER_OWNER_PROCESS, "0")))
    If lngOwner = 0 Then Exit Function
    If lngOwner = GetCurrentProcessId Then
        ReservationOwnerIsLive = True
        Exit Function
    End If
    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, lngOwner)
    If hProcess = 0 Then Exit Function
    If GetExitCodeProcess(hProcess, lngExit) <> 0 Then ReservationOwnerIsLive = (lngExit = STILL_ACTIVE)
    CloseHandle hProcess
End Function


'---------------------------------------------------------------------------------------
' Procedure : TimerRequestSucceeded
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Read explicit timer admission results.
'---------------------------------------------------------------------------------------
'
Public Function TimerRequestSucceeded(ByVal strResult As String) As Boolean
    Dim d As Dictionary
    Set d = ParseJson(strResult)
    If d.Exists("success") Then TimerRequestSucceeded = CBool(d("success"))
End Function


'---------------------------------------------------------------------------------------
' Procedure : FailNextTimerForTest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Inject one timer installation failure in a disposable qualification host.
'---------------------------------------------------------------------------------------
'
Public Sub FailNextTimerForTest()
    m_blnFailNextTimer = True
End Sub

' Admission and installation contain no DoEvents or network delivery. A missing token
' means a NEW request; it never acquires an existing root's continuation ownership.

'---------------------------------------------------------------------------------------
' Procedure : SetTimer
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Reserve and install a new callback only after root/token admission.
'           : strOpToken carries root ownership to the continuation; pass the token
'           : of the lease that was detached for this handoff. The session/1 envelope
'           : in force when the request is admitted is saved with it, so the queued
'           : work runs under the same session. Returns JSON; a refusal writes nothing.
'---------------------------------------------------------------------------------------
'
Public Function SetTimer(strOperation As String, _
    Optional strParam1 As String, Optional strParam2 As String, _
    Optional strOpToken As String, Optional sngSeconds As Single = 0.5, _
    Optional strCallbackInfo As String, Optional strRequestJson As String, _
    Optional blnReserveOnly As Boolean = False) As String

    Dim strReservation As String
    InitializeRequestErrors "SetTimer"
    On Error GoTo Failed
    If TimerIsPending Then
        SetTimer = RefusalJson("timer_busy", T("Another request is waiting for its callback timer."), False)
        Exit Function
    End If
    If Operation.IsActive Then
        If Len(strOpToken) = 0 Or Not Operation.RootLeaseIsActive(strOpToken) Then
            SetTimer = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
            Exit Function
        End If
    ElseIf Len(strOpToken) > 0 Then
        SetTimer = RefusalJson("stale_continuation", T("The continuation does not belong to the current operation."), False)
        Exit Function
    End If
    If StrComp(m_strFailTimerOperation, strOperation, vbBinaryCompare) = 0 And Len(m_strFailTimerOperation) > 0 Then
        m_strFailTimerOperation = vbNullString
        m_blnFailNextTimer = True
    End If
    If m_blnFailReservation Then
        m_blnFailReservation = False
        SetTimer = RefusalJson("timer_reservation_failed", T("The callback timer capacity could not be reserved."), False)
        Exit Function
    End If
    strReservation = NewOperationIdentity()
    SaveSetting PROJECT_NAME, "Timer", "Reservation", strReservation
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_OWNER_PROCESS, CStr(GetCurrentProcessId)
    SaveSetting PROJECT_NAME, "Timer", "Operation", strOperation
    SaveSetting PROJECT_NAME, "Timer", "Param1", strParam1
    SaveSetting PROJECT_NAME, "Timer", "Param2", strParam2
    SaveSetting PROJECT_NAME, "Timer", "OpToken", strOpToken
    SaveSetting PROJECT_NAME, "Timer", "CallbackInfo", strCallbackInfo
    SaveSetting PROJECT_NAME, "Timer", "RequestJson", strRequestJson
    SaveSetting PROJECT_NAME, "Timer", "CompatibilityEnvelope", CurrentCompatibilityEnvelope
    SaveSetting PROJECT_NAME, "Timer", "TimerID", "0"
    If blnReserveOnly Then
        SetTimer = "{""success"":true}"
    Else
        SetTimer = ActivateReservedTimer(strOpToken, sngSeconds)
    End If
    Exit Function
Failed:
    Dim strError As String
    strError = Err.Description
    If Len(strReservation) > 0 Then
        If GetSetting(PROJECT_NAME, "Timer", "Reservation") = strReservation Then KillTimer
        ClearTimerRequest strReservation
    End If
    SetTimer = RefusalJson("timer_install_failed", T("Could not install the callback timer: {0}", var0:=strError), False)
End Function


'---------------------------------------------------------------------------------------
' Procedure : TrySetTimer
' Purpose   : SetTimer for callers that only need to know whether the request was
'           : armed. The token is never defaulted: pass the owning root's token.
'---------------------------------------------------------------------------------------
'
Public Function TrySetTimer(ByVal strOperation As String, _
    Optional ByVal strParam1 As String, Optional ByVal strParam2 As String, _
    Optional ByVal strOpToken As String, Optional ByVal sngSeconds As Single = 0.5, _
    Optional ByVal strCallbackInfo As String) As Boolean
    TrySetTimer = TimerRequestSucceeded(SetTimer(strOperation, strParam1, strParam2, strOpToken, sngSeconds, strCallbackInfo))
End Function


'---------------------------------------------------------------------------------------
' Procedure : ActivateReservedTimer
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Activate matching reserved capacity and report Windows installation failure.
'           : The due time is recorded before arming, so the first event is never early.
'---------------------------------------------------------------------------------------
'
Public Function ActivateReservedTimer(ByVal strOpToken As String, Optional ByVal sngSeconds As Single = 0.5, _
    Optional ByVal blnRetainReservationOnFailure As Boolean = False) As String
    Dim strReservation As String
    Dim lngInterval As Long
    strReservation = GetSetting(PROJECT_NAME, "Timer", "Reservation")
    If Len(strOpToken) > 0 Then
        If Not Operation.RootLeaseIsActive(strOpToken) Then
            ActivateReservedTimer = RefusalJson("stale_continuation", T("The continuation no longer owns an active operation."), False)
            Exit Function
        End If
    ElseIf Operation.IsActive Then
        ActivateReservedTimer = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
        Exit Function
    End If
    If Len(strReservation) = 0 Or m_lngTimerID <> 0 Then
        ActivateReservedTimer = RefusalJson("timer_reservation_missing", T("The callback timer reservation is unavailable."), False)
        Exit Function
    End If
    If StrComp(GetSetting(PROJECT_NAME, "Timer", "OpToken"), strOpToken, vbBinaryCompare) <> 0 Then
        ActivateReservedTimer = RefusalJson("stale_continuation", T("The continuation does not own this timer reservation."), False)
        Exit Function
    End If
    lngInterval = CLng(1000 * sngSeconds)
    m_lngArmedTick = GetTickCount
    m_lngArmedInterval = lngInterval
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_ARMED_TICK, CStr(m_lngArmedTick)
    SaveSetting PROJECT_NAME, "Timer", REG_TIMER_ARMED_INTERVAL, CStr(lngInterval)
    If Not m_blnFailNextTimer Then m_lngTimerID = ApiSetTimer(0, 0, lngInterval, AddressOf WinAPITimerCallback)
    m_blnFailNextTimer = False
    If m_lngTimerID = 0 Then
        If Not blnRetainReservationOnFailure Then ClearTimerRequest strReservation
        ActivateReservedTimer = RefusalJson("timer_install_failed", T("Windows could not install the callback timer."), False)
        Exit Function
    End If
    SaveSetting PROJECT_NAME, "Timer", "TimerID", CStr(m_lngTimerID)
    ActivateReservedTimer = "{""success"":true}"
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReleaseTimerReservation
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Release only the reservation owned by the supplied root token.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseTimerReservation(ByVal strOpToken As String)
    If StrComp(GetSetting(PROJECT_NAME, "Timer", "OpToken"), strOpToken, vbBinaryCompare) <> 0 Then Exit Sub
    If Len(strOpToken) = 0 Then Exit Sub
    KillTimer
    ClearTimerRequest GetSetting(PROJECT_NAME, "Timer", "Reservation")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearTimerRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Clear fields only while the expected reservation still owns them.
'---------------------------------------------------------------------------------------
'
Private Sub ClearTimerRequest(ByVal strReservation As String)
    Dim varKey As Variant
    If Len(strReservation) = 0 Then Exit Sub
    If GetSetting(PROJECT_NAME, "Timer", "Reservation") <> strReservation Then Exit Sub
    For Each varKey In Array("Operation", "Param1", "Param2", "OpToken", "CallbackInfo", "RequestJson", _
        "CompatibilityEnvelope", REG_TIMER_ARMED_TICK, REG_TIMER_ARMED_INTERVAL, REG_TIMER_OWNER_PROCESS, "Reservation")
        SaveSetting PROJECT_NAME, "Timer", CStr(varKey), vbNullString
    Next varKey
    SaveSetting PROJECT_NAME, "Timer", "TimerID", "0"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : KillTimer
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Kill the currently recorded Windows timer.
'---------------------------------------------------------------------------------------
'
Private Sub KillTimer()
    If m_lngTimerID = 0 Then m_lngTimerID = CLngPtr(GetSetting(PROJECT_NAME, "Timer", "TimerID", "0"))
    If m_lngTimerID <> 0 Then ApiKillTimer 0, m_lngTimerID
    m_lngTimerID = 0
End Sub


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
'           : This covers failures on the callback's own VBA call stack, including the
'           : directly called request cases; an Application.Run target must trap its own
'           : errors (see DispatchTimerRequest).
'---------------------------------------------------------------------------------------
'
Public Sub WinAPITimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, ByVal dwTime As Long)

    On Error GoTo ErrHandler
    m_strRunningCommand = vbNullString
    If IsArmedEvent(uMsg, idEvent) Then RunPendingTimerRequest
    Exit Sub

ErrHandler:
    RecordTimerFailure Err.Number, Err.Description
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RecordTimerFailure
' Purpose   : Record a failed timer request as a warning, without raising.
'---------------------------------------------------------------------------------------
'
Private Sub RecordTimerFailure(ByVal lngErr As Long, ByVal strErr As String)
    Dim strCommand As String
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
        ArmedTimerID = CLngPtr(Val(GetSetting(PROJECT_NAME, "Timer", "TimerID", "0")))
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
    ' Nothing armed, nothing to consume.
    If ArmedTimerID = 0 Then Exit Sub
    DispatchTimerRequest GetSetting(PROJECT_NAME, "Timer", "Reservation")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DispatchTimerRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Consume matching queued work and resume only explicit owner tokens.
'           : Root ownership crosses this boundary in the saved token, never in mutable
'           : global state: a continuation resumes only the root it was armed for, and
'           : a stale or foreign token resumes nothing.
'---------------------------------------------------------------------------------------
'
Private Sub DispatchTimerRequest(ByVal strExpectedReservation As String)
    Dim strCommand As String
    Dim strParam1 As String
    Dim strParam2 As String
    Dim strToken As String
    Dim strRequest As String
    Dim strAdmission As String
    Dim strReservation As String
    Dim strQueueResult As String
    Dim cRequest As clsOperationRequest
    Dim cPrior As clsOperationRequest
    Dim lngErr As Long
    Dim strErr As String

    If Len(strExpectedReservation) = 0 Then Exit Sub
    If GetSetting(PROJECT_NAME, "Timer", "Reservation") <> strExpectedReservation Then Exit Sub
    strCommand = GetSetting(PROJECT_NAME, "Timer", "Operation")
    m_strRunningCommand = strCommand
    strParam1 = GetSetting(PROJECT_NAME, "Timer", "Param1")
    strParam2 = GetSetting(PROJECT_NAME, "Timer", "Param2")
    strToken = GetSetting(PROJECT_NAME, "Timer", "OpToken")
    strRequest = GetSetting(PROJECT_NAME, "Timer", "RequestJson")
    strAdmission = GetSetting(PROJECT_NAME, "Timer", "CompatibilityEnvelope")
    strReservation = GetSetting(PROJECT_NAME, "Timer", "Reservation")

    ' Consume before running, in case the operation sets another timer.
    KillTimer
    ClearTimerRequest strReservation

    Select Case strCommand
        Case "Build"
            ' Build from source (full or merge build)
            RunBuildFromContinuation strToken, strParam1, CBool(strParam2)

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
            If Not Operation.RootLeaseIsActive(strToken) Then GoTo Finished
            If Operation.Status = eosRunning Then Operation.DetachRootLease strToken
            strQueueResult = SetTimer("MergeResume", strParam1, strParam2, strToken)
            If Not TimerRequestSucceeded(strQueueResult) Then
                Operation.RecordRootRuntimeError strToken, T("Could not queue the merge continuation."), vbObjectError + 2410
                Operation.CompleteRootLease strToken, eorFailed
                GoTo Finished
            End If
            TraceInPlaceMerge "reset stage: merge timer armed"
            m_lngProjectResetCalls = m_lngProjectResetCalls + 1
            ResetProjectForInPlaceMerge

        Case "MergeResume"
            ' Continue a merge build after the database was prepared in place and its VBA
            ' project reset. (See modBuild.PrepareMergeInPlace.)
            RunBuildFromContinuation strToken, strParam1, False, Val(strParam2), vbNullString, True

        Case "APIAsyncOperation"
            ' Handle async operation with MCP callbacks
            HandleAPIAsyncOperation strParam1, strParam2, strRequest, strAdmission

        Case "QuitForRebuild"
            ' Close this instance so the rebuild worker can replace the files it holds.
            ' Armed by clsVersionControl.RebuildAddIn before the worker is launched,
            ' because the worker cannot close an instance it could not attach to, and it
            ' cannot attach when the current database is the add-in itself.
            ' (See clsWorker.Main.)
            If Not Operation.RootLeaseIsActive(strToken) Then GoTo Finished
            Application.Quit acQuitSaveAll

        Case Else
            ' A continuation runs only for the active root that armed it. Work queued
            ' without a token is a new command, refused while other work is admitted.
            If Len(strToken) > 0 Then
                If Not Operation.RootLeaseIsActive(strToken) Then GoTo Finished
            ElseIf ManualCommandRefused Then
                GoTo Finished
            End If
            Set cPrior = IncomingOperationRequest
            If Len(strRequest) > 0 Then
                Set cRequest = New clsOperationRequest
                cRequest.FromJson strRequest
                Set IncomingOperationRequest = cRequest
            End If
            ' Use the Run command to execute the specified operation with supplied parameters.
            ' A target that cannot be resolved fails here and is recorded as a warning.
            ' An error raised inside a resolved target is that target's own: in a native
            ' run, VBA presented it as an unhandled error in the target and did not return
            ' it here. A Run target must trap its own errors.
            On Error GoTo DispatchFailed
            If strParam2 <> vbNullString Then
                Application.Run strCommand, strParam1, strParam2
            ElseIf strParam1 <> vbNullString Then
                Application.Run strCommand, strParam1
            Else
                Application.Run strCommand
            End If
            On Error GoTo 0
            Set IncomingOperationRequest = cPrior
    End Select

Finished:
    m_strRunningCommand = vbNullString
    Exit Sub

DispatchFailed:
    lngErr = Err.Number
    strErr = Err.Description
    Set IncomingOperationRequest = cPrior
    RecordTimerFailure lngErr, strErr
End Sub

' This emitter is request-local on every path. In particular, an execution-time
' refusal never registers over another root's shared callback.

'---------------------------------------------------------------------------------------
' Procedure : HandleAPIAsyncOperation
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Run an admitted asynchronous request under the session it was admitted
'           : with. Admission is checked again, refusals are posted through the queued
'           : request's own callback, and exactly one terminal callback is posted for
'           : the call, even when no root was acquired. A root this call started posts
'           : its own outcome. Nothing is raised from here: it runs inside a Windows
'           : timer callback.
'---------------------------------------------------------------------------------------
'
Public Sub HandleAPIAsyncOperation(ByVal strMethod As String, ByVal strArgs As String, _
    ByVal strRequest As String, Optional ByVal strAdmission As String)

    Dim cRequest As New clsOperationRequest
    Dim cEmitter As New clsMCP
    Dim cOwnerEmitter As clsMCP
    Dim colArgs As Collection
    Dim varResult As Variant
    Dim dResult As Dictionary
    Dim strRefusal As String
    Dim strError As String
    Dim lngErr As Long
    Dim blnEntered As Boolean

    ' Re-admission failures must not log into another operation.
    InitializeRequestErrors "HandleAPIAsyncOperation"
    On Error GoTo Failed
    cRequest.FromJson strRequest
    cEmitter.RegisterCallback cRequest.CallbackInfo

    strRefusal = EnterCompatibilitySession(strAdmission)
    If Len(strRefusal) > 0 Then
        PostCompatibilityRefusal cRequest.CallbackInfo, strRefusal
        Exit Sub
    End If
    blnEntered = True
    strRefusal = CompatibilityCommandRefusal(strMethod)
    If Len(strRefusal) = 0 Then strRefusal = RequestAdmissionRefusal(blnQuiet:=True)
    If Len(strRefusal) = 0 And APIDispatchBusy Then strRefusal = RefusalJson("api_busy", T("Another API command is still running."), False)
    If Len(strRefusal) > 0 Then
        Set dResult = ParseJson(strRefusal)
        cEmitter.PostCallback "error", -1, -1, CStr(dResult("error")), dResult
        GoTo CleanUp
    End If

    Set colArgs = ParseJson(strArgs)
    CountCompatibilityDispatch
    ' Keep the owner's emitter alive across ReleaseObjects, so its terminal flag can
    ' still be read after the root completes.
    Set cOwnerEmitter = MCP
    varResult = DispatchOperationRequest(cRequest, strMethod, colArgs)
    MCPDebugLog "HandleAPIAsyncOperation: dispatch returned for " & strMethod

    ' The root posted its outcome, or a root this request started is staged and its
    ' continuation posts later.
    If cOwnerEmitter.TerminalPosted Then GoTo CleanUp
    If Operation.IsActive Then
        If Operation.PublicId = cRequest.PublicId Then GoTo CleanUp
    End If

    ' The call ended before a root posted its outcome: report it from its own return.
    Set dResult = Nothing
    If VarType(varResult) = vbString Then
        If Left$(Trim$(CStr(varResult)), 1) = "{" Then Set dResult = ParseJson(CStr(varResult))
    End If
    If Not dResult Is Nothing Then
        If dResult.Exists("success") Then
            If CBool(dResult("success")) Then
                cEmitter.PostCallback "complete", -1, -1, vbNullString, dResult
                GoTo CleanUp
            End If
        End If
    End If
    PostUnreportedOutcome cEmitter, strMethod, varResult, False

CleanUp:
    If blnEntered Then LeaveCompatibilitySession
    Exit Sub

Failed:
    lngErr = Err.Number
    strError = Err.Description
    Err.Clear
    On Error Resume Next
    If blnEntered Then LeaveCompatibilitySession
    If cEmitter.TerminalPosted Then Exit Sub
    If Not cOwnerEmitter Is Nothing Then
        If cOwnerEmitter.TerminalPosted Then Exit Sub
    End If
    Set dResult = New Dictionary
    dResult("success") = False
    dResult("error") = strError
    dResult("error_pattern") = "dispatch_error"
    dResult("runtime_error") = strError
    dResult("errorNumber") = lngErr
    cEmitter.PostCallback "error", -1, -1, strMethod & " failed: " & strError, dResult
    ' Never raise from a Windows timer callback after reporting its terminal outcome.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FailNextTimerReservationForTest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Inject one prelaunch callback-capacity failure.
'---------------------------------------------------------------------------------------
'
Public Sub FailNextTimerReservationForTest()
    m_blnFailReservation = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DispatchReservedTimerForTest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Enter actual dispatch for the currently reserved disposable request.
'---------------------------------------------------------------------------------------
'
Public Sub DispatchReservedTimerForTest()
    DispatchTimerRequest GetSetting(PROJECT_NAME, "Timer", "Reservation")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FailNamedTimerForTest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Inject failure only for the named continuation handoff.
'---------------------------------------------------------------------------------------
'
Public Sub FailNamedTimerForTest(ByVal strOperation As String)
    m_strFailTimerOperation = strOperation
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ProjectResetCallsForTest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Observe whether qualification reached the real project-reset boundary.
'---------------------------------------------------------------------------------------
'
Public Property Get ProjectResetCallsForTest() As Long
    ProjectResetCallsForTest = m_lngProjectResetCalls
End Property

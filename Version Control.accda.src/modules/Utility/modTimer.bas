Attribute VB_Name = "modTimer"
'---------------------------------------------------------------------------------------
' Module    : modTimer
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Reserves callback capacity and validates explicitly owned continuations.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

Private Declare PtrSafe Function ApiSetTimer Lib "user32" Alias "SetTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function ApiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private m_lngTimerID As LongPtr
Private m_blnFailNextTimer As Boolean
Private m_blnFailReservation As Boolean
Private m_strFailTimerOperation As String
Private m_lngProjectResetCalls As Long


'---------------------------------------------------------------------------------------
' Procedure : TimerIsPending
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Report installed or reserved callback capacity.
'---------------------------------------------------------------------------------------
'
Public Function TimerIsPending() As Boolean
    TimerIsPending = (m_lngTimerID <> 0) Or Len(GetSetting(PROJECT_NAME, "Timer", "Operation")) > 0
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
    SaveSetting PROJECT_NAME, "Timer", "Operation", strOperation
    SaveSetting PROJECT_NAME, "Timer", "Param1", strParam1
    SaveSetting PROJECT_NAME, "Timer", "Param2", strParam2
    SaveSetting PROJECT_NAME, "Timer", "OpToken", strOpToken
    SaveSetting PROJECT_NAME, "Timer", "CallbackInfo", strCallbackInfo
    SaveSetting PROJECT_NAME, "Timer", "RequestJson", strRequestJson
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
' Procedure : ActivateReservedTimer
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Activate matching reserved capacity and report Windows installation failure.
'---------------------------------------------------------------------------------------
'
Public Function ActivateReservedTimer(ByVal strOpToken As String, Optional ByVal sngSeconds As Single = 0.5, _
    Optional ByVal blnRetainReservationOnFailure As Boolean = False) As String
    Dim strReservation As String
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
    If Not m_blnFailNextTimer Then m_lngTimerID = ApiSetTimer(0, 0, CLng(1000 * sngSeconds), AddressOf WinAPITimerCallback)
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
    For Each varKey In Array("Operation", "Param1", "Param2", "OpToken", "CallbackInfo", "RequestJson", "Reservation")
        SaveSetting PROJECT_NAME, "Timer", CStr(varKey), vbNullString
    Next varKey
    SaveSetting PROJECT_NAME, "Timer", "TimerID", "0"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : KillTimer
' Author    : Codex
' Date      : 10/16/2026
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
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Ignore stale Windows events before touching a newer queue.
'---------------------------------------------------------------------------------------
'
Public Sub WinAPITimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal nIDEvent As LongPtr, ByVal dwTime As Long)
    Dim lngExpected As LongPtr
    lngExpected = m_lngTimerID
    If lngExpected = 0 Then lngExpected = CLngPtr(GetSetting(PROJECT_NAME, "Timer", "TimerID", "0"))
    ' A killed timer can have an already-queued event. Validate its Windows ID
    ' before touching the current timer or reading/clearing its request fields.
    If nIDEvent = 0 Or nIDEvent <> lngExpected Then Exit Sub
    DispatchTimerRequest GetSetting(PROJECT_NAME, "Timer", "Reservation")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DispatchTimerRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Consume matching queued work and resume only explicit owner tokens.
'---------------------------------------------------------------------------------------
'
Private Sub DispatchTimerRequest(ByVal strExpectedReservation As String)
    Dim strCommand As String
    Dim strParam1 As String
    Dim strParam2 As String
    Dim strToken As String
    Dim strCallback As String
    Dim strRequest As String
    Dim strReservation As String
    Dim strQueueResult As String
    Dim cRequest As clsOperationRequest
    Dim cPrior As clsOperationRequest

    If Len(strExpectedReservation) = 0 Then Exit Sub
    If GetSetting(PROJECT_NAME, "Timer", "Reservation") <> strExpectedReservation Then Exit Sub
    strCommand = GetSetting(PROJECT_NAME, "Timer", "Operation")
    strParam1 = GetSetting(PROJECT_NAME, "Timer", "Param1")
    strParam2 = GetSetting(PROJECT_NAME, "Timer", "Param2")
    strToken = GetSetting(PROJECT_NAME, "Timer", "OpToken")
    strCallback = GetSetting(PROJECT_NAME, "Timer", "CallbackInfo")
    strRequest = GetSetting(PROJECT_NAME, "Timer", "RequestJson")
    strReservation = GetSetting(PROJECT_NAME, "Timer", "Reservation")
    KillTimer
    ClearTimerRequest strReservation

    Select Case strCommand
        Case "Build"
            RunBuildFromContinuation strToken, strParam1, CBool(strParam2)
        Case "MergeReset"
            If Not Operation.RootLeaseIsActive(strToken) Then Exit Sub
            If Operation.Status = eosRunning Then Operation.DetachRootLease strToken
            strQueueResult = SetTimer("MergeResume", strParam1, strParam2, strToken)
            If Not TimerRequestSucceeded(strQueueResult) Then
                Operation.RecordRootRuntimeError strToken, T("Could not queue the merge continuation."), vbObjectError + 2410
                Operation.CompleteRootLease strToken, eorFailed
                Exit Sub
            End If
            TraceInPlaceMerge "reset stage: merge timer armed"
            m_lngProjectResetCalls = m_lngProjectResetCalls + 1
            ResetProjectForInPlaceMerge
        Case "MergeResume"
            RunBuildFromContinuation strToken, strParam1, False, Val(strParam2), vbNullString, True
        Case "APIAsyncOperation"
            HandleAPIAsyncOperation strParam1, strParam2, strRequest
        Case "QuitForRebuild"
            If Not Operation.RootLeaseIsActive(strToken) Then Exit Sub
            Application.Quit acQuitSaveAll
        Case Else
            If ManualCommandRefused Then Exit Sub
            Set cPrior = IncomingOperationRequest
            If Len(strRequest) > 0 Then
                Set cRequest = New clsOperationRequest
                cRequest.FromJson strRequest
                Set IncomingOperationRequest = cRequest
            End If
            LogUnhandledErrors
            On Error GoTo DispatchFailed
            If strParam2 <> vbNullString Then
                Application.Run strCommand, strParam1, strParam2
            ElseIf strParam1 <> vbNullString Then
                Application.Run strCommand, strParam1
            Else
                Application.Run strCommand
            End If
DispatchFailed:
            Set IncomingOperationRequest = cPrior
    End Select
End Sub

' This emitter is request-local on every path. In particular, an execution-time
' refusal never registers over another root's shared callback.

'---------------------------------------------------------------------------------------
' Procedure : HandleAPIAsyncOperation
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Recheck queued admission and emit refusals through the queued request owner.
'---------------------------------------------------------------------------------------
'
Public Sub HandleAPIAsyncOperation(ByVal strMethod As String, ByVal strArgs As String, ByVal strRequest As String)
    Dim cRequest As New clsOperationRequest
    Dim cEmitter As New clsMCP
    Dim cOwnerEmitter As clsMCP
    Dim colArgs As Collection
    Dim varResult As Variant
    Dim dResult As Object
    Dim strRefusal As String
    Dim strError As String

    ' Re-admission failures must not log into another operation.
    InitializeRequestErrors "HandleAPIAsyncOperation"
    On Error GoTo Failed
    cRequest.FromJson strRequest
    cEmitter.RegisterCallback cRequest.CallbackInfo
    strRefusal = RequestAdmissionRefusal(blnQuiet:=True)
    If APIDispatchBusy Then strRefusal = RefusalJson("api_busy", T("Another API command is still running."), False)
    If Len(strRefusal) > 0 Then
        Set dResult = ParseJson(strRefusal)
        cEmitter.PostCallback "error", -1, -1, CStr(dResult("error")), dResult
        Exit Sub
    End If
    Set colArgs = ParseJson(strArgs)
    Set cOwnerEmitter = MCP
    varResult = DispatchOperationRequest(cRequest, strMethod, colArgs)
    If cOwnerEmitter.TerminalPosted Then Exit Sub
    If Operation.IsActive Then
        If Operation.PublicId = cRequest.PublicId Then Exit Sub
    End If
    If VarType(varResult) = vbString And Len(CStr(varResult)) > 0 Then
        Set dResult = ParseJson(CStr(varResult))
    Else
        Set dResult = New Dictionary
        dResult("success") = False
        dResult("error_pattern") = "operation_not_started"
        dResult("error") = T("The operation returned without starting or reporting a result.")
    End If
    Dim strType As String
    strType = "error"
    If dResult.Exists("success") Then
        If CBool(dResult("success")) Then strType = "complete"
    End If
    cEmitter.PostCallback strType, -1, -1, vbNullString, dResult
    Exit Sub
Failed:
    strError = Err.Description
    If cEmitter.TerminalPosted Then Exit Sub
    If Not cOwnerEmitter Is Nothing Then
        If cOwnerEmitter.TerminalPosted Then Exit Sub
    End If
    Set dResult = New Dictionary
    dResult("success") = False
    dResult("error") = strError
    dResult("error_pattern") = "dispatch_error"
    cEmitter.PostCallback "error", -1, -1, strError, dResult
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

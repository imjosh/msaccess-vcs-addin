VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsOperation
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Root operation ownership, validated continuations, and pause scopes.
'           : Exactly one root controls registry state, heartbeat, logs, callbacks, and
'           : teardown. Core work never begins or finishes an operation directly.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Private Const HEARTBEAT_TIMEOUT_SECONDS As Integer = 60 * 10
Private Const REGISTRY_PULSE_SECONDS As Integer = 60
Private Const REG_KEY_ROOT_TOKEN As String = "RootToken"
Private Const REG_KEY_ROOT_TYPE As String = "RootOperationType"
Private Const REG_KEY_ROOT_PROJECT As String = "RootProject"


Public Heartbeat As Date
Public Result As eOperationResult
Private m_intIdleSource As eOperationSource
Private m_RootEmitter As clsMCP
Private m_strAdmissionRefusal As String
Private m_TransferOutcome As Dictionary
Private m_PolicyComparisonRefusal As Dictionary

' Set by headless entry points before TryBeginRoot/Begin so Attended stays False even
' when Source is still eosUserInterface (Immediate Window callers).
Public ForceUnattended As Boolean


Private Type udtRoot
    Token As String
    PublicId As String
    InitiatingSource As eOperationSource
    Policy As eDecisionPolicy
    CallbackInfo As String
    OperationType As eOperationType
    Attended As Boolean
    StartTime As Date
    FinishTime As Date
    State As eOperationState
    AppHwnd As Long
    ErrorTrapping As eVbeErrorTrapping
    Detached As Boolean
    Completed As Boolean
    RebuildHandoffHold As Boolean
    PauseDepth As Long
    RuntimeError As String
    RuntimeErrorNumber As Long
    CancelRequested As Boolean
    OutputPath As String
    NativeCloseCanceled As Boolean
End Type
Private m_Root As udtRoot

' Owns interaction mode, decision policy, blocked flag, decision journal and error level.
Private m_Scope As clsInteractionScope

' Journal of the operation that finished last (see LastDecisions).
Private m_LastDecisions As Collection

' Terminal callback of the operation that finished last (see LastCompletion).
Private m_LastCompletion As Dictionary

' Heartbeat value last copied to the registry, throttling Pulse's registry write.
Private m_datRegistryPulse As Date

' Lease issued by Begin. It has to be held here for the life of the root: a lease
' released while the root is still active abandons it (clsRootOperationLease.
' Class_Terminate), and Begin's caller does not complete until Finish.
Private m_cSyncLease As clsRootOperationLease

' Instance-owned, one-shot completion-construction fault for lifecycle regressions.
Private m_lngCompletionFaultForTest As Long

' Test seam for the cancel poll (see SetCancelPollForTest). Zero when disarmed. Roots
' do not reset it, so a test can arm it before the run it is meant to stop.
Private m_lngFakeCancelOnCheck As Long
Private m_lngFakeCancelChecks As Long


'---------------------------------------------------------------------------------------
' Procedure : TryBeginRoot
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Acquire exclusive root ownership. Returns Nothing when refused.
'           : A policy other than edpAsk opens an operation-owned noninteractive scope,
'           : but only once the root is granted. A refused request leaves the running
'           : operation's interaction state untouched and never shows a dialog.
'---------------------------------------------------------------------------------------
'
Public Function TryBeginRoot(intType As eOperationType, Optional ByVal intPolicy As eDecisionPolicy = edpAsk, _
    Optional ByVal blnUnattended As Boolean = False, Optional ByVal intMode As Long = -1) As clsRootOperationLease

    Dim cLease As clsRootOperationLease

    If Not CanBeginRoot(intPolicy) Then Exit Function

    Dim cRequest As clsOperationRequest
    If IsSessionSingleton Then Set cRequest = IncomingOperationRequest
    If cRequest Is Nothing Then Set cRequest = CreateOperationRequest(m_intIdleSource)
    m_strAdmissionRefusal = ClaimOperationRequest(cRequest, IsSessionSingleton)
    If Len(m_strAdmissionRefusal) > 0 Then Exit Function
    InitializeRoot intType, intPolicy, cRequest, blnUnattended, intMode
    cRequest.Consumed = True
    Set cLease = New clsRootOperationLease
    cLease.Init Me, m_Root.Token
    Set TryBeginRoot = cLease

End Function


'---------------------------------------------------------------------------------------
' Procedure : Begin
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Synchronous root entry for work that begins and ends in one procedure.
'           : Wraps TryBeginRoot and discards the lease handle; the caller completes
'           : with Finish on the same stack. Use TryBeginRoot when ownership crosses a
'           : timer, a form lifetime, or a procedure boundary instead.
'---------------------------------------------------------------------------------------
'
Public Function Begin(intType As eOperationType, Optional ByVal intPolicy As eDecisionPolicy = edpAsk, _
    Optional ByVal blnUnattended As Boolean = False, Optional ByVal intMode As Long = -1) As Boolean

    Dim cLease As clsRootOperationLease

    ' Acquire into a local. A refused Begin returns Nothing, and storing that would
    ' release the lease of the root that is already running here, abandoning it.
    Set cLease = TryBeginRoot(intType, intPolicy, blnUnattended, intMode)
    If cLease Is Nothing Then Exit Function

    Set m_cSyncLease = cLease
    Begin = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResumeRoot
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Resume a detached root from a timer continuation. Fails closed on mismatch.
'---------------------------------------------------------------------------------------
'
Public Function ResumeRoot(strToken As String) As clsRootOperationLease

    Dim cLease As clsRootOperationLease

    If Len(strToken) = 0 Then Exit Function
    If m_Root.Completed Then Exit Function
    If m_Root.State <> eosStaged Then Exit Function
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Function
    If m_Root.AppHwnd <> hWndAccessApp Then Exit Function

    EnsureRootEmitter
    m_Root.Detached = False
    m_Root.State = eosRunning
    SetErrorTrapping
    UpdateRegistry
    Pulse

    Set cLease = New clsRootOperationLease
    cLease.Init Me, m_Root.Token
    Set ResumeRoot = cLease

End Function


'---------------------------------------------------------------------------------------
' Procedure : CompleteRootLease
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Complete the root when the caller holds a matching lease token.
'---------------------------------------------------------------------------------------
'
Public Sub CompleteRootLease(strToken As String, Optional intResult As eOperationResult = eorUnknown)

    If m_Root.Completed Then Exit Sub
    If Len(m_Root.Token) = 0 Then Exit Sub
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Sub

    FinishRootInternal intResult

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RecordRootRuntimeError
' Date      : 09/30/2026
' Purpose   : Keep the runtime error that ended the root, when the caller holds a
'           : matching lease token. Finish puts it on the terminal callback, since an
'           : async caller never sees the procedure's own return value.
'---------------------------------------------------------------------------------------
'
Public Sub RecordRootRuntimeError(strToken As String, strDescription As String, lngNumber As Long)

    If m_Root.Completed Then Exit Sub
    If Len(m_Root.Token) = 0 Then Exit Sub
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Sub

    m_Root.RuntimeError = strDescription
    m_Root.RuntimeErrorNumber = lngNumber

End Sub





'---------------------------------------------------------------------------------------
' Procedure : RecordOutputPath
' Purpose   : Record the file a successful build actually wrote before completion.
'---------------------------------------------------------------------------------------
'
Public Sub RecordOutputPath(ByVal strPath As String)
    If m_Root.Completed Then Exit Sub
    m_Root.OutputPath = strPath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DetachRootLease
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Stage the root for an asynchronous continuation.
'---------------------------------------------------------------------------------------
'
Public Function DetachRootLease(strToken As String) As String

    If m_Root.Completed Then Exit Function
    If Len(m_Root.Token) = 0 Then Exit Function
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Function
    If m_Root.State <> eosRunning Then Exit Function

    m_Root.Detached = True
    m_Root.State = eosStaged
    RestoreUserErrorTrapping
    UpdateRegistry
    Pulse
    DetachRootLease = m_Root.Token

End Function


'---------------------------------------------------------------------------------------
' Procedure : AbandonRootLease
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Safety-net cleanup when a lease is dropped without Complete.
'---------------------------------------------------------------------------------------
'
Public Sub AbandonRootLease(strToken As String, strSource As String)

    If m_Root.Completed Then Exit Sub
    If Len(m_Root.Token) = 0 Then Exit Sub
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Sub

    LogDiagnostic "Root lease abandoned", strSource
    FinishRootInternal eorFailed

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RootLeaseIsActive
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : True when the token matches the live, incomplete root.
'---------------------------------------------------------------------------------------
'
Public Function RootLeaseIsActive(strToken As String) As Boolean

    RootLeaseIsActive = (Not m_Root.Completed) _
        And Len(m_Root.Token) > 0 _
        And StrComp(m_Root.Token, strToken, vbBinaryCompare) = 0 _
        And (m_Root.State = eosRunning Or m_Root.State = eosStaged)

End Function


'---------------------------------------------------------------------------------------
' Procedure : Finish
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Synchronous root completion for work begun with Begin on the same stack.
'           : Finish supplies the live root token to itself, so it performs no ownership
'           : check: only call it from the procedure that began the root.
'---------------------------------------------------------------------------------------
'
Public Function Finish(Optional intResult As eOperationResult = eorUnknown)

    If Len(m_Root.Token) = 0 Then Exit Function
    CompleteRootLease m_Root.Token, intResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : TryPause
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Begin a short synchronous pause around foreign code: user VBA invoked
'           : through Application.Run, or closing a form whose unload handler would
'           : otherwise read a running operation as a cancel request. The root is
'           : suspended, not detached - no continuation token is issued, and the pause
'           : ends on the same stack that started it.
'---------------------------------------------------------------------------------------
'
Public Function TryPause() As clsOperationPause

    Dim cPause As clsOperationPause

    ' Pauses nest: an already-paused root is staged, so accept that too, but never a
    ' root that was detached for an asynchronous continuation.
    If m_Root.State <> eosRunning Then
        If m_Root.PauseDepth = 0 Then Exit Function
        If m_Root.State <> eosStaged Or m_Root.Detached Then Exit Function
    End If

    Set cPause = New clsOperationPause
    cPause.Init Me
    Set TryPause = cPause

End Function


'---------------------------------------------------------------------------------------
' Procedure : BeginPauseScope
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Internal entry for clsOperationPause.
'---------------------------------------------------------------------------------------
'
Public Sub BeginPauseScope()

    m_Root.PauseDepth = m_Root.PauseDepth + 1
    If m_Root.PauseDepth > 1 Then Exit Sub

    If m_Root.State = eosRunning Then
        m_Root.State = eosStaged
        m_Root.Detached = False
        RestoreUserErrorTrapping
        UpdateRegistry
        Pulse
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EndPauseScope
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Internal exit for clsOperationPause.
'---------------------------------------------------------------------------------------
'
Public Sub EndPauseScope()

    If m_Root.PauseDepth <= 0 Then Exit Sub
    m_Root.PauseDepth = m_Root.PauseDepth - 1
    If m_Root.PauseDepth > 0 Then Exit Sub

    ' A root detached for a continuation while paused is no longer ours to resume.
    If m_Root.State = eosStaged And Not m_Root.Detached Then
        m_Root.State = eosRunning
        SetErrorTrapping
        UpdateRegistry
        Pulse
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Pulse
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Pulse the heartbeat. Callers pulse this once per component, so the
'           : in-memory stamp is all that is updated on most calls. The registry copy
'           : only exists so a VBA project reset can recognize a root that is still
'           : running, and is refreshed on a timer rather than per component.
'---------------------------------------------------------------------------------------
'
Public Function Pulse()

    Heartbeat = Now
    If m_Root.State = eosReady Then Exit Function
    If DateDiff("s", m_datRegistryPulse, Heartbeat) < REGISTRY_PULSE_SECONDS Then Exit Function
    m_datRegistryPulse = Heartbeat
    UpdateRegistry

End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckCancelRequest
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : The one checkpoint for an MCP cancel request. The test loop and the
'           : per-object loops of export and build call it at their safe points. A
'           : request goes to RequestCancel, which sends the signal that closing the
'           : main form sends. Only a running root polls. A paused root is running
'           : foreign code. Returns True once this root has seen a request.
'---------------------------------------------------------------------------------------
'
Public Function CheckCancelRequest() As Boolean

    If m_Root.CancelRequested Then
        CheckCancelRequest = True
        Exit Function
    End If
    If m_Root.State <> eosRunning Then Exit Function
    If Not CancelPolled() Then Exit Function

    CheckCancelRequest = RequestCancel("clsOperation.CheckCancelRequest")

End Function


'---------------------------------------------------------------------------------------
' Procedure : RequestCancel
' Author    : Josh
' Date      : 10/19/2026
' Purpose   : The one owner of a cancel request for this root. A request can come from
'           : the MCP poll, from closing the main form, or from a dismiss cancel. This
'           : records it and sends the signal the loops stop on. A test run gets
'           : TestRunner.Cancel. Export, build and merge get a critical error. Only a
'           : running root accepts a request. With no root, a finished one, or a
'           : staged or paused one, nothing changes and the result is False. A request
'           : already recorded returns True and is not sent again. A request is not a
'           : result. Finish reports the root cancelled unless a blocked decision or a
'           : runtime error takes precedence.
'---------------------------------------------------------------------------------------
'
Public Function RequestCancel(Optional ByVal strSource As String = "clsOperation.RequestCancel") As Boolean

    If m_Root.Completed Then Exit Function
    If m_Root.CancelRequested Then
        RequestCancel = True
        Exit Function
    End If
    If m_Root.State <> eosRunning Then Exit Function

    m_Root.CancelRequested = True
    RequestCancel = True
    If m_Root.OperationType = eotTestRun Then
        If IsSessionSingleton Then TestRunner.Cancel
    ElseIf IsSessionSingleton Then
        Log.Error eelCritical, T("Canceled Operation"), strSource
    Else
        ' A private instance must not write to the running operation's log.
        Me.ErrorLevel = eelCritical
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CancelRequested
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : True once RequestCancel has accepted a cancel request for this root. Kept
'           : after Finish until the next root begins.
'---------------------------------------------------------------------------------------
'
Public Property Get CancelRequested() As Boolean
    CancelRequested = m_Root.CancelRequested
End Property


'---------------------------------------------------------------------------------------
' Procedure : HasRootRuntimeError
' Author    : Josh
' Date      : 10/19/2026
' Purpose   : True once RecordRootRuntimeError has kept an error for this root. A
'           : wrap-up that reports its own result reads it, so a runtime error keeps
'           : precedence over a cancel request as it does in Finish.
'---------------------------------------------------------------------------------------
'
Public Property Get HasRootRuntimeError() As Boolean
    HasRootRuntimeError = (Len(m_Root.RuntimeError) > 0)
End Property


'---------------------------------------------------------------------------------------
' Procedure : CancelPolled
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : The seam around the cancel poll. An armed fake answers in place of MCP.
'           : Otherwise only the session operation asks, through the emitter that
'           : holds its callback. CheckCancelled returns at once, with no HTTP
'           : request, when that emitter holds no callback.
'---------------------------------------------------------------------------------------
'
Private Function CancelPolled() As Boolean

    If m_lngFakeCancelOnCheck > 0 Then
        m_lngFakeCancelChecks = m_lngFakeCancelChecks + 1
        If m_lngFakeCancelChecks < m_lngFakeCancelOnCheck Then Exit Function
        SetCancelPollForTest 0
        CancelPolled = True
    ElseIf IsSessionSingleton Then
        If Not m_RootEmitter Is Nothing Then CancelPolled = m_RootEmitter.CheckCancelled
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CompletionFaultForTest
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : One-shot fault while constructing a completion payload.
'---------------------------------------------------------------------------------------
'
Friend Property Let CompletionFaultForTest(ByVal lngNumber As Long)
    m_lngCompletionFaultForTest = lngNumber
End Property


'---------------------------------------------------------------------------------------
' Procedure : SetCancelPollForTest
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Test seam. The cancel poll reports a request on the Nth check from now,
'           : once, in place of asking MCP. Pass 0 to disarm. Arm a private instance
'           : unless the test means to cancel the session's running operation.
'---------------------------------------------------------------------------------------
'
Friend Sub SetCancelPollForTest(ByVal lngCancelOnCheck As Long)
    m_lngFakeCancelOnCheck = IIf(lngCancelOnCheck > 0, lngCancelOnCheck, 0)
    m_lngFakeCancelChecks = 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsActive
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True while an export, build, merge, or test run has not finished.
'---------------------------------------------------------------------------------------
'
Public Function IsActive() As Boolean
    IsActive = (Me.Status = eosRunning Or Me.Status = eosStaged)
End Function


'---------------------------------------------------------------------------------------
' Procedure : StatusName
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Status as a string for JSON results: running, staged, or ready.
'---------------------------------------------------------------------------------------
'
Public Function StatusName() As String

    Select Case Me.Status
        Case eosRunning: StatusName = "running"
        Case eosStaged: StatusName = "staged"
        Case Else: StatusName = "ready"
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : Status
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Return the current status of the operation.
'---------------------------------------------------------------------------------------
'
Public Property Get Status() As eOperationState
    ' A paused root is sitting inside foreign code (a user hook such as AfterExport,
    ' or a form closing) that can legitimately run past the timeout and that we have
    ' no way to pulse through. Expiring it there would strand the root: EndPauseScope
    ' only restores a root it still finds staged.
    If m_Root.State <> eosReady And m_Root.PauseDepth = 0 And Not m_Root.RebuildHandoffHold Then
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, Me.Heartbeat) < Now Then
            FinishRootInternal eorTimeout
        End If
    End If
    Status = m_Root.State
End Property


'---------------------------------------------------------------------------------------
' Procedure : OperationType
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : The type of work the active root is doing.
'---------------------------------------------------------------------------------------
'
Public Property Get OperationType() As eOperationType
    OperationType = m_Root.OperationType
End Property


Public Property Let OperationType(intType As eOperationType)
    m_Root.OperationType = intType
End Property


'---------------------------------------------------------------------------------------
' Procedure : InteractionMode
' Purpose   : Controls UI prompts during the operation. eimNormal (0) shows MsgBox2
'           : dialogs; eimSilent (1) routes them to the log and returns the supplied
'           : default; eimNonInteractive (2) answers them from DecisionPolicy. Values are
'           : ordered by strictness (Normal < Silent < NonInteractive).
'           : Interaction can only get stricter while a root is active: a caller that
'           : asked for silence must not have it revoked by unrelated code deeper in the
'           : same operation. Between operations the value is just the default for the
'           : next root, so relaxing it is allowed - that is how a headless caller returns
'           : the session to normal. Inside a noninteractive scope, a request for a weaker
'           : mode is ignored.
'---------------------------------------------------------------------------------------
'
Public Property Get InteractionMode() As eInteractionMode
    InteractionMode = Scope.InteractionMode
End Property


Public Property Let InteractionMode(ByVal eimMode As eInteractionMode)

    If eimMode < Scope.InteractionMode _
        And (m_Root.State <> eosReady) And (Len(m_Root.Token) > 0) Then
        LogDiagnostic "Ignored attempt to relax interaction mode", "clsOperation.InteractionMode"
        Exit Property
    End If
    Scope.InteractionMode = eimMode

End Property


Public Property Get Attended() As Boolean
    Attended = m_Root.Attended
End Property


Public Property Get CurrentRootToken() As String
    CurrentRootToken = m_Root.Token
End Property


Public Property Get SourceName() As String
    Select Case Me.Source
        Case eosMCPTool:     SourceName = "MCP Tool"
        Case eosExternalAPI: SourceName = "External API"
    End Select
End Property


Public Property Get AutomationSource() As Boolean
    AutomationSource = (Me.Source = eosExternalAPI Or Me.Source = eosMCPTool)
End Property


Private Function CanBeginRoot(Optional ByVal intPolicy As eDecisionPolicy = edpAsk) As Boolean

    m_strAdmissionRefusal = vbNullString
    Select Case Me.Status
        Case eosRunning, eosStaged
            m_strAdmissionRefusal = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
            ReportRefusal T("Another Operation Already Running"), _
                T("Please complete the current operation before beginning a new one."), _
                (intPolicy <> edpAsk)
        Case eosReady
            CanBeginRoot = True
        Case Else
            ReportRefusal T("Unknown Operation State"), T("Current state: {0}", , , , m_Root.State), _
                (intPolicy <> edpAsk)
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReportRefusal
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Explain a refused root. This runs before a root exists, so the usual
'           : Attended/InteractionMode state describes the previous operation and cannot
'           : be trusted - decide from the incoming caller's own intent instead. An
'           : automated or explicitly unattended caller gets a log entry, never a dialog.
'           : A caller that passed a decision policy is automated (blnQuiet).
'---------------------------------------------------------------------------------------
'
Private Sub ReportRefusal(strTitle As String, strMessage As String, Optional ByVal blnQuiet As Boolean = False)

    If IsSessionSingleton _
        And Scope.InteractionMode = eimNormal _
        And Not blnQuiet _
        And Not (Me.AutomationSource Or Me.ForceUnattended) _
        And RequestSource = eosUserInterface Then
        MsgBox2 strTitle, strMessage, , vbExclamation
    Else
        If IsSessionSingleton Then
            SysCmd acSysCmdSetStatus, strTitle & ": " & strMessage
        Else
            Debug.Print strTitle & ": " & strMessage
        End If
    End If

End Sub


Private Sub InitializeRoot(intType As eOperationType, ByVal intPolicy As eDecisionPolicy, _
    ByVal cRequest As clsOperationRequest, ByVal blnUnattended As Boolean, ByVal intMode As Long)

    m_Root.Token = NewRootToken()
    m_Root.PublicId = cRequest.PublicId
    m_Root.InitiatingSource = cRequest.Source
    m_Root.Policy = intPolicy
    m_Root.CallbackInfo = cRequest.CallbackInfo
    m_Root.OperationType = intType
    m_Root.Attended = Not (Me.AutomationSource Or Me.ForceUnattended Or blnUnattended)
    m_Root.StartTime = Now
    m_Root.FinishTime = 0
    m_Root.State = eosRunning
    m_Root.Detached = False
    m_Root.Completed = False
    m_Root.RebuildHandoffHold = False
    m_Root.PauseDepth = 0
    m_Root.AppHwnd = hWndAccessApp
    m_Root.RuntimeError = vbNullString
    m_Root.RuntimeErrorNumber = 0
    m_Root.CancelRequested = False
    m_Root.OutputPath = vbNullString
    m_Root.NativeCloseCanceled = False
    Set m_TransferOutcome = Nothing
    Set m_PolicyComparisonRefusal = Nothing

    Me.ErrorLevel = eelNoError
    Me.Result = eorUnknown
    Set m_LastDecisions = Nothing
    Set m_LastCompletion = Nothing

    ' A test run's callback carries the results file, so a run that ends before it
    ' saves must not find the previous run's path still set.
    If intType = eotTestRun And IsSessionSingleton Then TestRunner.ClearLastResultsPath

    ' Every root owns a fresh journal. An omitted policy never inherits activation.
    If intPolicy <> edpAsk Then
        Scope.OpenScope eimNonInteractive, intPolicy, True
    Else
        If intMode < 0 Then intMode = Scope.InteractionMode
        If intMode = eimNonInteractive Then intMode = eimNormal
        Scope.OpenScope intMode, edpAsk, True
    End If
    If IsSessionSingleton Then
        Set m_RootEmitter = MCP
        m_RootEmitter.RegisterCallback m_Root.CallbackInfo
    End If

    SetErrorTrapping
    UpdateRegistry
    Pulse

End Sub


Private Sub FinishRootInternal(Optional intResult As eOperationResult = eorUnknown)

    Dim blnWasRunning As Boolean
    Dim dExtra As Dictionary
    Dim cCallback As clsMCP
    Dim dPartial As Dictionary
    Dim varKey As Variant
    Dim strType As String
    Dim strMessage As String
    Dim lngError As Long
    Dim strError As String
    Dim lngFault As Long

    If m_Root.Completed Then Exit Sub
    If m_Root.State = eosReady And Len(m_Root.Token) = 0 Then Exit Sub

    blnWasRunning = (m_Root.State = eosRunning)
    LogUnhandledErrors
    On Error GoTo PayloadError

    ' An unresolved prompt is not a successful or merely cancelled run.
    If Not ApplyBlockedResult() Then
        If intResult > eorUnknown Then Me.Result = intResult
        ' A cancel request is a result only when no runtime error took precedence.
        ' A test run keeps the result its runner reports, with the partial results.
        If m_Root.CancelRequested And m_Root.OperationType <> eotTestRun _
            And Len(m_Root.RuntimeError) = 0 Then Me.Result = eorCanceled
        If Me.Result = eorUnknown Then
            Me.Result = IIf(Me.ErrorLevel >= eelError, eorFailed, eorSuccess)
        End If
    End If

    ' Claim completion before any code that can yield or fail. Cleanup below is
    ' unconditional: a completed flag alone cannot release the root's owned state.
    m_Root.Completed = True
    m_Root.State = eosReady
    m_Root.FinishTime = Now
    EnsureRootEmitter
    Set cCallback = m_RootEmitter
    Set m_LastDecisions = Scope.Decisions

    Set dExtra = New Dictionary
    If Not m_TransferOutcome Is Nothing Then
        For Each varKey In m_TransferOutcome.Keys
            dExtra.Add varKey, m_TransferOutcome(varKey)
        Next varKey
    End If
    If Not m_PolicyComparisonRefusal Is Nothing Then
        For Each varKey In m_PolicyComparisonRefusal.Keys
            If dExtra.Exists(varKey) Then dExtra.Remove varKey
            dExtra.Add varKey, m_PolicyComparisonRefusal(varKey)
        Next varKey
    End If
    AddOperationIdentity dExtra, Me
    dExtra("error_level") = Me.ErrorLevel
    dExtra("success") = (Me.Result = eorSuccess)
    dExtra("log_path") = Log.SavedLogFilePath
    If m_Root.OperationType = eotTestRun Then
        If Len(TestRunner.LastResultsPath) > 0 Then
            dExtra("results_path") = TestRunner.LastResultsPath
        End If
        If Len(TestRunner.RunError) > 0 Then
            dExtra("success") = False
            dExtra("cancelled") = False
            dExtra("error") = TestRunner.RunError
            dExtra("error_pattern") = TestRunner.RunErrorPattern
        End If
        If Len(TestRunner.ResultsSaveError) > 0 Then
            dExtra("results_error") = TestRunner.ResultsSaveError
        End If
    End If
    If m_LastDecisions.Count > 0 Then Set dExtra("decisions") = m_LastDecisions
    If Len(m_Root.RuntimeError) > 0 Then
        dExtra("runtime_error") = m_Root.RuntimeError
        dExtra("errorNumber") = m_Root.RuntimeErrorNumber
    End If
    If Len(m_Root.OutputPath) > 0 Then dExtra("output_path") = m_Root.OutputPath

    ' A saved file already preserves results. Only carry them inline when that
    ' file is unavailable. Capture the small diagnostic payload before fallible
    ' JSON construction, so even its failure keeps the operation's own evidence.
    If m_Root.OperationType = eotTestRun Then
        If Len(TestRunner.LastResultsPath) = 0 And _
            (TestRunner.State = etrsComplete Or TestRunner.State = etrsCancelled _
            Or TestRunner.State = etrsFailed) Then
            Set dPartial = ParseJson(TestRunner.GetResultsAsJson)
            For Each varKey In dPartial.Keys
                If Not dExtra.Exists(varKey) Then dExtra.Add varKey, dPartial(varKey)
            Next varKey
        End If
    End If

    Select Case Me.Result
        Case eorSuccess
            strType = "complete"
            strMessage = T("Operation completed successfully")
        Case eorFailed
            strType = "error"
            strMessage = T("Operation failed")
            If Len(m_Root.RuntimeError) > 0 Then strMessage = m_Root.RuntimeError
            If dExtra.Exists("error") Then strMessage = dExtra("error")
        Case eorCanceled
            strType = "cancelled"
            strMessage = T("Operation was cancelled")
        Case eorTimeout
            strType = "error"
            strMessage = T("Operation timed out")
        Case eorDecisionRequired
            strType = "error"
            SetDecisionRequiredOutcome dExtra, m_LastDecisions
            strMessage = dExtra("error")
        Case Else
            strType = "complete"
            strMessage = T("Operation completed")
    End Select
    If strType <> "complete" Then dExtra("success") = False

    If m_lngCompletionFaultForTest <> 0 Then
        lngFault = m_lngCompletionFaultForTest
        m_lngCompletionFaultForTest = 0
        Err.Raise lngFault, "clsOperation.FinishRootInternal", "Completion payload fixture"
    End If
    GoTo CleanUp

PayloadError:
    lngError = Err.Number
    strError = Err.Description
    Resume PayloadFailed

PayloadFailed:
    ' Keep fields already captured. A completion fault is secondary to the
    ' operation's decision/compile/runtime error, and cannot skip restoration.
    On Error Resume Next
    If dExtra Is Nothing Then Set dExtra = New Dictionary
    dExtra("completion_error") = strError
    dExtra("completion_error_number") = lngError
    dExtra("success") = False
    strType = "error"
    strMessage = strError
    If Me.Result = eorDecisionRequired Then
        SetDecisionRequiredOutcome dExtra, m_LastDecisions
        strMessage = dExtra("error")
    Else
        Me.Result = eorFailed
        If Not dExtra.Exists("error") Then dExtra("error") = strError
    End If

CleanUp:
    ' Each restoration is attempted even if another one or callback submission
    ' fails. Retain the admitted emitter while releasing shared infrastructure.
    On Error Resume Next
    m_Root.Completed = True
    m_Root.State = eosReady
    m_Root.FinishTime = Now
    Set m_cSyncLease = Nothing
    If m_LastDecisions Is Nothing Then Set m_LastDecisions = Scope.Decisions
    Scope.ReleaseOperationScopes
    If IsSessionSingleton Then ReleaseTimerReservation m_Root.Token
    Set m_RootEmitter = Nothing
    Me.ForceUnattended = False
    Scope.InteractionMode = eimNormal
    m_Root.Token = vbNullString
    m_Root.RebuildHandoffHold = False
    m_Root.PauseDepth = 0
    m_Root.Detached = False
    UpdateRegistry
    If blnWasRunning Then RestoreUserErrorTrapping
    If IsSessionSingleton Then
        If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagBegin "op.release"
        modObjects.ReleaseObjects
        If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagEnd "op.release"
    End If

    Set m_LastCompletion = dExtra
    If Not cCallback Is Nothing Then
        If strType = "complete" Then
            cCallback.PostCallback strType, 100, 100, strMessage, dExtra
        Else
            cCallback.PostCallback strType, -1, -1, strMessage, dExtra
        End If
    End If
    ' Reserved envelope keys belong to PostCallback, so add them only afterwards.
    dExtra("type") = strType
    dExtra("message") = strMessage
    Err.Clear

End Sub


Private Function NewRootToken() As String
    NewRootToken = NewOperationIdentity()
End Function


Private Sub LogDiagnostic(strMessage As String, strSource As String)
    On Error Resume Next
    Log.Add "[Operation] " & strMessage & IIf(Len(strSource) > 0, " (" & strSource & ")", vbNullString), False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsSessionSingleton
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : True only for the one instance published by modObjects. Everything that
'           : reaches outside this instance - the registry state used for crash recovery,
'           : the MCP completion callback, VBE error trapping, and singleton teardown -
'           : is gated on this so a test can drive a private instance through a full root
'           : lifecycle without disturbing the operation that is actually running.
'---------------------------------------------------------------------------------------
'
Private Function IsSessionSingleton() As Boolean
    On Error Resume Next
    IsSessionSingleton = (Me Is modObjects.Operation)
End Function


Private Sub UpdateRegistry()
    If Not IsSessionSingleton Then Exit Sub
    If m_Root.AppHwnd <> 0 Then
        CheckSetting "Heartbeat", Me.Heartbeat
        CheckSetting "Result", Me.Result
        CheckSetting "StartTime", m_Root.StartTime
        CheckSetting "FinishTime", m_Root.FinishTime
        CheckSetting "State", m_Root.State
        CheckSetting "AppHwnd", m_Root.AppHwnd
        CheckSetting "ErrorTrapping", m_Root.ErrorTrapping
        CheckSetting REG_KEY_ROOT_TOKEN, m_Root.Token
        CheckSetting REG_KEY_ROOT_TYPE, m_Root.OperationType
        CheckSetting "PublicOperationId", m_Root.PublicId
        CheckSetting "InitiatingSource", m_Root.InitiatingSource
        CheckSetting "Attended", m_Root.Attended
        CheckSetting "RequestPolicy", m_Root.Policy
        CheckSetting "CallbackInfo", m_Root.CallbackInfo
        CheckSetting "RebuildHandoffHold", m_Root.RebuildHandoffHold
        CheckSetting "DecisionJournal", ConvertToJson(Scope.Decisions)
        CheckSetting "DecisionBlocked", Scope.Blocked
        CheckSetting REG_KEY_ROOT_PROJECT, CodeProject.FullName
        CheckSetting "InteractionMode", Scope.InteractionMode
        CheckSetting "DecisionPolicy", Scope.DecisionPolicy
    End If
End Sub


Private Sub CheckSetting(strKey As String, varValue)
    If GetSetting(PROJECT_NAME, "State", strKey) <> CStr(varValue) Then
        SaveSetting PROJECT_NAME, "State", strKey, CStr(varValue)
    End If
End Sub


Private Function RestoreFromRegistry() As Boolean

    If CanAdoptRegistryState(ReadReg("AppHwnd"), ReadReg("Heartbeat"), ReadReg(REG_KEY_ROOT_PROJECT), _
        (ReadReg("RebuildHandoffHold") = "True")) Then
        ' The heartbeat came from the same registry value the timeout check just
        ' accepted; without it Status would see a zero heartbeat and report the
        ' restored operation as timed out.
        Me.Heartbeat = CDate(ReadReg("Heartbeat"))
        Me.Result = ReadReg("Result")
        m_Root.StartTime = ReadReg("StartTime")
        m_Root.FinishTime = ReadReg("FinishTime")
        m_Root.State = ReadReg("State")
        m_Root.AppHwnd = ReadReg("AppHwnd")
        m_Root.ErrorTrapping = CInt(ReadReg("ErrorTrapping"))
        m_Root.Token = ReadReg(REG_KEY_ROOT_TOKEN)
        m_Root.OperationType = Val(ReadReg(REG_KEY_ROOT_TYPE))
        m_Root.PublicId = ReadReg("PublicOperationId")
        m_Root.InitiatingSource = Val(ReadReg("InitiatingSource"))
        m_Root.Attended = (ReadReg("Attended") = "True")
        m_Root.Policy = Val(ReadReg("RequestPolicy"))
        m_Root.CallbackInfo = ReadReg("CallbackInfo")
        m_Root.RebuildHandoffHold = (ReadReg("RebuildHandoffHold") = "True")
        m_Root.Completed = False
        RestoreScopeFromRegistry
        Scope.RestoreJournal ReadReg("DecisionJournal"), (ReadReg("DecisionBlocked") = "True")
        RestoreFromRegistry = True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CanAdoptRegistryState
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : True when the saved state belongs to this Access window, has a heartbeat
'           : inside the timeout, and was written by this add-in file. Every copy of the
'           : add-in shares one registry key, so a development copy loaded beside the
'           : installed add-in would otherwise adopt the installed add-in's running test
'           : run - one it can never finish - and refuse to close frmVCSMain until the
'           : heartbeat timed out. State saved before the project was recorded is
'           : not adopted.
'---------------------------------------------------------------------------------------
'
Friend Function CanAdoptRegistryState(ByVal strAppHwnd As String, ByVal strHeartbeat As String, _
    ByVal strProject As String, Optional ByVal blnRebuildHandoffHold As Boolean = False) As Boolean

    If strAppHwnd <> CStr(hWndAccessApp) Then Exit Function
    If Not IsDate(strHeartbeat) Then Exit Function
    If Not blnRebuildHandoffHold Then
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, CDate(strHeartbeat)) <= Now Then Exit Function
    End If
    CanAdoptRegistryState = (StrComp(strProject, CodeProject.FullName, vbTextCompare) = 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : RestoreScopeFromRegistry
' Date      : 09/29/2026
' Purpose   : Put back the interaction mode and policy saved with an operation that is
'           : still running or staged, so the second half of a timer-driven operation
'           : runs under the same policy. A noninteractive policy the live scope already
'           : holds is left alone. Otherwise it reopens as an operation-owned scope, so
'           : Finish still releases it.
'---------------------------------------------------------------------------------------
'
Private Sub RestoreScopeFromRegistry()

    Dim intMode As eInteractionMode

    If m_Root.State = eosReady Then Exit Sub
    intMode = Val(ReadReg("InteractionMode"))

    If intMode = eimNonInteractive Then
        If Not Scope.InNonInteractiveScope Then
            Scope.OpenScope eimNonInteractive, Val(ReadReg("DecisionPolicy")), True
        End If
    Else
        Scope.InteractionMode = intMode
    End If

End Sub


Private Function ReadReg(strKey As String) As String
    ReadReg = GetSetting(PROJECT_NAME, "State", strKey)
End Function


Private Sub SetErrorTrapping()
    If Not IsSessionSingleton Then Exit Sub
    m_Root.ErrorTrapping = SaveUserErrorTrapping()
    ApplyUserErrorTrapping EffectiveVbeErrorTrapping(m_Root.ErrorTrapping, eetBreakInClassModule)
End Sub


Private Sub RestoreUserErrorTrapping()
    If Not IsSessionSingleton Then Exit Sub
    modErrorHandling.RestoreUserErrorTrapping m_Root.ErrorTrapping
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Scope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : The interaction scope object this operation delegates to. Callers use
'           : the members below rather than the scope, which stays private so the
'           : scope can change without touching them. The property-style members
'           : that only forward (DecisionPolicy, DecisionBlocked, ErrorLevel,
'           : PushInteractionScope, CloseInteractionScope, Decisions) are the
'           : operation's contract: the export, build and merge code reads and sets
'           : them on Operation, and the API class and modDialogPolicy open and close
'           : scopes through it. The members that did nothing else, such as
'           : ResolveConflictAction, were removed.
'---------------------------------------------------------------------------------------
'
Private Property Get Scope() As clsInteractionScope
    If m_Scope Is Nothing Then Set m_Scope = New clsInteractionScope
    Set Scope = m_Scope
End Property


'---------------------------------------------------------------------------------------
' Procedure : DecisionPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Policy used while InteractionMode is eimNonInteractive.
'---------------------------------------------------------------------------------------
'
Public Property Get DecisionPolicy() As eDecisionPolicy
    DecisionPolicy = Scope.DecisionPolicy
End Property

Public Property Let DecisionPolicy(ByVal intPolicy As eDecisionPolicy)
    Scope.DecisionPolicy = intPolicy
End Property


'---------------------------------------------------------------------------------------
' Procedure : DecisionBlocked
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Set when a required prompt was not covered by DecisionPolicy.
'           : Finish reports eorDecisionRequired and does not report success.
'---------------------------------------------------------------------------------------
'
Public Property Get DecisionBlocked() As Boolean
    DecisionBlocked = Scope.Blocked
End Property

Public Property Let DecisionBlocked(ByVal blnBlocked As Boolean)
    Scope.Blocked = blnBlocked
End Property


'---------------------------------------------------------------------------------------
' Procedure : ErrorLevel
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Highest error severity seen by this operation. Saved and restored by
'           : interaction scopes.
'---------------------------------------------------------------------------------------
'
Public Property Get ErrorLevel() As eErrorLevel
    ErrorLevel = Scope.ErrorLevel
End Property

Public Property Let ErrorLevel(ByVal intLevel As eErrorLevel)
    Scope.ErrorLevel = intLevel
End Property


'---------------------------------------------------------------------------------------
' Procedure : PushInteractionScope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Open an interaction scope and return its owner token. blnPopOnFinish
'           : makes it operation-owned: Finish closes it, including after failure
'           : and cancellation. Otherwise only CloseInteractionScope closes it.
'---------------------------------------------------------------------------------------
'
Public Function PushInteractionScope(ByVal intMode As eInteractionMode, _
    ByVal intPolicy As eDecisionPolicy, _
    Optional ByVal blnPopOnFinish As Boolean = True) As Long
    PushInteractionScope = Scope.OpenScope(intMode, intPolicy, blnPopOnFinish)
End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseInteractionScope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Close the scope for this token, restoring the state saved when it
'           : opened. An unknown or already closed token does nothing.
'---------------------------------------------------------------------------------------
'
Public Sub CloseInteractionScope(ByVal lngToken As Long)
    Scope.CloseScope lngToken
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolvePrompt
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Answer a message box from the current policy and record the outcome.
'---------------------------------------------------------------------------------------
'
Public Function ResolvePrompt(ByVal intButtons As Long, ByVal strTitle As String, _
    ByVal strMessage As String) As clsPromptOutcome
    Set ResolvePrompt = Scope.ResolvePrompt(intButtons, strTitle, strMessage)
    ApplyBlockedResult
End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordDecision
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Append one prompt outcome. decision_required blocks the operation
'           : and prevents Finish from reporting success.
'---------------------------------------------------------------------------------------
'
Public Sub RecordDecision(ByVal cRecord As clsDecisionRecord)

    Scope.RecordDecision cRecord
    ApplyBlockedResult
    UpdateRegistry

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyBlockedResult
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : The one place a blocked prompt becomes eorDecisionRequired. Called when a
'           : prompt is recorded, because closing a nested scope restores the outer
'           : scope's blocked flag and would otherwise lose the block, and again by
'           : Finish. Returns True when the operation is blocked.
'---------------------------------------------------------------------------------------
'
Private Function ApplyBlockedResult() As Boolean
    If Scope.Blocked Then
        Me.Result = eorDecisionRequired
        ApplyBlockedResult = True
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Decisions
' Date      : 09/29/2026
' Purpose   : Prompts recorded in the current scope, as a collection of dictionaries
'           : for the JSON encoder.
'---------------------------------------------------------------------------------------
'
Public Function Decisions() As Collection
    Set Decisions = Scope.Decisions
End Function


'---------------------------------------------------------------------------------------
' Procedure : LastDecisions
' Date      : 09/29/2026
' Purpose   : Prompts recorded by the operation that most recently finished. Finish
'           : keeps them because releasing the operation's scope discards the journal.
'           : Nothing until an operation has finished.
'---------------------------------------------------------------------------------------
'
Public Function LastDecisions() As Collection
    Set LastDecisions = m_LastDecisions
End Function


'---------------------------------------------------------------------------------------
' Procedure : LastCompletion
' Date      : 09/30/2026
' Purpose   : Terminal callback fields of the operation that most recently finished,
'           : plus its "type" and "message". Built whether or not MCP is active, so the
'           : payload can be checked on a private instance. Nothing until one finishes.
'---------------------------------------------------------------------------------------
'
Public Function LastCompletion() As Dictionary
    Set LastCompletion = m_LastCompletion
End Function


Private Sub Class_Initialize()
    m_Root.State = eosReady
    m_Root.Attended = True
    If modObjects.CreatingOperationSingleton Then
        If RestoreFromRegistry Then Exit Sub
    End If
    m_Root.AppHwnd = Application.hWndAccessApp
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Source
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Read captured owner source or configure only an idle direct caller.
'---------------------------------------------------------------------------------------
'
Public Property Get Source() As eOperationSource
    If Len(m_Root.Token) > 0 Then
        Source = m_Root.InitiatingSource
    ElseIf IsSessionSingleton And Not IncomingOperationRequest Is Nothing Then
        Source = RequestSource
    Else
        Source = m_intIdleSource
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : Source
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Read captured owner source or configure only an idle direct caller.
'---------------------------------------------------------------------------------------
'
Public Property Let Source(ByVal intSource As eOperationSource)
    ' Compatibility for internal callers/tests. Dispatch uses its own request object.
    If IsActive Then Exit Property
    m_intIdleSource = intSource
End Property


'---------------------------------------------------------------------------------------
' Procedure : PublicId
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Return the stable admitted public identity independently of callbacks.
'---------------------------------------------------------------------------------------
'
Public Property Get PublicId() As String
    PublicId = m_Root.PublicId
End Property


'---------------------------------------------------------------------------------------
' Procedure : PolicyName
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Report the captured policy, including an explicitly absent policy.
'---------------------------------------------------------------------------------------
'
Public Property Get PolicyName() As String
    PolicyName = OperationPolicyName(m_Root.Policy)
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceLabel
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Report the captured initiating source for public metadata.
'---------------------------------------------------------------------------------------
'
Public Property Get SourceLabel() As String
    Select Case m_Root.InitiatingSource
        Case eosMCPTool: SourceLabel = "mcp"
        Case eosExternalAPI: SourceLabel = "api"
        Case Else: SourceLabel = "user_interface"
    End Select
End Property


'---------------------------------------------------------------------------------------
' Procedure : AdmissionRefusal
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Return the last request-local root-admission refusal.
'---------------------------------------------------------------------------------------
'
Public Property Get AdmissionRefusal() As String
    AdmissionRefusal = m_strAdmissionRefusal
    If Len(AdmissionRefusal) = 0 Then AdmissionRefusal = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
End Property


'---------------------------------------------------------------------------------------
' Procedure : SetTransferOutcome
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Retain completed/preserved transfer counts for final callbacks.
'---------------------------------------------------------------------------------------
'
Public Sub SetTransferOutcome(ByVal dTransfer As Dictionary)
    Dim varKey As Variant
    Set m_TransferOutcome = New Dictionary
    If dTransfer Is Nothing Then Exit Sub
    For Each varKey In dTransfer.Keys
        Select Case CStr(varKey)
            Case "success", "decisions", "operation_id", "type", "message", "decision_policy", "initiating_source", "error_level"
                ' The final operation, rather than a component, owns these fields.
            Case Else
                m_TransferOutcome.Add varKey, dTransfer(varKey)
        End Select
    Next varKey
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreRootInteraction
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Restore launcher interaction only while its exact root is retained.
'---------------------------------------------------------------------------------------
'
Public Sub RestoreRootInteraction(ByVal strToken As String, ByVal intMode As eInteractionMode)
    If Not RootLeaseIsActive(strToken) Then Exit Sub
    Scope.InteractionMode = intMode
    UpdateRegistry
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnsureRootEmitter
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Recreate callback registration only for the matching resumed owner.
'---------------------------------------------------------------------------------------
'
Private Sub EnsureRootEmitter()
    If Not IsSessionSingleton Then Exit Sub
    If Not m_RootEmitter Is Nothing Then Exit Sub
    Set m_RootEmitter = MCP
    m_RootEmitter.RegisterCallback m_Root.CallbackInfo
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestorePersistedRootForTest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Restore owned registry state on a private disposable probe instance.
'---------------------------------------------------------------------------------------
'
Friend Function RestorePersistedRootForTest() As Boolean
    RestorePersistedRootForTest = RestoreFromRegistry()
End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordPolicyComparisonRefusal
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Retain specific comparison diagnostics across generic transfer summaries.
'---------------------------------------------------------------------------------------
'
Public Sub RecordPolicyComparisonRefusal(ByVal dRefusal As Dictionary)
    Dim varKey As Variant
    Set m_PolicyComparisonRefusal = New Dictionary
    For Each varKey In dRefusal.Keys
        m_PolicyComparisonRefusal.Add varKey, dRefusal(varKey)
    Next varKey
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PolicyComparisonRefusal
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Expose the owned comparison refusal for synchronous final results.
'---------------------------------------------------------------------------------------
'
Public Property Get PolicyComparisonRefusal() As Dictionary
    Set PolicyComparisonRefusal = m_PolicyComparisonRefusal
End Property


'---------------------------------------------------------------------------------------
' Procedure : TransferOutcome
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Expose the current owned transfer outcome to final-result assembly.
'---------------------------------------------------------------------------------------
'
Public Property Get TransferOutcome() As Dictionary
    Set TransferOutcome = m_TransferOutcome
End Property

' A launched helper can still act after a handoff or quit failure. Heartbeat expiry
' cannot prove otherwise. Recovery must explicitly complete this same owned root.

'---------------------------------------------------------------------------------------
' Procedure : RetainRebuildHandoff
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Prevent heartbeat expiry from releasing a launched helper reservation.
'---------------------------------------------------------------------------------------
'
Public Sub RetainRebuildHandoff(ByVal strToken As String)
    If Not RootLeaseIsActive(strToken) Then Exit Sub
    m_Root.RebuildHandoffHold = True
    UpdateRegistry
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RebuildHandoffHeld
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Report a root retained pending explicit rebuild recovery.
'---------------------------------------------------------------------------------------
'
Public Property Get RebuildHandoffHeld() As Boolean
    RebuildHandoffHeld = m_Root.RebuildHandoffHold
End Property


Public Property Get NativeCloseCanceled() As Boolean
    NativeCloseCanceled = m_Root.NativeCloseCanceled
End Property


Public Property Let NativeCloseCanceled(ByVal blnCanceled As Boolean)
    m_Root.NativeCloseCanceled = blnCanceled
End Property

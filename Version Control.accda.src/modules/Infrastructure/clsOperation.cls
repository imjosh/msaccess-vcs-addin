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


Public Heartbeat As Date
Public Result As eOperationResult
Public Source As eOperationSource

' Set by headless entry points before TryBeginRoot/Begin so Attended stays False even
' when Source is still eosUserInterface (Immediate Window callers).
Public ForceUnattended As Boolean


Private Type udtRoot
    Token As String
    OperationType As eOperationType
    Attended As Boolean
    StartTime As Date
    FinishTime As Date
    State As eOperationState
    AppHwnd As Long
    ErrorTrapping As eVbeErrorTrapping
    Detached As Boolean
    Completed As Boolean
    PauseDepth As Long
    RuntimeError As String
    RuntimeErrorNumber As Long
End Type
Private m_Root As udtRoot

' Owns interaction mode, decision policy, blocked flag, decision journal and error level.
Private m_Scope As clsInteractionScope

' Owner token of the scope OpenSessionScope opened. Zero when none is open.
Private m_SessionScope As Long

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
Public Function TryBeginRoot(intType As eOperationType, Optional ByVal intPolicy As eDecisionPolicy = edpAsk) As clsRootOperationLease

    Dim cLease As clsRootOperationLease

    If Not CanBeginRoot(intPolicy) Then Exit Function

    InitializeRoot intType, intPolicy
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
Public Function Begin(intType As eOperationType, Optional ByVal intPolicy As eDecisionPolicy = edpAsk) As Boolean

    Set m_cSyncLease = TryBeginRoot(intType, intPolicy)
    Begin = (Not m_cSyncLease Is Nothing)

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
    If m_Root.State <> eosReady And m_Root.PauseDepth = 0 Then
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, Me.Heartbeat) < Now Then
            m_Root.State = eosReady
            Me.Result = eorTimeout
            Scope.ReleaseOperationScopes
            RestoreUserErrorTrapping
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

    Select Case m_Root.State
        Case eosRunning, eosStaged
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
        And Not (Me.AutomationSource Or Me.ForceUnattended) Then
        MsgBox2 strTitle, strMessage, , vbExclamation
    Else
        LogDiagnostic strTitle & ": " & strMessage, "clsOperation.CanBeginRoot"
    End If

End Sub


Private Sub InitializeRoot(intType As eOperationType, Optional ByVal intPolicy As eDecisionPolicy = edpAsk)

    m_Root.Token = NewRootToken()
    m_Root.OperationType = intType
    m_Root.Attended = Not (Me.AutomationSource Or Me.ForceUnattended)
    m_Root.StartTime = Now
    m_Root.FinishTime = 0
    m_Root.State = eosRunning
    m_Root.Detached = False
    m_Root.Completed = False
    m_Root.PauseDepth = 0
    m_Root.AppHwnd = hWndAccessApp
    m_Root.RuntimeError = vbNullString
    m_Root.RuntimeErrorNumber = 0

    Me.ErrorLevel = eelNoError
    Me.Result = eorUnknown
    Set m_LastDecisions = Nothing
    Set m_LastCompletion = Nothing

    ' A test run's callback carries the results file, so a run that ends before it
    ' saves must not find the previous run's path still set.
    If intType = eotTestRun And IsSessionSingleton Then TestRunner.ClearLastResultsPath

    ' The policy scope opens only now that the root is granted. Finish closes it, and
    ' UpdateRegistry saves it so RestoreFromRegistry can reopen it. Under a caller-owned
    ' session policy the root still opens its own scope, with that policy, so its
    ' blocked flag and decisions start empty and do not carry into the next root.
    If intPolicy <> edpAsk Then
        Scope.OpenScope eimNonInteractive, intPolicy, True
    ElseIf Scope.InNonInteractiveScope Then
        Scope.OpenScope eimNonInteractive, Scope.DecisionPolicy, True
    End If

    SetErrorTrapping
    UpdateRegistry
    Pulse

End Sub


Private Sub FinishRootInternal(Optional intResult As eOperationResult = eorUnknown)

    Dim blnWasRunning As Boolean
    Dim dExtra As Dictionary
    Dim strType As String
    Dim strMessage As String

    If m_Root.Completed Then Exit Sub
    If m_Root.State = eosReady And Len(m_Root.Token) = 0 Then Exit Sub

    ' An unresolved prompt is not a successful or merely cancelled run.
    If Not ApplyBlockedResult() And intResult > eorUnknown Then Me.Result = intResult
    blnWasRunning = (m_Root.State = eosRunning)

    m_Root.Completed = True
    m_Root.State = eosReady
    m_Root.FinishTime = Now

    ' Release any lease Begin was holding. The root is already flagged complete, so
    ' the lease cannot abandon it on the way out.
    Set m_cSyncLease = Nothing

    If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagBegin "op.registry"
    UpdateRegistry
    If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagEnd "op.registry"

    Set dExtra = New Dictionary
    dExtra.Add "log_path", Log.SavedLogFilePath
    If Scope.DecisionCount > 0 Then dExtra.Add "decisions", Scope.Decisions
    ' Failed tests complete as eorFailed (MCP "error"), so the results file
    ' has to ride on every terminal callback or the client only sees
    ' "Operation failed" with no JSON to rerun from.
    If m_Root.OperationType = eotTestRun Then
        If Len(TestRunner.LastResultsPath) > 0 Then
            dExtra.Add "results_path", TestRunner.LastResultsPath
        End If
        ' A run that finished but could not save its results file says why.
        If Len(TestRunner.ResultsSaveError) > 0 Then
            dExtra.Add "results_error", TestRunner.ResultsSaveError
        End If
    End If
    ' The error text is otherwise only in the sync return, which an async caller
    ' never reads.
    If Len(m_Root.RuntimeError) > 0 Then
        dExtra.Add "runtime_error", m_Root.RuntimeError
        dExtra.Add "errorNumber", m_Root.RuntimeErrorNumber
    End If

    Select Case Me.Result
        Case eorSuccess
            strType = "complete"
            strMessage = "Operation completed successfully"
        Case eorFailed
            strType = "error"
            strMessage = "Operation failed"
            If Len(m_Root.RuntimeError) > 0 Then strMessage = m_Root.RuntimeError
        Case eorCanceled
            strType = "cancelled"
            strMessage = "Operation was cancelled"
        Case eorTimeout
            strType = "error"
            strMessage = "Operation timed out"
        Case eorDecisionRequired
            strType = "error"
            strMessage = "Decision required"
            dExtra.Add "decision_required", True
            dExtra.Add "error_pattern", ERR_DECISION_REQUIRED
        Case Else
            strType = "complete"
            strMessage = "Operation completed"
    End Select
    If MCP.IsActive And IsSessionSingleton Then
        If strType = "complete" Then
            MCP.PostCallback strType, 100, 100, strMessage, dExtra
        Else
            MCP.PostCallback strType, -1, -1, strMessage, dExtra
        End If
    End If
    dExtra.Add "type", strType
    dExtra.Add "message", strMessage
    Set m_LastCompletion = dExtra

    ' Restore interaction state after the callback has read the journal.
    Set m_LastDecisions = Scope.Decisions
    Scope.ReleaseOperationScopes
    UpdateRegistry

    If IsSessionSingleton Then
        If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagBegin "op.release"
        modObjects.ReleaseObjects
        If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagEnd "op.release"
    End If

    If blnWasRunning Then RestoreUserErrorTrapping

    Me.ForceUnattended = False
    ' Each root starts from normal unless its caller asks otherwise. A caller-owned
    ' noninteractive scope that is still open ignores this.
    Scope.InteractionMode = eimNormal

    m_Root.Token = vbNullString

End Sub


Private Function NewRootToken() As String
    NewRootToken = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Timer * 1000000@)
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

    If ReadReg("AppHwnd") = CStr(hWndAccessApp) Then
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, CDate(Nz2(ReadReg("Heartbeat"), 0))) > Now Then
            ' The heartbeat came from the same registry value the timeout check above
            ' just accepted; without it Status would see a zero heartbeat and report the
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
            m_Root.Completed = False
            RestoreScopeFromRegistry
            RestoreFromRegistry = True
        End If
    End If

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
' Procedure : OpenSessionScope
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Open the caller-owned noninteractive scope a session policy runs under,
'           : replacing the one already open rather than nesting under it. The
'           : operation keeps the token, since API runs each call on a new
'           : clsVersionControl and no caller lives from the set to the clear.
'---------------------------------------------------------------------------------------
'
Public Sub OpenSessionScope(ByVal intPolicy As eDecisionPolicy)
    CloseSessionScope
    m_SessionScope = Scope.OpenScope(eimNonInteractive, intPolicy, False)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloseSessionScope
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Close the scope OpenSessionScope opened. Safe when none is open, and
'           : safe to call twice.
'---------------------------------------------------------------------------------------
'
Public Sub CloseSessionScope()
    Scope.CloseScope m_SessionScope
    m_SessionScope = 0
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

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
' Purpose   : This class manages the overall state of user operations such as Export,
'           : Merge, Import, custom action, etc...
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


' This constant determines whether the current operation is still valid or if it is
' considered to have timed out. (Or exited without cleaning up properly.)
' The heartbeat is incremented periodically during the export/build process.
Private Const HEARTBEAT_TIMEOUT_SECONDS As Integer = 60 * 10    ' (10 minutes)


' Periodically update this to Now() to manage timeouts.
Public Heartbeat As Date
Public OperationType As eOperationType
Public Result As eOperationResult
Public Source As eOperationSource     ' Default: eosUserInterface (0)

' Use a private type to manage local variables
Private Type udtThis
    StartTime As Date
    FinishTime As Date
    State As eOperationState

    ' Since multiple Access instances may be running, we need to
    ' track which instance initiated the operation.
    ' (We don't want multiple operations happening on different
    ' instances at the same time, since we also interact with
    ' potentially shared registry values.)
    AppHwnd As Long

    ' Other variables we are caching
    ErrorTrapping As Integer
End Type
Private this As udtThis

' Owns interaction mode, decision policy, blocked flag, decision journal and error level.
Private m_Scope As clsInteractionScope

' Journal of the operation that finished last (see LastDecisions).
Private m_LastDecisions As Collection


'---------------------------------------------------------------------------------------
' Procedure : Pulse
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Pulse the heartbeat
'---------------------------------------------------------------------------------------
'
Public Function Pulse()
    Heartbeat = Now
End Function


'---------------------------------------------------------------------------------------
' Procedure : Begin
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Begin an operation (I.e. Export, Merge, Import, custom action, etc...)
'           : Returns true if we are safe to begin a new operation. (No forms open,
'           : no operations currently in progress.)
'           :
'           : To run an operation unattended, set Me.InteractionMode = eimSilent
'           : on the singleton before (or any time during) Begin. Like Source,
'           : it is sticky -- callers that need to restore the prior value are
'           : responsible for caching and restoring it themselves.
'---------------------------------------------------------------------------------------
'
Public Function Begin(intType As eOperationType, Optional ByVal intPolicy As eDecisionPolicy = edpAsk) As Boolean
    Select Case this.State
        Case eosRunning, eosStaged
            ' Another operation is already running (or staged for a timer restart).
            ' A policy request comes from automation, so refuse without a dialog and
            ' leave the running operation's interaction state untouched.
            If intPolicy = edpAsk Then
                MsgBox2 T("Another Operation Already Running"), _
                    T("Please complete the current operation before beginning a new one."), , vbExclamation
            End If
        Case eosReady
            ' Ready to go!
            Me.OperationType = intType
            Me.ErrorLevel = eelNoError
            Me.Result = eorUnknown
            this.State = eosRunning
            this.StartTime = Now
            Set m_LastDecisions = Nothing
            ' The policy scope opens only now that the operation has begun. Finish
            ' closes it, and UpdateRegistry saves it so Restore can reopen it.
            If intPolicy <> edpAsk Then Scope.OpenScope eimNonInteractive, intPolicy, True
            SetErrorTrapping
            UpdateRegistry
            Me.Pulse
            Begin = True
        Case Else
            MsgBox2 T("Unknown Operation State"), T("Current state: {0}", , , , this.State)
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : Stage
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Stage the current operation for a timer restart, or to run a user hook.
'---------------------------------------------------------------------------------------
'
Public Function Stage()
    this.State = eosStaged
    UpdateRegistry
    RestoreUserErrorTrapping
    Me.Pulse
End Function


'---------------------------------------------------------------------------------------
' Procedure : Restore
' Author    : Adam Waller
' Date      : 5/10/2025
' Purpose   : Restore the staged process
'---------------------------------------------------------------------------------------
'
Public Function Restore()
    SetErrorTrapping
    RestoreFromRegistry
    this.State = eosRunning
    UpdateRegistry
    Me.Pulse
End Function


'---------------------------------------------------------------------------------------
' Procedure : Finish
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Complete the current operation, release all objects.
'---------------------------------------------------------------------------------------
'
Public Function Finish(Optional intResult As eOperationResult)
    Dim blnWasRunning As Boolean
    If this.State <> eosReady Then
        If Me.DecisionBlocked Then
            ' An unresolved prompt is not a successful or merely cancelled run.
            Me.Result = eorDecisionRequired
        ElseIf intResult > eorUnknown Then
            Me.Result = intResult
        End If
        blnWasRunning = (this.State = eosRunning)
        this.State = eosReady
        this.FinishTime = Now
        UpdateRegistry

        ' Post MCP completion callback BEFORE releasing objects
        ' (ReleaseObjects destroys the MCP singleton)
        If MCP.IsActive Then
            ' Build extra data with log path
            Dim dExtra As Dictionary
            Set dExtra = New Dictionary
            dExtra.Add "log_path", Log.LogFilePath
            If Scope.DecisionCount > 0 Then dExtra.Add "decisions", Scope.Decisions

            Select Case Me.Result
                Case eorSuccess
                    MCP.PostCallback "complete", 100, 100, "Operation completed successfully", dExtra
                Case eorFailed
                    MCP.PostCallback "error", -1, -1, "Operation failed", dExtra
                Case eorCanceled
                    MCP.PostCallback "cancelled", -1, -1, "Operation was cancelled", dExtra
                Case eorTimeout
                    MCP.PostCallback "error", -1, -1, "Operation timed out", dExtra
                Case eorDecisionRequired
                    dExtra.Add "decision_required", True
                    dExtra.Add "error_pattern", "decision_required"
                    MCP.PostCallback "error", -1, -1, "Decision required", dExtra
                Case Else
                    ' Unknown or unset result - assume success since we got here
                    MCP.PostCallback "complete", 100, 100, "Operation completed", dExtra
            End Select
        End If

        ' Restore interaction state after the callback has read the journal.
        Set m_LastDecisions = Scope.Decisions
        Scope.ReleaseOperationScopes
        UpdateRegistry
        modObjects.ReleaseObjects
        ' Only restore error trapping if we were in a running state.
        If blnWasRunning Then RestoreUserErrorTrapping
    End If
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
    ' Check for timeout
    If this.State <> eosReady Then
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, Me.Heartbeat) < Now Then
            ' Timeout exceeded from last heartbeat
            this.State = eosReady
            Me.Result = eorTimeout
            Scope.ReleaseOperationScopes
            RestoreUserErrorTrapping
        End If
    End If
    Status = this.State
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceName
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Return the display name for the operation source. Returns empty string
'           : for UI operations (the default) since we don't need to log that.
'---------------------------------------------------------------------------------------
'
Public Property Get SourceName() As String
    Select Case Me.Source
        Case eosMCPTool:     SourceName = "MCP Tool"
        Case eosExternalAPI: SourceName = "External API"
        ' eosUserInterface: return empty string (default case, no output needed)
    End Select
End Property


'---------------------------------------------------------------------------------------
' Procedure : UpdateRegistry
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Updates values in registry to reflect current class values
'---------------------------------------------------------------------------------------
'
Private Sub UpdateRegistry()
    ' Make sure we still have a valid cache
    If this.AppHwnd <> 0 Then
        CheckSetting "Heartbeat", Me.Heartbeat
        CheckSetting "Result", Me.Result
        CheckSetting "StartTime", this.StartTime
        CheckSetting "FinishTime", this.FinishTime
        CheckSetting "State", this.State
        CheckSetting "AppHwnd", this.AppHwnd
        CheckSetting "ErrorTrapping", this.ErrorTrapping
        CheckSetting "InteractionMode", Scope.InteractionMode
        CheckSetting "DecisionPolicy", Scope.DecisionPolicy
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckSetting
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Update the setting if it has changed. (Faster to read than write)
'---------------------------------------------------------------------------------------
'
Private Sub CheckSetting(strKey As String, varValue)
    If GetSetting(PROJECT_NAME, "State", strKey) <> CStr(varValue) Then
        SaveSetting PROJECT_NAME, "State", strKey, CStr(varValue)
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearRegistryCache
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Clear the entire key for the state.
'---------------------------------------------------------------------------------------
'
Private Sub ClearRegistryCache()
    DeleteSetting PROJECT_NAME, "State"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreFromRegistry
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Restore the operation state from the registry cache. Returns true if
'           : successful.
'---------------------------------------------------------------------------------------
'
Private Function RestoreFromRegistry() As Boolean
    ' Only restore to the same Access instance
    If ReadReg("AppHwnd") = CStr(hWndAccessApp) Then
        ' Check for timeout
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, CDate(Nz2(ReadReg("Heartbeat"), 0))) > Now Then
            ' Load results
            Me.Result = ReadReg("Result")
            this.StartTime = ReadReg("StartTime")
            this.FinishTime = ReadReg("FinishTime")
            this.State = ReadReg("State")
            this.AppHwnd = ReadReg("AppHwnd")
            this.ErrorTrapping = CInt(ReadReg("ErrorTrapping"))
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

    If this.State = eosReady Then Exit Sub
    intMode = Val(ReadReg("InteractionMode"))

    If intMode = eimNonInteractive Then
        If Not Scope.InNonInteractiveScope Then
            Scope.OpenScope eimNonInteractive, Val(ReadReg("DecisionPolicy")), True
        End If
    Else
        Scope.InteractionMode = intMode
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReadReg
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Read a setting from the registry (wrapper for readability)
'---------------------------------------------------------------------------------------
'
Private Function ReadReg(strKey As String) As String
    ReadReg = GetSetting(PROJECT_NAME, "State", strKey)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasFormOpen
' Author    : Adam Waller
' Date      : 3/29/2024
' Purpose   : Make sure no other forms are currently loaded before initiating a new
'           : operation. (Force linear usage of one operation at a time.)
'           : Also warns user that they must close the form before they can proceed.
'---------------------------------------------------------------------------------------
'
Private Function HasFormOpen(Optional blnWarnUser As Boolean = True) As Boolean

    Dim objForm As AccessObject
    Dim frm As Access.Form
    Dim strCaption As String

    ' Cache collection of form names used by this add-in
    Static dFormNames As Dictionary

    ' Build cached list of form names
    If dFormNames Is Nothing Then
        Set dFormNames = New Dictionary
        For Each objForm In CodeProject.AllForms
            dFormNames.Add objForm.Name, vbNullString
        Next objForm
    End If

    ' Check for open forms (from current database, add-in, or library database)
    For Each frm In Application.Forms
        If dFormNames.Exists(frm.Name) Then
            ' Found a matching form name.
            If frm.CurrentView = acCurViewDesign Then
                ' Design view is okay, since the add-in won't have a form in design
                ' view when running as an add-in. (No conflict)
            Else
                ' A form with the same name. This could be a problem.
                If blnWarnUser Then
                    If frm.Caption = vbNullString Then
                        strCaption = frm.Name
                    Else
                        ' Users may recognize caption better than name
                        strCaption = frm.Caption & " (" & frm.Name & ")"
                    End If
                    MsgBox2 "Add-in Form Already Open", _
                        "Please close '" & strCaption & "' before running this action.", , vbInformation
                    HasFormOpen = True
                    Exit Function
                End If
            End If
        End If
    Next

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetErrorTrapping
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : We need this to "Break in Class Modules" for some functionality in this
'           : add-in. (Such as debugging VBA code)
'---------------------------------------------------------------------------------------
'
Private Sub SetErrorTrapping()
    this.ErrorTrapping = Application.GetOption("Error Trapping")
    If this.ErrorTrapping <> 1 Then Application.SetOption "Error Trapping", 1
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreUserErrorTrapping
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Restore the user's preference for error trapping.
'---------------------------------------------------------------------------------------
'
Private Sub RestoreUserErrorTrapping()
    If this.ErrorTrapping <> Application.GetOption("Error Trapping") Then
        Application.SetOption "Error Trapping", this.ErrorTrapping
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Scope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : The interaction scope object this operation delegates to.
'---------------------------------------------------------------------------------------
'
Private Property Get Scope() As clsInteractionScope
    If m_Scope Is Nothing Then Set m_Scope = New clsInteractionScope
    Set Scope = m_Scope
End Property


'---------------------------------------------------------------------------------------
' Procedure : InteractionMode
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Controls UI prompts during the operation. eimNormal (0) shows MsgBox2
'           : dialogs; eimSilent (1) routes them to the log and returns the supplied
'           : default; eimNonInteractive (2) answers them from DecisionPolicy.
'           : Outside a noninteractive scope the mode is sticky, like Source; callers
'           : that need to restore the prior value cache and restore it themselves.
'           : Inside a noninteractive scope, a request for a weaker mode is ignored.
'---------------------------------------------------------------------------------------
'
Public Property Get InteractionMode() As eInteractionMode
    InteractionMode = Scope.InteractionMode
End Property

Public Property Let InteractionMode(ByVal intMode As eInteractionMode)
    Scope.InteractionMode = intMode
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
    ByVal strMessage As String, ByRef strResolution As String) As VbMsgBoxResult
    ResolvePrompt = Scope.ResolvePrompt(intButtons, strTitle, strMessage, strResolution)
    If Scope.Blocked Then Me.Result = eorDecisionRequired
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveConflictAction
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Resolution for one conflict under the current policy.
'---------------------------------------------------------------------------------------
'
Public Function ResolveConflictAction(ByVal intRequested As eResolveConflict) As eResolveConflict
    ResolveConflictAction = Scope.ResolveConflictAction(intRequested)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordDecision
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Append one prompt outcome. decision_required blocks the operation
'           : and prevents Finish from reporting success.
'---------------------------------------------------------------------------------------
'
Public Sub RecordDecision(ByVal strKind As String, ByVal strTitle As String, _
    ByVal strMessage As String, ByVal strResolution As String, _
    Optional ByVal strDetail As String)

    Scope.RecordDecision strKind, strTitle, strMessage, strResolution, strDetail
    If Scope.Blocked Then Me.Result = eorDecisionRequired

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DecisionsJson
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : JSON array of prompts recorded in the current scope.
'---------------------------------------------------------------------------------------
'
Public Function DecisionsJson() As String
    DecisionsJson = Scope.DecisionsJson()
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
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Load initial settings for class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    ' Attempt to restore from registry
    If Not RestoreFromRegistry Then
        ' Set initial values
        this.AppHwnd = Application.hWndAccessApp
    End If
End Sub

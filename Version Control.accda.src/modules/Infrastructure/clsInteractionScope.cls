VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInteractionScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsInteractionScope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Owns the interaction mode, decision policy, blocked flag, decision journal
'           : and error level for an operation. Scopes are opened and closed by owner
'           : token, so a close can only ever undo the scope that was opened for it.
'           : Closing a token restores the state saved when it opened (including the
'           : error level) and discards any scopes opened after it. Closing a token
'           : that is not open does nothing.
'           :
'           : A scope is either operation-owned (released by ReleaseOperationScopes
'           : when the operation finishes) or caller-owned (closed only by its token).
'           : While any open scope is noninteractive, requests for a weaker mode
'           : (silent or normal) leave the effective mode noninteractive.
'           :
'           : This class does not touch the live operation, registry or singletons,
'           : so tests can create their own instance.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")



' State saved when a scope opens, restored when it closes.
Private Type udtFrame
    Token As Long
    OperationOwned As Boolean
    AppliedMode As eInteractionMode
    SavedMode As eInteractionMode
    SavedPolicy As eDecisionPolicy
    SavedBlocked As Boolean
    SavedErrorLevel As eErrorLevel
    SavedJournal As Collection
End Type

Private Type udtThis
    Mode As eInteractionMode
    Policy As eDecisionPolicy
    Blocked As Boolean
    ErrorLevel As eErrorLevel
    Journal As Collection
    Frames() As udtFrame
    Depth As Long
    LastToken As Long
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : OpenScope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Save the current state and apply a scoped mode and policy. Returns the
'           : owner token to pass to CloseScope. The journal and blocked flag start
'           : empty for the new scope. Inside a noninteractive scope, a request for a
'           : weaker mode keeps the enclosing mode and policy.
'---------------------------------------------------------------------------------------
'
Public Function OpenScope(ByVal intMode As eInteractionMode, _
    ByVal intPolicy As eDecisionPolicy, _
    Optional ByVal blnOperationOwned As Boolean = False) As Long

    Dim blnWeaken As Boolean

    blnWeaken = (intMode <> eimNonInteractive) And InNonInteractiveScope

    this.Depth = this.Depth + 1
    If this.Depth = 1 Then
        ReDim this.Frames(1 To 1)
    Else
        ReDim Preserve this.Frames(1 To this.Depth)
    End If
    this.LastToken = this.LastToken + 1

    With this.Frames(this.Depth)
        .Token = this.LastToken
        .OperationOwned = blnOperationOwned
        .SavedMode = this.Mode
        .SavedPolicy = this.Policy
        .SavedBlocked = this.Blocked
        .SavedErrorLevel = this.ErrorLevel
        Set .SavedJournal = this.Journal
    End With

    If Not blnWeaken Then
        this.Mode = intMode
        this.Policy = intPolicy
    End If
    this.Blocked = False
    Set this.Journal = New Collection
    this.Frames(this.Depth).AppliedMode = this.Mode

    OpenScope = this.LastToken

End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseScope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Restore the state saved when this token's scope opened, and discard any
'           : scopes opened after it. An unknown, already closed, or zero token does
'           : nothing.
'---------------------------------------------------------------------------------------
'
Public Sub CloseScope(ByVal lngToken As Long)

    Dim lngIdx As Long

    lngIdx = FindFrame(lngToken)
    If lngIdx = 0 Then Exit Sub
    RestoreFrame lngIdx
    TrimFrames lngIdx - 1

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReleaseOperationScopes
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Close every operation-owned scope, leaving caller-owned scopes open.
'           : A caller-owned scope opened inside an operation-owned one keeps its
'           : own state and inherits what the released scope had saved.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseOperationScopes()

    Dim lngIdx As Long
    Dim lngPos As Long

    For lngIdx = this.Depth To 1 Step -1
        If this.Frames(lngIdx).OperationOwned Then
            If lngIdx = this.Depth Then
                RestoreFrame lngIdx
            Else
                ' The scope above takes over the state saved by the one being removed.
                With this.Frames(lngIdx + 1)
                    .SavedMode = this.Frames(lngIdx).SavedMode
                    .SavedPolicy = this.Frames(lngIdx).SavedPolicy
                    .SavedBlocked = this.Frames(lngIdx).SavedBlocked
                    .SavedErrorLevel = this.Frames(lngIdx).SavedErrorLevel
                    Set .SavedJournal = this.Frames(lngIdx).SavedJournal
                End With
                For lngPos = lngIdx To this.Depth - 1
                    this.Frames(lngPos) = this.Frames(lngPos + 1)
                Next lngPos
            End If
            TrimFrames this.Depth - 1
        End If
    Next lngIdx

End Sub


'---------------------------------------------------------------------------------------
' Procedure : InteractionMode
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Effective mode. A request for silent or normal inside a noninteractive
'           : scope is ignored; it can never weaken the scope.
'---------------------------------------------------------------------------------------
'
Public Property Get InteractionMode() As eInteractionMode
    InteractionMode = this.Mode
End Property

Public Property Let InteractionMode(ByVal intMode As eInteractionMode)
    If intMode <> eimNonInteractive And InNonInteractiveScope Then Exit Property
    this.Mode = intMode
End Property


'---------------------------------------------------------------------------------------
' Procedure : DecisionPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Policy for the current scope. Not changeable from inside a
'           : noninteractive scope; open a scope to apply a different policy.
'---------------------------------------------------------------------------------------
'
Public Property Get DecisionPolicy() As eDecisionPolicy
    DecisionPolicy = this.Policy
End Property

Public Property Let DecisionPolicy(ByVal intPolicy As eDecisionPolicy)
    If InNonInteractiveScope Then Exit Property
    this.Policy = intPolicy
End Property


'---------------------------------------------------------------------------------------
' Procedure : Blocked
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True when a required prompt in this scope was not covered by the policy.
'---------------------------------------------------------------------------------------
'
Public Property Get Blocked() As Boolean
    Blocked = this.Blocked
End Property

Public Property Let Blocked(ByVal blnBlocked As Boolean)
    this.Blocked = blnBlocked
End Property


'---------------------------------------------------------------------------------------
' Procedure : ErrorLevel
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Operation error level. Saved when a scope opens and restored when it
'           : closes, so a blocked prompt in a nested scope does not raise the
'           : outer operation's level.
'---------------------------------------------------------------------------------------
'
Public Property Get ErrorLevel() As eErrorLevel
    ErrorLevel = this.ErrorLevel
End Property

Public Property Let ErrorLevel(ByVal intLevel As eErrorLevel)
    this.ErrorLevel = intLevel
End Property


'---------------------------------------------------------------------------------------
' Procedure : InNonInteractiveScope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True while any open scope is noninteractive.
'---------------------------------------------------------------------------------------
'
Public Property Get InNonInteractiveScope() As Boolean

    Dim lngIdx As Long

    For lngIdx = 1 To this.Depth
        If this.Frames(lngIdx).AppliedMode = eimNonInteractive Then
            InNonInteractiveScope = True
            Exit Property
        End If
    Next lngIdx

End Property


'---------------------------------------------------------------------------------------
' Procedure : IsOpen
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : True when this token's scope is still open.
'---------------------------------------------------------------------------------------
'
Public Function IsOpen(ByVal lngToken As Long) As Boolean
    IsOpen = (FindFrame(lngToken) > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordDecision
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Append one prompt outcome to the journal. A decision_required outcome
'           : blocks the scope and raises the error level to critical.
'---------------------------------------------------------------------------------------
'
Public Sub RecordDecision(ByVal cRecord As clsDecisionRecord)

    Dim dItem As Dictionary

    Set dItem = New Dictionary
    dItem.Add "kind", cRecord.Kind
    dItem.Add "title", cRecord.Title
    dItem.Add "message", cRecord.Message
    dItem.Add "resolution", cRecord.Resolution
    dItem.Add "detail", cRecord.Detail
    this.Journal.Add dItem

    If StrComp(cRecord.Kind, DECISION_REQUIRED, vbTextCompare) = 0 Then
        this.Blocked = True
        If this.ErrorLevel < eelCritical Then this.ErrorLevel = eelCritical
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolvePrompt
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Answer a message box for the current policy without showing it, and
'           : record the outcome. The outcome holds the button to report to the
'           : caller and the resolution: acknowledged, declined or decision_required.
'---------------------------------------------------------------------------------------
'
Public Function ResolvePrompt(ByVal intButtons As Long, _
    Optional ByVal strTitle As String, _
    Optional ByVal strMessage As String) As clsPromptOutcome

    Dim cOutcome As clsPromptOutcome

    Set cOutcome = ResolveNonInteractivePrompt(intButtons, this.Policy)
    If cOutcome.Blocked Then
        RecordDecision NewDecision(DECISION_REQUIRED, strTitle, strMessage, cOutcome.Resolution, _
            "buttons=" & CStr(intButtons And MB_BUTTON_STYLE_MASK))
    ElseIf cOutcome.Resolution = DECISION_ACKNOWLEDGED Then
        RecordDecision NewDecision(DECISION_ACKNOWLEDGED, strTitle, strMessage, cOutcome.Resolution)
    Else
        RecordDecision NewDecision(DECISION_APPLIED, strTitle, strMessage, cOutcome.Resolution, _
            "result=" & CStr(cOutcome.Result))
    End If
    Set ResolvePrompt = cOutcome

End Function


'---------------------------------------------------------------------------------------
' Procedure : DecisionCount
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Number of decisions recorded in the current scope.
'---------------------------------------------------------------------------------------
'
Public Property Get DecisionCount() As Long
    DecisionCount = this.Journal.Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : Decisions
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Decisions recorded in the current scope, as a collection of dictionaries
'           : (kind, title, message, resolution, detail).
'---------------------------------------------------------------------------------------
'
Public Function Decisions() As Collection
    Set Decisions = this.Journal
End Function


'---------------------------------------------------------------------------------------
' Procedure : FindFrame
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Index of the open scope for this token, or 0 when it is not open.
'---------------------------------------------------------------------------------------
'
Private Function FindFrame(ByVal lngToken As Long) As Long

    Dim lngIdx As Long

    If lngToken = 0 Then Exit Function
    For lngIdx = this.Depth To 1 Step -1
        If this.Frames(lngIdx).Token = lngToken Then
            FindFrame = lngIdx
            Exit Function
        End If
    Next lngIdx

End Function


'---------------------------------------------------------------------------------------
' Procedure : RestoreFrame
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Put back the state saved by this scope.
'---------------------------------------------------------------------------------------
'
Private Sub RestoreFrame(ByVal lngIdx As Long)

    With this.Frames(lngIdx)
        this.Mode = .SavedMode
        this.Policy = .SavedPolicy
        this.Blocked = .SavedBlocked
        this.ErrorLevel = .SavedErrorLevel
        Set this.Journal = .SavedJournal
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TrimFrames
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Keep only the first lngDepth scopes.
'---------------------------------------------------------------------------------------
'
Private Sub TrimFrames(ByVal lngDepth As Long)

    this.Depth = lngDepth
    If lngDepth = 0 Then
        Erase this.Frames
    Else
        ReDim Preserve this.Frames(1 To lngDepth)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Start in normal mode with an empty journal.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set this.Journal = New Collection
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestPromptSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestPromptSites
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Wiring tests for the prompt sites. They call the real MsgBox2 and the real
'           : conflict-action mapping under an interaction scope on the live operation,
'           : and check that no dialog blocks the run, the return value, the recorded
'           : decision, and that the live state is the same after the scope closes.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")


' Live state captured before each test and checked again after it.
Private m_Mode As eInteractionMode
Private m_Policy As eDecisionPolicy
Private m_Blocked As Boolean
Private m_ErrorLevel As eErrorLevel


Private Sub Class_Initialize()
    m_Mode = Operation.InteractionMode
    m_Policy = Operation.DecisionPolicy
    m_Blocked = Operation.DecisionBlocked
    m_ErrorLevel = Operation.ErrorLevel
End Sub


Private Function LiveStateUnchanged() As Boolean
    LiveStateUnchanged = (Operation.InteractionMode = m_Mode) _
        And (Operation.DecisionPolicy = m_Policy) _
        And (Operation.DecisionBlocked = m_Blocked) _
        And (Operation.ErrorLevel = m_ErrorLevel)
End Function


Private Function LastDecisionKind() As String
    With Operation.Decisions
        If .Count > 0 Then LastDecisionKind = .Item(.Count)("kind")
    End With
End Function


Public Sub TestMsgBox2NoDialogUnderEachPolicy()

    Dim varPolicy As Variant
    Dim lngToken As Long
    Dim intResult As VbMsgBoxResult

    For Each varPolicy In Array(edpBlock, edpPreferSource, edpPreferDatabase, edpSkip, edpDecline)
        lngToken = Operation.PushInteractionScope(eimNonInteractive, varPolicy, False)

        intResult = MsgBox2("Information", , , vbOKOnly)
        TestAssert intResult = vbOK, "an OK-only prompt is acknowledged (policy " & varPolicy & ")"
        TestAssert Not Operation.DecisionBlocked, "an OK-only prompt does not block (policy " & varPolicy & ")"

        intResult = MsgBox2("Overwrite?", , , vbYesNo)
        If varPolicy = edpDecline Then
            TestAssert intResult = vbNo, "decline answers No"
            TestAssert Not Operation.DecisionBlocked, "decline does not block the operation"
        Else
            TestAssert intResult = vbNo, "a conflict policy never approves a generic confirmation (policy " & varPolicy & ")"
            TestAssert Operation.DecisionBlocked, "the unanswered confirmation blocks (policy " & varPolicy & ")"
            TestAssert LastDecisionKind = "decision_required", "the confirmation is recorded as decision_required (policy " & varPolicy & ")"
        End If

        Operation.CloseInteractionScope lngToken
        TestAssert LiveStateUnchanged, "closing the scope restores the live state (policy " & varPolicy & ")"
    Next varPolicy

End Sub


Public Sub TestMsgBox2CancelStyleReturnsNonDestructiveButton()

    Dim lngToken As Long

    lngToken = Operation.PushInteractionScope(eimNonInteractive, edpDecline, False)
    TestAssert MsgBox2("Continue?", , , vbOKCancel) = vbCancel, "OK/Cancel is declined with Cancel"
    TestAssert MsgBox2("Retry?", , , vbAbortRetryIgnore) = vbAbort, "Abort/Retry/Ignore is declined with Abort"
    TestAssert MsgBox2("Save?", , , vbYesNoCancel) = vbNo, "Yes/No/Cancel is declined with No"
    Operation.CloseInteractionScope lngToken
    TestAssert LiveStateUnchanged, "closing the scope restores the live state"

End Sub


Public Sub TestConflictActionAppliesRequestedActionOnlyForPreferSource()

    Dim lngToken As Long

    lngToken = Operation.PushInteractionScope(eimNonInteractive, edpPreferSource, False)
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercSkip) = ercSkip, "prefer_source applies the conflict's requested action"
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercNone) = ercOverwrite, "prefer_source falls back to overwrite"
    Operation.CloseInteractionScope lngToken

    lngToken = Operation.PushInteractionScope(eimNonInteractive, edpPreferDatabase, False)
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercOverwrite) = ercSkip, "prefer_database keeps the database object"
    Operation.CloseInteractionScope lngToken

    lngToken = Operation.PushInteractionScope(eimNonInteractive, edpSkip, False)
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercOverwrite) = ercSkip, "skip is equivalent to prefer_database"
    Operation.CloseInteractionScope lngToken

    lngToken = Operation.PushInteractionScope(eimNonInteractive, edpBlock, False)
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercOverwrite) = ercNone, "block leaves the conflict unresolved"
    Operation.CloseInteractionScope lngToken

    TestAssert LiveStateUnchanged, "closing the scopes restores the live state"

End Sub


Public Sub TestPredicatesFollowTheScope()

    Dim lngToken As Long

    lngToken = Operation.PushInteractionScope(eimNonInteractive, edpBlock, False)
    TestAssert InteractionIsNonInteractive, "the predicate is true inside a noninteractive scope"
    TestAssert Not InteractionIsNormal, "the normal predicate is false inside a noninteractive scope"
    Operation.CloseInteractionScope lngToken

    TestAssert InteractionIsNonInteractive = (m_Mode = eimNonInteractive), "the predicate follows the restored mode"
    TestAssert LiveStateUnchanged, "closing the scope restores the live state"

End Sub

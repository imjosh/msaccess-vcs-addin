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
Private m_Journal As Collection
Private m_JournalJson As String
Private m_Token As Long
Private m_SetupError As String
Private m_FaultPoint As String
Private m_CleanupFault As Boolean
Private m_CaptureFailure As Boolean
Private m_Report As Dictionary


Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    m_Mode = Operation.InteractionMode
    m_Policy = Operation.DecisionPolicy
    m_Blocked = Operation.DecisionBlocked
    m_ErrorLevel = Operation.ErrorLevel
    Set m_Journal = Operation.Decisions
    m_JournalJson = ConvertToJson(m_Journal)
    Exit Sub
ErrHandler:
    m_SetupError = Err.Number & ": " & Err.Description
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    ' Fallback for an abandoned fixture; explicit teardown reports before destruction.
    Dim strCleanup As String
    strCleanup = ReleaseOwnedScope()
    If Len(strCleanup) > 0 Then TestAssert False, "prompt fixture teardown: " & strCleanup
End Sub


Private Sub OpenOwnedScope(ByVal intPolicy As eDecisionPolicy)
    If Len(m_SetupError) > 0 Then Error 5
    m_Token = Operation.PushInteractionScope(eimNonInteractive, intPolicy, False)
    InjectFault "acquired"
End Sub


Private Sub InjectFault(ByVal strPoint As String)
    Dim lngZero As Long
    Dim dblValue As Double
    If m_FaultPoint = strPoint Then
        Operation.ErrorLevel = eelCritical
        dblValue = 1 / lngZero
    End If
End Sub


Private Function ReleaseOwnedScope() As String
    On Error Resume Next
    Dim intOverflow As Integer
    ' A fallible cleanup step must not prevent the actual owned-token release.
    Err.Clear
    If m_CleanupFault Then intOverflow = 40000
    m_CleanupFault = False
    If Err.Number <> 0 Then ReleaseOwnedScope = "cleanup step " & Err.Number & ": " & Err.Description
    Err.Clear
    If m_Token <> 0 Then Operation.CloseInteractionScope m_Token
    If Err.Number <> 0 Then
        ReleaseOwnedScope = ReleaseOwnedScope & " scope close " & Err.Number & ": " & Err.Description
    Else
        m_Token = 0
    End If
    Err.Clear
End Function


Private Sub FinishTest(ByVal lngErr As Long, ByVal strErr As String, ByVal strSource As String)
    On Error Resume Next
    Dim strCleanup As String
    strCleanup = ReleaseOwnedScope()
    Set m_Report = New Dictionary
    m_Report.Add "error_number", lngErr
    m_Report.Add "error_description", strErr
    m_Report.Add "error_source", strSource
    m_Report.Add "cleanup_error", strCleanup
    If Not m_CaptureFailure Then
        If lngErr <> 0 Then TestAssert False, "prompt test failure " & lngErr & ": " & strErr & " (" & strSource & ")"
        If Len(strCleanup) > 0 Then TestAssert False, "prompt test cleanup: " & strCleanup
    End If
End Sub


Private Function LiveStateUnchanged() As Boolean
    LiveStateUnchanged = (Operation.InteractionMode = m_Mode) _
        And (Operation.DecisionPolicy = m_Policy) _
        And (Operation.DecisionBlocked = m_Blocked) _
        And (Operation.ErrorLevel = m_ErrorLevel) _
        And (Operation.Decisions Is m_Journal) _
        And (ConvertToJson(Operation.Decisions) = m_JournalJson)
End Function


Private Function LastDecisionKind() As String
    InjectFault "journal"
    With Operation.Decisions
        If .Count > 0 Then LastDecisionKind = .Item(.Count)("kind")
    End With
End Function


Public Sub TestMsgBox2NoDialogUnderEachPolicy()
    On Error GoTo ErrHandler

    Dim varPolicy As Variant
    Dim intResult As VbMsgBoxResult

    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String

    For Each varPolicy In Array(edpBlock, edpPreferSource, edpPreferDatabase, edpSkip, edpDecline)
        OpenOwnedScope varPolicy

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

        InjectFault "work"
        CloseOwnedScope
        TestAssert LiveStateUnchanged, "closing the scope restores the live state (policy " & varPolicy & ")"
    Next varPolicy

CleanUp:
    On Error Resume Next
    FinishTest lngErr, strErr, strSource
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    Resume CleanUp
End Sub


Public Sub TestMsgBox2CancelStyleReturnsNonDestructiveButton()
    On Error GoTo ErrHandler


    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String

    OpenOwnedScope edpDecline
    TestAssert MsgBox2("Continue?", , , vbOKCancel) = vbCancel, "OK/Cancel is declined with Cancel"
    TestAssert MsgBox2("Retry?", , , vbAbortRetryIgnore) = vbAbort, "Abort/Retry/Ignore is declined with Abort"
    TestAssert MsgBox2("Save?", , , vbYesNoCancel) = vbNo, "Yes/No/Cancel is declined with No"
    InjectFault "work"
    CloseOwnedScope
    TestAssert LiveStateUnchanged, "closing the scope restores the live state"

CleanUp:
    On Error Resume Next
    FinishTest lngErr, strErr, strSource
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    Resume CleanUp
End Sub


Public Sub TestConflictActionAppliesRequestedActionOnlyForPreferSource()
    On Error GoTo ErrHandler


    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String

    OpenOwnedScope edpPreferSource
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercSkip) = ercSkip, "prefer_source applies the conflict's requested action"
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercNone) = ercOverwrite, "prefer_source falls back to overwrite"
    InjectFault "work"
    CloseOwnedScope

    OpenOwnedScope edpPreferDatabase
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercOverwrite) = ercSkip, "prefer_database keeps the database object"
    InjectFault "work"
    CloseOwnedScope

    OpenOwnedScope edpSkip
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercOverwrite) = ercSkip, "skip is equivalent to prefer_database"
    InjectFault "work"
    CloseOwnedScope

    OpenOwnedScope edpBlock
    TestAssert ConflictActionForPolicy(Operation.DecisionPolicy, ercOverwrite) = ercNone, "block leaves the conflict unresolved"
    InjectFault "work"
    CloseOwnedScope

    TestAssert LiveStateUnchanged, "closing the scopes restores the live state"

CleanUp:
    On Error Resume Next
    FinishTest lngErr, strErr, strSource
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    Resume CleanUp
End Sub


Public Sub TestPredicatesFollowTheScope()
    On Error GoTo ErrHandler


    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String

    OpenOwnedScope edpBlock
    TestAssert InteractionIsNonInteractive, "the predicate is true inside a noninteractive scope"
    TestAssert Not InteractionIsNormal, "the normal predicate is false inside a noninteractive scope"
    InjectFault "work"
    CloseOwnedScope

    TestAssert InteractionIsNonInteractive = (m_Mode = eimNonInteractive), "the predicate follows the restored mode"
    TestAssert LiveStateUnchanged, "closing the scope restores the live state"

CleanUp:
    On Error Resume Next
    FinishTest lngErr, strErr, strSource
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    Resume CleanUp
End Sub


Private Sub CloseOwnedScope()
    Operation.CloseInteractionScope m_Token
    m_Token = 0
End Sub


' Parameterized public boundary drives the actual tests, not a duplicate body.
' Expected faults are returned to the outer regression rather than logged as failures.
Public Function ProbeExit(ByVal strTest As String, ByVal strFault As String, _
    ByVal blnCleanupFault As Boolean) As Dictionary
    On Error GoTo ErrHandler
    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String
    m_CaptureFailure = True
    m_FaultPoint = strFault
    m_CleanupFault = blnCleanupFault
    Select Case strTest
        Case "policies": TestMsgBox2NoDialogUnderEachPolicy
        Case "buttons": TestMsgBox2CancelStyleReturnsNonDestructiveButton
        Case "conflicts": TestConflictActionAppliesRequestedActionOnlyForPreferSource
        Case "predicates": TestPredicatesFollowTheScope
    End Select
    Set ProbeExit = m_Report
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    Resume CleanUp
CleanUp:
    On Error Resume Next
    FinishTest lngErr, strErr, strSource
    Set ProbeExit = m_Report
End Function


Public Sub HoldForTermination(ByVal intPolicy As eDecisionPolicy)
    On Error GoTo ErrHandler
    OpenOwnedScope intPolicy
    MsgBox2 "Abandoned fixture?", , , vbYesNo
    Operation.ErrorLevel = eelCritical
    Exit Sub
ErrHandler:
    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    Resume CleanUp
CleanUp:
    On Error Resume Next
    FinishTest lngErr, strErr, strSource
End Sub


Public Sub TestExceptionalAndNormalExitsRestoreCallerState()
    On Error GoTo ErrHandler
    Dim cFixture As clsTestPromptSites
    Dim dFailure As Dictionary
    Dim varTest As Variant
    Dim varFault As Variant
    Dim strCallerJournal As String
    Dim colCallerJournal As Collection
    Dim lngErr As Long
    Dim strErr As String
    Dim strSource As String

    ' Non-default caller-owned baseline catches accidental clearing as well as leaks.
    OpenOwnedScope edpBlock
    MsgBox2 "Caller decision?", , , vbYesNo
    Operation.ErrorLevel = eelWarning
    Set colCallerJournal = Operation.Decisions
    strCallerJournal = ConvertToJson(colCallerJournal)
    For Each varTest In Array("policies", "buttons", "conflicts", "predicates")
        For Each varFault In Array("acquired", "work", "")
            Set cFixture = New clsTestPromptSites
            Set dFailure = cFixture.ProbeExit(CStr(varTest), CStr(varFault), Len(varFault) > 0)
            TestAssert dFailure("error_number") = IIf(Len(varFault) > 0, 11, 0), "original runtime failure: " & varTest & "/" & varFault
            If Len(varFault) > 0 Then
                TestAssert Len(dFailure("error_description")) > 0, "original description retained"
                TestAssert Len(dFailure("error_source")) > 0, "original source retained"
                TestAssert InStr(dFailure("cleanup_error"), "cleanup step 6:") > 0, "cleanup overflow recorded separately"
            Else
                TestAssert dFailure("cleanup_error") = "", "normal cleanup succeeds"
            End If
            TestAssert Operation.InteractionMode = eimNonInteractive, "caller mode restored"
            TestAssert Operation.DecisionPolicy = edpBlock, "caller policy restored"
            TestAssert Operation.DecisionBlocked, "caller blocked state restored"
            TestAssert Operation.ErrorLevel = eelWarning, "caller error level restored"
            TestAssert Operation.Decisions Is colCallerJournal, "caller journal identity restored"
            TestAssert ConvertToJson(Operation.Decisions) = strCallerJournal, "caller journal contents restored"

            ' The same fixture's next real prompt test must be clean after failure.
            Set dFailure = cFixture.ProbeExit(CStr(varTest), "", False)
            TestAssert dFailure("error_number") = 0, "subsequent test completes cleanly"
            TestAssert dFailure("cleanup_error") = "", "subsequent teardown completes cleanly"
            Set cFixture = Nothing
        Next varFault
    Next varTest
    Set cFixture = New clsTestPromptSites
    Set dFailure = cFixture.ProbeExit("policies", "journal", False)
    TestAssert dFailure("error_number") = 11, "fallible journal assertion support also tears down"
    Set cFixture = Nothing
    TestAssert ConvertToJson(Operation.Decisions) = strCallerJournal, "journal-support failure preserves caller journal"
    Set cFixture = New clsTestPromptSites
    Set dFailure = cFixture.ProbeExit("buttons", "", True)
    TestAssert dFailure("error_number") = 0, "cleanup-only failure never invents a primary failure"
    TestAssert InStr(dFailure("cleanup_error"), "cleanup step 6:") > 0, "cleanup-only failure recorded separately"
    Set cFixture = Nothing
    Set cFixture = New clsTestPromptSites
    cFixture.HoldForTermination edpBlock
    Set cFixture = Nothing
    TestAssert Operation.ErrorLevel = eelWarning, "termination fallback restores error level"
    TestAssert Operation.DecisionBlocked And Operation.DecisionPolicy = edpBlock, "termination fallback restores caller decision"
    TestAssert Operation.Decisions Is colCallerJournal, "termination fallback restores caller journal"
    CloseOwnedScope
    TestAssert LiveStateUnchanged, "outer fixture restores all pre-test state"
CleanUp:
    On Error Resume Next
    Set cFixture = Nothing
    FinishTest lngErr, strErr, strSource
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    strSource = Err.Source
    Resume CleanUp
End Sub

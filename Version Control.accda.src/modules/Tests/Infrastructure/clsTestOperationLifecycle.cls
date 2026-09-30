VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestOperationLifecycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestOperationLifecycle
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Wiring tests for policy requests against a running operation: a policy request
'           : made while an operation runs is refused without changing that operation's
'           : interaction state. Every test that needs a running operation uses a private
'           : instance, so no result depends on where the host run started. The session
'           : operation is only read, by the clearing test. Start, restore and exit paths
'           : of MergeBuild and RunFilteredTests need a free operation and are covered by
'           : the scripted pass in X01.
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


Public Sub TestBeginWithPolicyRefusedWhileRunning()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim intPolicy As eDecisionPolicy
    Dim intPolicyBefore As eDecisionPolicy
    Dim intMode As eInteractionMode
    Dim blnBlocked As Boolean
    Dim intLevel As eErrorLevel

    ' A private instance, so the result does not depend on whether the session operation
    ' happens to hold the host run's root. Through the MCP the runner lives in the
    ' installed add-in and this project's session operation only adopts that root from
    ' the registry when it is created after the root begins.
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "the first root is granted"
    intMode = cOp.InteractionMode
    intPolicyBefore = cOp.DecisionPolicy
    blnBlocked = cOp.DecisionBlocked
    intLevel = cOp.ErrorLevel

    For intPolicy = edpBlock To edpDecline
        TestAssert Not cOp.Begin(eotOther, intPolicy), _
            "a policy request is refused while an operation is running (policy " & intPolicy & ")"
        TestAssert (cOp.InteractionMode = intMode) And (cOp.DecisionBlocked = blnBlocked) _
            And (cOp.ErrorLevel = intLevel) And (cOp.DecisionPolicy = intPolicyBefore), _
            "a refused policy request leaves mode, policy, blocked flag and error level alone (policy " & intPolicy & ")"
    Next intPolicy
    TestAssert cOp.Status = eosRunning, "the running operation is not disturbed"
    TestAssert cRoot.IsValid, "the running operation keeps its lease"

    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "the first root still completes"

End Sub


Public Sub TestPolicyRequestRefusedWhileRunning()

    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strJson As String
    Dim intModeBefore As eInteractionMode
    Dim intPolicyBefore As eDecisionPolicy

    ' SetOperationPolicy asks PolicyRequestRefusal about the session operation. A private
    ' instance stands in for it, so the running state is set up here and never depends on
    ' where the host run started.
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "the root is granted"
    intModeBefore = cOp.InteractionMode
    intPolicyBefore = cOp.DecisionPolicy

    strJson = PolicyRequestRefusal(cOp)
    TestAssert InStr(strJson, """success"":false") > 0, "setting a policy while an operation runs fails"
    TestAssert InStr(strJson, "operation_already_running") > 0, "the refusal names operation_already_running"
    TestAssert (cOp.InteractionMode = intModeBefore) And (cOp.DecisionPolicy = intPolicyBefore), _
        "the refused request did not touch the running operation"

    cRoot.Complete eorSuccess

End Sub


Public Sub TestPolicyRequestAcceptedWhenReady()

    Dim cOp As clsOperation

    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    TestAssert cOp.Status = eosReady, "a new operation is ready"
    TestAssert Len(PolicyRequestRefusal(cOp)) = 0, "a ready operation is not refused"

End Sub


Public Sub TestClearOperationPolicyIsHarmlessWhileRunning()

    TestAssert InStr(VCS.ClearOperationPolicy, """success"":true") > 0, "clearing a policy that was never set succeeds"
    TestAssert InStr(VCS.ClearOperationPolicy, """success"":true") > 0, "clearing twice succeeds"
    TestAssert LiveStateUnchanged, "clearing an unset policy leaves the running operation alone"

End Sub

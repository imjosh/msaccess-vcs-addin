VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestInteractionScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestInteractionScope
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Tests for clsInteractionScope: nesting, release by owner token, silent
'           : requests inside a noninteractive scope, error-level restore, policy
'           : answers, decision recording, and decision JSON. Every test builds its
'           : own instance, so the live operation and registry state are never touched.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestNewScopeStartsNormalWithNoDecisions()

    Dim cScope As New clsInteractionScope

    TestAssert cScope.InteractionMode = eimNormal, "a new scope is in normal mode"
    TestAssert cScope.DecisionPolicy = edpAsk, "a new scope asks"
    TestAssert Not cScope.Blocked, "a new scope is not blocked"
    TestAssert cScope.DecisionCount = 0, "a new scope has no decisions"
    TestAssert ConvertToJson(cScope.Decisions) = "[]", "no decisions is an empty JSON array"
    TestAssert Not cScope.InNonInteractiveScope, "no scope is open"

End Sub


Public Sub TestOpenAppliesAndCloseRestores()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
    TestAssert lngToken <> 0, "open returns a token"
    TestAssert cScope.InteractionMode = eimNonInteractive, "open applies the mode"
    TestAssert cScope.DecisionPolicy = edpBlock, "open applies the policy"
    TestAssert cScope.IsOpen(lngToken), "the scope is open"

    cScope.CloseScope lngToken
    TestAssert cScope.InteractionMode = eimNormal, "close restores the mode"
    TestAssert cScope.DecisionPolicy = edpAsk, "close restores the policy"
    TestAssert Not cScope.IsOpen(lngToken), "the scope is closed"
    TestAssert Not cScope.InNonInteractiveScope, "no noninteractive scope remains"

End Sub


Public Sub TestCloseRestoresPriorSilentMode()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long

    cScope.InteractionMode = eimSilent
    lngToken = cScope.OpenScope(eimNonInteractive, edpDecline)
    cScope.CloseScope lngToken
    TestAssert cScope.InteractionMode = eimSilent, "close returns to the silent mode that was set before"

End Sub


Public Sub TestNestedScopesCloseInnerFirst()

    Dim cScope As New clsInteractionScope
    Dim lngOuter As Long
    Dim lngInner As Long

    lngOuter = cScope.OpenScope(eimNonInteractive, edpBlock)
    lngInner = cScope.OpenScope(eimNonInteractive, edpPreferSource)
    TestAssert cScope.DecisionPolicy = edpPreferSource, "inner policy applies"

    cScope.CloseScope lngInner
    TestAssert cScope.DecisionPolicy = edpBlock, "closing the inner scope returns to the outer policy"
    TestAssert cScope.InteractionMode = eimNonInteractive, "outer scope is still noninteractive"

    cScope.CloseScope lngOuter
    TestAssert cScope.InteractionMode = eimNormal, "closing the outer scope returns to normal"
    TestAssert cScope.DecisionPolicy = edpAsk, "closing the outer scope returns to the base policy"

End Sub


Public Sub TestClosingOuterDiscardsInnerScopes()

    Dim cScope As New clsInteractionScope
    Dim lngOuter As Long
    Dim lngInner As Long

    lngOuter = cScope.OpenScope(eimNonInteractive, edpBlock)
    lngInner = cScope.OpenScope(eimNonInteractive, edpPreferSource)

    cScope.CloseScope lngOuter
    TestAssert cScope.InteractionMode = eimNormal, "out-of-order close lands on the state before the outer scope"
    TestAssert cScope.DecisionPolicy = edpAsk, "out-of-order close restores the base policy"
    TestAssert Not cScope.IsOpen(lngInner), "the inner scope was discarded"

    cScope.CloseScope lngInner
    TestAssert cScope.InteractionMode = eimNormal, "closing a discarded token does nothing"

End Sub


Public Sub TestCloseTwiceDoesNotDisturbALaterScope()

    Dim cScope As New clsInteractionScope
    Dim lngFirst As Long
    Dim lngSecond As Long

    lngFirst = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.CloseScope lngFirst
    lngSecond = cScope.OpenScope(eimNonInteractive, edpDecline)

    cScope.CloseScope lngFirst
    TestAssert cScope.IsOpen(lngSecond), "a repeated close does not close another scope"
    TestAssert cScope.DecisionPolicy = edpDecline, "a repeated close leaves the later policy in force"

    cScope.CloseScope lngSecond
    TestAssert cScope.InteractionMode = eimNormal, "the later scope still closes by its own token"

End Sub


Public Sub TestUnknownTokenIsANoOp()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long

    cScope.CloseScope 0
    cScope.CloseScope 12345
    TestAssert cScope.InteractionMode = eimNormal, "unknown token with nothing open changes nothing"

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.CloseScope 0
    cScope.CloseScope lngToken + 1000
    TestAssert cScope.IsOpen(lngToken), "an unknown token does not close an open scope"
    TestAssert cScope.InteractionMode = eimNonInteractive, "an unknown token leaves the mode alone"

    cScope.CloseScope lngToken

End Sub


Public Sub TestReleaseClosesOnlyOperationOwnedScopes()

    Dim cScope As New clsInteractionScope
    Dim lngCaller As Long
    Dim lngOwned As Long

    lngCaller = cScope.OpenScope(eimNonInteractive, edpDecline, False)
    lngOwned = cScope.OpenScope(eimNonInteractive, edpBlock, True)

    cScope.ReleaseOperationScopes
    TestAssert Not cScope.IsOpen(lngOwned), "the operation-owned scope is released"
    TestAssert cScope.IsOpen(lngCaller), "the caller-owned scope stays open"
    TestAssert cScope.DecisionPolicy = edpDecline, "the caller's policy is back in force"

    cScope.ReleaseOperationScopes
    TestAssert cScope.IsOpen(lngCaller), "releasing again still leaves the caller-owned scope"

    cScope.CloseScope lngCaller
    TestAssert cScope.InteractionMode = eimNormal, "the caller closes its own scope"
    TestAssert cScope.DecisionPolicy = edpAsk, "the caller's close restores the base state"

End Sub


Public Sub TestReleaseKeepsCallerScopeOpenedInsideOwnedScope()

    Dim cScope As New clsInteractionScope
    Dim lngOwned As Long
    Dim lngCaller As Long

    lngOwned = cScope.OpenScope(eimNonInteractive, edpBlock, True)
    lngCaller = cScope.OpenScope(eimNonInteractive, edpPreferSource, False)

    cScope.ReleaseOperationScopes
    TestAssert Not cScope.IsOpen(lngOwned), "the operation-owned scope is released"
    TestAssert cScope.IsOpen(lngCaller), "the caller-owned scope stays open"
    TestAssert cScope.DecisionPolicy = edpPreferSource, "the caller's policy is untouched by the release"

    cScope.CloseScope lngCaller
    TestAssert cScope.InteractionMode = eimNormal, "closing the caller's scope lands on the state before the released scope"
    TestAssert cScope.DecisionPolicy = edpAsk, "the base policy is restored"

End Sub


Public Sub TestReleaseWithNothingOwnedIsANoOp()

    Dim cScope As New clsInteractionScope

    cScope.ReleaseOperationScopes
    TestAssert cScope.InteractionMode = eimNormal, "release with no scopes changes nothing"

End Sub


Public Sub TestSilentRequestOutsideNoninteractiveScope()

    Dim cScope As New clsInteractionScope

    cScope.InteractionMode = eimSilent
    TestAssert cScope.InteractionMode = eimSilent, "silent applies outside a noninteractive scope"
    cScope.InteractionMode = eimNormal
    TestAssert cScope.InteractionMode = eimNormal, "normal applies outside a noninteractive scope"

End Sub


Public Sub TestSilentRequestDoesNotWeakenNoninteractiveScope()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)

    cScope.InteractionMode = eimSilent
    TestAssert cScope.InteractionMode = eimNonInteractive, "a silent request leaves the mode noninteractive"
    cScope.InteractionMode = eimNormal
    TestAssert cScope.InteractionMode = eimNonInteractive, "a normal request leaves the mode noninteractive"
    TestAssert cScope.DecisionPolicy = edpBlock, "the policy is unchanged"

    cScope.CloseScope lngToken
    TestAssert cScope.InteractionMode = eimNormal, "close restores the mode saved at open"

End Sub


Public Sub TestSilentScopeInsideNoninteractiveScopeKeepsPolicy()

    Dim cScope As New clsInteractionScope
    Dim lngOuter As Long
    Dim lngInner As Long

    lngOuter = cScope.OpenScope(eimNonInteractive, edpBlock)
    lngInner = cScope.OpenScope(eimSilent, edpPreferSource)
    TestAssert cScope.InteractionMode = eimNonInteractive, "a silent scope does not weaken the enclosing scope"
    TestAssert cScope.DecisionPolicy = edpBlock, "a silent scope does not replace the enclosing policy"

    cScope.CloseScope lngInner
    TestAssert cScope.InteractionMode = eimNonInteractive, "closing the silent scope keeps the outer mode"
    TestAssert cScope.InNonInteractiveScope, "the outer scope is still open"

    cScope.CloseScope lngOuter

End Sub


Public Sub TestPolicyCannotBeChangedInsideNoninteractiveScope()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.DecisionPolicy = edpPreferSource
    TestAssert cScope.DecisionPolicy = edpBlock, "a direct policy change is ignored inside the scope"

    cScope.CloseScope lngToken
    cScope.DecisionPolicy = edpPreferSource
    TestAssert cScope.DecisionPolicy = edpPreferSource, "the policy can be set outside a noninteractive scope"

End Sub


Public Sub TestCloseRestoresErrorLevel()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long

    cScope.ErrorLevel = eelWarning
    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
    TestAssert cScope.ErrorLevel = eelWarning, "opening a scope keeps the current error level"

    cScope.ErrorLevel = eelError
    cScope.CloseScope lngToken
    TestAssert cScope.ErrorLevel = eelWarning, "close restores the error level saved at open"

End Sub


Public Sub TestNestedBlockedPromptDoesNotPoisonOuterScope()

    Dim cScope As New clsInteractionScope
    Dim lngOuter As Long
    Dim lngInner As Long

    lngOuter = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.RecordDecision NewDecision("acknowledged", "Outer", "note", "acknowledged")
    lngInner = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.ResolvePrompt vbYesNo, "Inner", "Delete everything?"
    TestAssert cScope.Blocked, "the inner scope is blocked"
    TestAssert cScope.ErrorLevel = eelCritical, "the inner block raises the error level"

    cScope.CloseScope lngInner
    TestAssert Not cScope.Blocked, "the outer scope is not blocked by the inner prompt"
    TestAssert cScope.ErrorLevel = eelNoError, "the outer error level is restored"
    TestAssert cScope.DecisionCount = 1, "the outer scope keeps its own journal"

    cScope.CloseScope lngOuter

End Sub


Public Sub TestOkOnlyIsAcknowledgedUnderEveryPolicy()

    Dim varPolicies As Variant
    Dim varPolicy As Variant
    Dim cScope As New clsInteractionScope
    Dim lngToken As Long
    Dim cOutcome As clsPromptOutcome

    varPolicies = AllPolicies
    For Each varPolicy In varPolicies
        lngToken = cScope.OpenScope(eimNonInteractive, varPolicy)
        Set cOutcome = cScope.ResolvePrompt(vbOKOnly + vbInformation, "Info", "Done")
        TestAssert cOutcome.Result = vbOK, PolicyLabel(varPolicy) & ": OK-only returns OK"
        TestAssert cOutcome.Resolution = "acknowledged", PolicyLabel(varPolicy) & ": OK-only is acknowledged"
        TestAssert Not cScope.Blocked, PolicyLabel(varPolicy) & ": OK-only never blocks"
        cScope.CloseScope lngToken
    Next varPolicy

End Sub


Public Sub TestEveryPolicyAgainstEveryConfirmationStyle()

    ' The expected answers are the policy table in docs/noninteractive-dialogs.md, written
    ' out as data rather than derived from the code under test. block, prefer_source,
    ' prefer_database and skip leave a confirmation as decision_required and blocked:
    ' conflict policies never answer a generic confirmation. Only decline answers it.
    ' Whatever the policy, the button returned is the refusing one for the style:
    ' Cancel, Abort or No.
    Dim varPolicies As Variant
    Dim varResolutions As Variant
    Dim varBlocked As Variant
    Dim varStyles As Variant
    Dim varRefusals As Variant
    Dim lngPolicy As Long
    Dim lngStyle As Long
    Dim cScope As New clsInteractionScope
    Dim lngToken As Long
    Dim cOutcome As clsPromptOutcome
    Dim strLabel As String

    varPolicies = Array(edpBlock, edpPreferSource, edpPreferDatabase, edpSkip, edpDecline)
    varResolutions = Array("decision_required", "decision_required", "decision_required", "decision_required", "declined")
    varBlocked = Array(True, True, True, True, False)
    varStyles = Array(vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel)
    varRefusals = Array(vbCancel, vbAbort, vbNo, vbNo, vbCancel)

    For lngPolicy = 0 To UBound(varPolicies)
        For lngStyle = 0 To UBound(varStyles)
            strLabel = PolicyLabel(varPolicies(lngPolicy)) & " x style " & CStr(varStyles(lngStyle))
            lngToken = cScope.OpenScope(eimNonInteractive, varPolicies(lngPolicy))

            Set cOutcome = cScope.ResolvePrompt(varStyles(lngStyle) + vbQuestion, "Confirm", "Proceed?")
            TestAssert cOutcome.Result = varRefusals(lngStyle), strLabel & ": answers with the refusing button"
            TestAssert cOutcome.Resolution = varResolutions(lngPolicy), strLabel & ": resolution matches the documented table"
            TestAssert cOutcome.Blocked = varBlocked(lngPolicy), strLabel & ": outcome blocked flag matches the documented table"
            TestAssert cScope.Blocked = varBlocked(lngPolicy), strLabel & ": scope blocked flag matches the documented table"
            TestAssert cScope.DecisionCount = 1, strLabel & ": records the prompt"

            cScope.CloseScope lngToken
        Next lngStyle
    Next lngPolicy

End Sub


Public Sub TestConflictActionForEachPolicyAndRequestedAction()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long
    Dim varRequested As Variant
    Dim varPolicy As Variant

    For Each varRequested In Array(ercNone, ercSkip, ercOverwrite, ercDelete)
        lngToken = cScope.OpenScope(eimNonInteractive, edpPreferSource)
        If varRequested = ercNone Then
            TestAssert ConflictActionForPolicy(cScope.DecisionPolicy, varRequested) = ercOverwrite, _
                "prefer_source overwrites when the conflict asks for nothing"
        Else
            TestAssert ConflictActionForPolicy(cScope.DecisionPolicy, varRequested) = varRequested, _
                "prefer_source applies the conflict's own action (" & CStr(varRequested) & ")"
        End If
        cScope.CloseScope lngToken

        For Each varPolicy In Array(edpPreferDatabase, edpSkip, edpDecline)
            lngToken = cScope.OpenScope(eimNonInteractive, varPolicy)
            TestAssert ConflictActionForPolicy(cScope.DecisionPolicy, varRequested) = ercSkip, _
                PolicyLabel(varPolicy) & " keeps the database object (" & CStr(varRequested) & ")"
            cScope.CloseScope lngToken
        Next varPolicy

        lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
        TestAssert ConflictActionForPolicy(cScope.DecisionPolicy, varRequested) = ercNone, _
            "block does not resolve a conflict (" & CStr(varRequested) & ")"
        cScope.CloseScope lngToken
    Next varRequested

End Sub


Public Sub TestDecisionRequiredIsRecordedAndBlocks()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long
    Dim colDecisions As Collection

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.RecordDecision NewDecision("decision_required", "Select Source Folder", "No folder", "blocked", "picker")

    TestAssert cScope.Blocked, "decision_required blocks the scope"
    TestAssert cScope.ErrorLevel = eelCritical, "decision_required raises the error level to critical"
    TestAssert cScope.DecisionCount = 1, "the decision is recorded"

    Set colDecisions = ParseJson(ConvertToJson(cScope.Decisions))
    TestAssert colDecisions(1)("kind") = "decision_required", "kind is reported"
    TestAssert colDecisions(1)("title") = "Select Source Folder", "title is reported"
    TestAssert colDecisions(1)("message") = "No folder", "message is reported"
    TestAssert colDecisions(1)("resolution") = "blocked", "resolution is reported"
    TestAssert colDecisions(1)("detail") = "picker", "detail is reported"

    cScope.CloseScope lngToken
    TestAssert Not cScope.Blocked, "closing the scope clears the block"
    TestAssert cScope.DecisionCount = 0, "closing the scope restores the outer journal"

End Sub


Public Sub TestOtherDecisionKindsDoNotBlock()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long

    lngToken = cScope.OpenScope(eimNonInteractive, edpDecline)
    cScope.RecordDecision NewDecision("acknowledged", "Info", "Done", "acknowledged")
    cScope.RecordDecision NewDecision("applied", "Merge conflicts", "2 conflict(s)", "keep_database")

    TestAssert Not cScope.Blocked, "acknowledged and applied do not block"
    TestAssert cScope.ErrorLevel = eelNoError, "acknowledged and applied leave the error level alone"
    TestAssert cScope.DecisionCount = 2, "both decisions are recorded"

    cScope.CloseScope lngToken

End Sub


Public Sub TestDecisionsJsonIsValidForControlCharacters()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long
    Dim strTitle As String
    Dim strMessage As String
    Dim strJson As String
    Dim colDecisions As Collection
    Dim lngIdx As Long

    strTitle = "Say ""hi"" to C:\temp\new"
    strMessage = "line1" & vbCr & "line2" & vbLf & "line3" & vbCrLf & "tab" & vbTab & "end" & _
        Chr$(0) & Chr$(1) & Chr$(8) & Chr$(12) & Chr$(31) & Chr$(127) & "/ unicode " & ChrW$(233) & ChrW$(8364)

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.RecordDecision NewDecision("decision_required", strTitle, strMessage, "blocked", strMessage)
    strJson = ConvertToJson(cScope.Decisions)
    cScope.CloseScope lngToken

    For lngIdx = 1 To Len(strJson)
        TestAssert AscW(Mid$(strJson, lngIdx, 1)) >= 32, "the JSON contains no raw control character (position " & lngIdx & ")"
    Next lngIdx

    Set colDecisions = ParseJson(strJson)
    TestAssert colDecisions.Count = 1, "the JSON parses back to one decision"
    TestAssert colDecisions(1)("title") = strTitle, "quotes and backslashes round-trip"
    TestAssert colDecisions(1)("message") = strMessage, "control characters round-trip"
    TestAssert colDecisions(1)("detail") = strMessage, "detail round-trips"

End Sub


Public Sub TestPromptDecisionTextRoundTripsThroughJson()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long
    Dim colDecisions As Collection
    Dim strMessage As String

    strMessage = "Overwrite ""x""?" & vbCrLf & "Path: C:\a\b" & vbTab & Chr$(7)

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock)
    cScope.ResolvePrompt vbYesNo, "Confirm" & vbLf, strMessage
    Set colDecisions = ParseJson(ConvertToJson(cScope.Decisions))
    cScope.CloseScope lngToken

    TestAssert colDecisions(1)("message") = strMessage, "a blocked prompt's text round-trips"
    TestAssert colDecisions(1)("title") = "Confirm" & vbLf, "a blocked prompt's title round-trips"
    TestAssert colDecisions(1)("kind") = "decision_required", "the prompt is recorded as decision_required"

End Sub


Public Sub TestScopesDoNotTouchTheLiveOperation()

    Dim cScope As New clsInteractionScope
    Dim lngToken As Long
    Dim eimBefore As eInteractionMode
    Dim edpBefore As eDecisionPolicy
    Dim blnBlockedBefore As Boolean
    Dim eelBefore As eErrorLevel

    eimBefore = Operation.InteractionMode
    edpBefore = Operation.DecisionPolicy
    blnBlockedBefore = Operation.DecisionBlocked
    eelBefore = Operation.ErrorLevel

    lngToken = cScope.OpenScope(eimNonInteractive, edpBlock, True)
    cScope.ResolvePrompt vbYesNo, "Confirm", "Proceed?"
    cScope.ReleaseOperationScopes
    cScope.CloseScope lngToken

    TestAssert Operation.InteractionMode = eimBefore, "the live interaction mode is unchanged"
    TestAssert Operation.DecisionPolicy = edpBefore, "the live policy is unchanged"
    TestAssert Operation.DecisionBlocked = blnBlockedBefore, "the live blocked flag is unchanged"
    TestAssert Operation.ErrorLevel = eelBefore, "the live error level is unchanged"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AllPolicies
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Every noninteractive policy.
'---------------------------------------------------------------------------------------
'
Private Function AllPolicies() As Variant
    AllPolicies = Array(edpBlock, edpPreferSource, edpPreferDatabase, edpSkip, edpDecline)
End Function


'---------------------------------------------------------------------------------------
' Procedure : PolicyLabel
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Policy name for assertion messages.
'---------------------------------------------------------------------------------------
'
Private Function PolicyLabel(ByVal intPolicy As eDecisionPolicy) As String

    Select Case intPolicy
        Case edpBlock: PolicyLabel = "block"
        Case edpPreferSource: PolicyLabel = "prefer_source"
        Case edpPreferDatabase: PolicyLabel = "prefer_database"
        Case edpSkip: PolicyLabel = "skip"
        Case edpDecline: PolicyLabel = "decline"
        Case Else: PolicyLabel = "policy " & CStr(intPolicy)
    End Select

End Function

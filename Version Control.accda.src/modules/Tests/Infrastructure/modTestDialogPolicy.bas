Attribute VB_Name = "modTestDialogPolicy"
'---------------------------------------------------------------------------------------
' Module    : modTestDialogPolicy
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : Decision-policy rules for noninteractive prompts and merge conflicts.
'           : These tests do not open dialogs.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestParseDecisionPolicyNames()

    Dim intPolicy As eDecisionPolicy

    TestAssert ParseDecisionPolicy("block", intPolicy), "block is a policy"
    TestAssert intPolicy = edpBlock, "block maps to edpBlock"
    TestAssert ParseDecisionPolicy("prefer_source", intPolicy), "prefer_source is a policy"
    TestAssert intPolicy = edpPreferSource, "prefer_source maps"
    TestAssert ParseDecisionPolicy("PREFER_DATABASE", intPolicy), "policy names are case-insensitive"
    TestAssert intPolicy = edpPreferDatabase, "prefer_database maps"
    TestAssert ParseDecisionPolicy("skip", intPolicy), "skip is a policy"
    TestAssert intPolicy = edpSkip, "skip maps"
    TestAssert ParseDecisionPolicy("decline", intPolicy), "decline is a policy"
    TestAssert intPolicy = edpDecline, "decline maps"
    TestAssert Not ParseDecisionPolicy("accept_all", intPolicy), "unknown policy is rejected"
    TestAssert Not ParseDecisionPolicy("", intPolicy), "empty policy is rejected"

End Sub


Public Sub TestPolicyNamesRoundTripThroughTheTable()

    Dim intPolicy As eDecisionPolicy
    Dim varName As Variant

    For Each varName In Array("block", "prefer_source", "prefer_database", "skip", "decline")
        TestAssert ParseDecisionPolicy(CStr(varName), intPolicy), varName & " parses"
        TestAssert DecisionPolicyName(intPolicy) = varName, varName & " names itself"
    Next varName
    TestAssert Len(DecisionPolicyName(edpAsk)) = 0, "edpAsk is not a named policy"
    TestAssert InStr(InvalidPolicyMessage(), "block, prefer_source, prefer_database, skip, or decline") > 0, _
        "the invalid-policy message lists every name in table order"
    TestAssert ConflictResolutionName(edpPreferSource) = "prefer_source", "prefer_source reports its own name"
    TestAssert ConflictResolutionName(edpSkip) = "keep_database", "other policies keep the database object"

End Sub


Public Sub TestDecisionWireValuesAreUnchanged()

    ' Pinned to literals on purpose: these strings are the wire contract.
    TestAssert DECISION_REQUIRED = "decision_required", "decision_required"
    TestAssert ERR_DECISION_REQUIRED = "decision_required", "the error pattern shares it"
    TestAssert DECISION_ACKNOWLEDGED = "acknowledged", "acknowledged"
    TestAssert DECISION_APPLIED = "applied", "applied"
    TestAssert DECISION_DECLINED = "declined", "declined"
    TestAssert DECISION_BLOCKED = "blocked", "blocked"
    TestAssert DECISION_KEEP_DATABASE = "keep_database", "keep_database"

End Sub


Public Sub TestNewDecisionCarriesEveryField()

    Dim cRecord As clsDecisionRecord

    Set cRecord = NewDecision("applied", "Title", "Message", "keep_database", "detail")
    TestAssert cRecord.Kind = "applied", "the kind is carried"
    TestAssert cRecord.Title = "Title", "the title is carried"
    TestAssert cRecord.Message = "Message", "the message is carried"
    TestAssert cRecord.Resolution = "keep_database", "the resolution is carried"
    TestAssert cRecord.Detail = "detail", "the detail is carried"
    TestAssert NewDecision("acknowledged", "T", "M", "acknowledged").Detail = vbNullString, "the detail is optional"

End Sub


Public Sub TestOkOnlyPromptIsAcknowledged()

    Dim cOutcome As clsPromptOutcome

    Set cOutcome = ResolveNonInteractivePrompt(vbOKOnly + vbInformation, edpBlock)
    TestAssert Not cOutcome.Blocked, "an OK-only message is not a decision"
    TestAssert cOutcome.Result = vbOK, "OK-only returns OK"
    TestAssert cOutcome.Resolution = "acknowledged", "OK-only is acknowledged"

End Sub


Public Sub TestBlockYesNoDoesNotApprove()

    Dim cOutcome As clsPromptOutcome

    Set cOutcome = ResolveNonInteractivePrompt(vbYesNo + vbQuestion, edpBlock)
    TestAssert cOutcome.Blocked, "Yes/No without an accept policy is blocked"
    TestAssert cOutcome.Result = vbNo, "the non-destructive answer is No"
    TestAssert cOutcome.Result <> vbYes, "block does not approve"
    TestAssert cOutcome.Resolution = "decision_required", "resolution is decision_required"

End Sub


Public Sub TestBlockOkCancelDoesNotApprove()

    Dim cOutcome As clsPromptOutcome

    Set cOutcome = ResolveNonInteractivePrompt(vbOKCancel + vbExclamation, edpBlock)
    TestAssert cOutcome.Blocked, "OK/Cancel is a decision"
    TestAssert cOutcome.Result = vbCancel, "block returns Cancel"
    TestAssert cOutcome.Result <> vbOK, "block does not click OK"

End Sub


Public Sub TestDeclineYesNoIsApplied()

    Dim cOutcome As clsPromptOutcome

    Set cOutcome = ResolveNonInteractivePrompt(vbYesNo, edpDecline)
    TestAssert Not cOutcome.Blocked, "decline is an explicit answer"
    TestAssert cOutcome.Result = vbNo, "decline answers No"
    TestAssert cOutcome.Resolution = "declined", "resolution is declined"

End Sub


Public Sub TestConflictPolicyDoesNotAnswerGenericPrompt()

    Dim cOutcome As clsPromptOutcome

    Set cOutcome = ResolveNonInteractivePrompt(vbYesNo, edpPreferSource)
    TestAssert cOutcome.Blocked, "prefer_source does not answer a generic confirmation"
    TestAssert cOutcome.Result = vbNo, "generic confirmation stays unapproved"

End Sub


Public Sub TestAbortRetryDeclineStops()

    Dim cOutcome As clsPromptOutcome

    Set cOutcome = ResolveNonInteractivePrompt(vbAbortRetryIgnore, edpDecline)
    TestAssert Not cOutcome.Blocked, "decline covers Abort/Retry/Ignore"
    TestAssert cOutcome.Result = vbAbort, "decline aborts rather than retrying"

End Sub


Public Sub TestConflictActionForPolicy()

    TestAssert ConflictActionForPolicy(edpPreferSource) = ercOverwrite, "prefer_source overwrites from source"
    TestAssert ConflictActionForPolicy(edpPreferDatabase) = ercSkip, "prefer_database keeps the database object"
    TestAssert ConflictActionForPolicy(edpSkip) = ercSkip, "skip keeps the database object"
    TestAssert ConflictActionForPolicy(edpDecline) = ercSkip, "decline does not overwrite"
    TestAssert ConflictActionForPolicy(edpBlock) = ercNone, "block does not pick a resolution"

End Sub


Public Sub TestRefusalJsonShape()

    Dim dResult As Dictionary

    Set dResult = ParseJson(RefusalJson(ERR_INVALID_DECISION_POLICY, InvalidPolicyMessage(), False))
    TestAssert dResult("success") = False, "a refusal is not a success"
    TestAssert dResult("error_pattern") = "invalid_decision_policy", "the refusal carries its error pattern"
    TestAssert InStr(dResult("error"), "prefer_source") > 0, "the message lists the valid policy names"
    TestAssert Not dResult.Exists("started"), "a refusal has no started marker"

    Set dResult = ParseJson(RefusalJson(vbNullString, "not started", False))
    TestAssert Not dResult.Exists("error_pattern"), "an empty pattern is left out"

End Sub


Public Sub TestStartedJsonIsNotAFinalResult()

    Dim dResult As Dictionary

    Set dResult = ParseJson(StartedJson())
    TestAssert dResult("success") = True, "a start result reports success"
    TestAssert dResult("started") = True, "a start result is marked started"
    TestAssert Not dResult.Exists("error_pattern"), "a start result has no error pattern"
    TestAssert Not dResult.Exists("decisions"), "a start result has no decisions"

End Sub


Public Sub TestRuntimeErrorJsonEncodesAnyText()

    Dim dResult As Dictionary
    Dim strText As String
    Dim blnBlocked As Boolean

    ' RuntimeErrorJson reads the session operation's blocked flag. Clear it for
    ' the call so the result does not depend on the host run, then restore it.
    blnBlocked = Operation.DecisionBlocked
    Operation.DecisionBlocked = False
    strText = "bad ""quote"" \ back" & vbCrLf & "tab" & vbTab & "bell" & Chr$(7)
    Set dResult = ParseJson(RuntimeErrorJson(strText, 5))
    Operation.DecisionBlocked = blnBlocked
    TestAssert dResult("success") = False, "a runtime error is not a success"
    TestAssert dResult("error") = strText, "the error text survives encoding"
    TestAssert dResult("errorNumber") = 5, "the error number is included"

End Sub


Public Sub TestMCPCallBeforeRootKeepsLoggedErrorsOutOfBoxes()

    ' Log.Error skips its message box only in the gap between an MCP call
    ' registering its callback and its root beginning.
    TestAssert MCPCallBeforeRoot(True, False), "MCP registered, no root yet"
    TestAssert Not MCPCallBeforeRoot(True, True), "inside the root the policy scope decides"
    TestAssert Not MCPCallBeforeRoot(False, False), "no MCP call: interactive rules apply"
    TestAssert Not MCPCallBeforeRoot(False, True), "a root without MCP is unchanged"

End Sub

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


Public Sub TestOkOnlyPromptIsAcknowledged()

    Dim intResult As VbMsgBoxResult
    Dim strResolution As String
    Dim blnBlocked As Boolean

    ResolveNonInteractivePrompt vbOKOnly + vbInformation, edpBlock, intResult, strResolution, blnBlocked
    TestAssert Not blnBlocked, "an OK-only message is not a decision"
    TestAssert intResult = vbOK, "OK-only returns OK"
    TestAssert strResolution = "acknowledged", "OK-only is acknowledged"

End Sub


Public Sub TestBlockYesNoDoesNotApprove()

    Dim intResult As VbMsgBoxResult
    Dim strResolution As String
    Dim blnBlocked As Boolean

    ResolveNonInteractivePrompt vbYesNo + vbQuestion, edpBlock, intResult, strResolution, blnBlocked
    TestAssert blnBlocked, "Yes/No without an accept policy is blocked"
    TestAssert intResult = vbNo, "the non-destructive answer is No"
    TestAssert intResult <> vbYes, "block does not approve"
    TestAssert strResolution = "decision_required", "resolution is decision_required"

End Sub


Public Sub TestBlockOkCancelDoesNotApprove()

    Dim intResult As VbMsgBoxResult
    Dim strResolution As String
    Dim blnBlocked As Boolean

    ResolveNonInteractivePrompt vbOKCancel + vbExclamation, edpBlock, intResult, strResolution, blnBlocked
    TestAssert blnBlocked, "OK/Cancel is a decision"
    TestAssert intResult = vbCancel, "block returns Cancel"
    TestAssert intResult <> vbOK, "block does not click OK"

End Sub


Public Sub TestDeclineYesNoIsApplied()

    Dim intResult As VbMsgBoxResult
    Dim strResolution As String
    Dim blnBlocked As Boolean

    ResolveNonInteractivePrompt vbYesNo, edpDecline, intResult, strResolution, blnBlocked
    TestAssert Not blnBlocked, "decline is an explicit answer"
    TestAssert intResult = vbNo, "decline answers No"
    TestAssert strResolution = "declined", "resolution is declined"

End Sub


Public Sub TestConflictPolicyDoesNotAnswerGenericPrompt()

    Dim intResult As VbMsgBoxResult
    Dim strResolution As String
    Dim blnBlocked As Boolean

    ResolveNonInteractivePrompt vbYesNo, edpPreferSource, intResult, strResolution, blnBlocked
    TestAssert blnBlocked, "prefer_source does not answer a generic confirmation"
    TestAssert intResult = vbNo, "generic confirmation stays unapproved"

End Sub


Public Sub TestAbortRetryDeclineStops()

    Dim intResult As VbMsgBoxResult
    Dim strResolution As String
    Dim blnBlocked As Boolean

    ResolveNonInteractivePrompt vbAbortRetryIgnore, edpDecline, intResult, strResolution, blnBlocked
    TestAssert Not blnBlocked, "decline covers Abort/Retry/Ignore"
    TestAssert intResult = vbAbort, "decline aborts rather than retrying"

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

    strText = "bad ""quote"" \ back" & vbCrLf & "tab" & vbTab & "bell" & Chr$(7)
    Set dResult = ParseJson(RuntimeErrorJson(strText, 5))
    TestAssert dResult("success") = False, "a runtime error is not a success"
    TestAssert dResult("error") = strText, "the error text survives encoding"
    TestAssert dResult("errorNumber") = 5, "the error number is included"

End Sub

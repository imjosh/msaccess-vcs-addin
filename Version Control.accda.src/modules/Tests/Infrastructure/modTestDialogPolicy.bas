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

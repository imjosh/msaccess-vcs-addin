Attribute VB_Name = "modTestAsyncTerminalCallback"
'---------------------------------------------------------------------------------------
' Module    : modTestAsyncTerminalCallback
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : Tests for the terminal callback the timer posts when a launched call
'           : finished without one (modDialogPolicy.PostUnreportedOutcome).
'           : Each test posts through a private clsMCP registered to an unreachable
'           : host, so nothing reaches the MCP server that may be hosting this run.
'           : Run: ?VCS.RunTests("modTestAsyncTerminalCallback")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


' Fails at once: the malformed host is rejected before any connection is attempted.
Private Const UNREACHABLE_CALLBACK As String = "http://invalid..host/callback"


'---------------------------------------------------------------------------------------
' Procedure : NewActiveMCP
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : A private MCP instance with a registered callback.
'---------------------------------------------------------------------------------------
'
Private Function NewActiveMCP() As clsMCP
    Dim cMCP As clsMCP
    Set cMCP = New clsMCP
    cMCP.RegisterCallback "{""callback_url"":""" & UNREACHABLE_CALLBACK & """,""operation_id"":""a27-unit""}"
    Set NewActiveMCP = cMCP
End Function


'---------------------------------------------------------------------------------------
' Procedure : TerminalPosts
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : How many complete, error or cancelled callbacks an MCP instance attempted.
'---------------------------------------------------------------------------------------
'
Private Function TerminalPosts(cMCP As clsMCP) As Long
    Dim dData As Dictionary
    Set dData = cMCP.GetDiagnosticData
    TerminalPosts = dData("terminalPostAttempts")
End Function


'---------------------------------------------------------------------------------------
' Procedure : FakePreRootFailure
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : What a method returns when it fails before its root exists, as
'           : ExecuteTests does when installing the helper raises. The session
'           : operation's blocked flag is cleared for the call so the shape does not
'           : depend on the run hosting this test.
'---------------------------------------------------------------------------------------
'
Private Function FakePreRootFailure() As String
    Dim blnBlocked As Boolean
    blnBlocked = Operation.DecisionBlocked
    Operation.DecisionBlocked = False
    FakePreRootFailure = RuntimeErrorJson("Method 'Name' of object '_VBComponent' failed", 32813)
    Operation.DecisionBlocked = blnBlocked
End Function


Public Sub TestPreRootRuntimeErrorPostsOneErrorCallback()

    Dim cMCP As clsMCP
    Dim strResult As String
    Dim dPayload As Dictionary
    Dim strMessage As String

    strResult = FakePreRootFailure()
    Set cMCP = NewActiveMCP
    TestAssert Not cMCP.TerminalPosted, "a new registration has posted nothing"

    PostUnreportedOutcome cMCP, "RunFilteredTests", strResult, False
    TestAssert TerminalPosts(cMCP) = 1, "the unreported failure is posted"
    TestAssert cMCP.TerminalPosted, "the post is recorded"

    PostUnreportedOutcome cMCP, "RunFilteredTests", strResult, False
    TestAssert TerminalPosts(cMCP) = 1, "a second call posts nothing more"

    ' The payload is the method's own result.
    Set dPayload = UnreportedOutcome("RunFilteredTests", strResult, strMessage)
    TestAssert dPayload("success") = False, "the callback is a failure"
    TestAssert dPayload("errorNumber") = 32813, "the callback carries errorNumber"
    TestAssert dPayload("runtime_error") = "Method 'Name' of object '_VBComponent' failed", _
        "the callback carries runtime_error"
    TestAssert dPayload("error") = dPayload("runtime_error"), "error keeps the method's text"
    TestAssert strMessage = dPayload("error"), "the message is the error text"

End Sub


Public Sub TestNoSecondCallbackAfterARootOrRefusalPosted()

    Dim cMCP As clsMCP
    Dim strRefusal As String
    Dim dRefusal As Dictionary

    ' A root's completion posts its own callback (Operation.Finish).
    Set cMCP = NewActiveMCP
    cMCP.PostCallback "complete", 100, 100, "Operation completed successfully"
    PostUnreportedOutcome cMCP, "Export", Empty, False
    TestAssert TerminalPosts(cMCP) = 1, "nothing is added after a root posted"

    ' RefusalJson posts the refusal it returns (PostRefusal).
    strRefusal = RefusalJson(ERR_MERGE_NOT_AVAILABLE, "Merge build not available", False)
    Set dRefusal = ParseJson(strRefusal)
    Set cMCP = NewActiveMCP
    cMCP.PostCallback "error", -1, -1, "Merge build not available", dRefusal
    PostUnreportedOutcome cMCP, "MergeBuild", strRefusal, False
    TestAssert TerminalPosts(cMCP) = 1, "nothing is added after a refusal posted"

    ' Root completion releases the session MCP; its replacement is not active.
    Set cMCP = New clsMCP
    PostUnreportedOutcome cMCP, "Build", Empty, False
    TestAssert TerminalPosts(cMCP) = 0, "an inactive MCP posts nothing"

    ' A root the call left staged posts when its continuation finishes.
    Set cMCP = NewActiveMCP
    PostUnreportedOutcome cMCP, "Build", Empty, True
    TestAssert TerminalPosts(cMCP) = 0, "a staged root is left to post its own outcome"

    ' Streaming callbacks are not an outcome, and a new registration starts over.
    Set cMCP = NewActiveMCP
    cMCP.PostCallback "progress", 1, 2, "Working"
    TestAssert Not cMCP.TerminalPosted, "a progress callback is not terminal"
    cMCP.PostCallback "cancelled", -1, -1, "Operation was cancelled"
    TestAssert cMCP.TerminalPosted, "a cancelled callback is terminal"
    cMCP.RegisterCallback "{""callback_url"":""" & UNREACHABLE_CALLBACK & """,""operation_id"":""a27-next""}"
    TestAssert Not cMCP.TerminalPosted, "a new operation has posted nothing"

End Sub


Public Sub TestEveryTimerMethodReportsACallThatDidNotStart()

    Dim varMethod As Variant
    Dim cMCP As clsMCP
    Dim dPayload As Dictionary
    Dim strMessage As String
    Dim strLabel As String

    For Each varMethod In Array("Export", "FullExport", "ExportVBA", "Build", "BuildAs", _
        "MergeBuild", "RunFilteredTests")
        strLabel = " (" & varMethod & ")"
        Set cMCP = NewActiveMCP
        PostUnreportedOutcome cMCP, CStr(varMethod), Empty, False
        TestAssert TerminalPosts(cMCP) = 1, "one error callback" & strLabel

        Set dPayload = UnreportedOutcome(CStr(varMethod), Empty, strMessage)
        TestAssert dPayload("success") = False, "the callback is a failure" & strLabel
        TestAssert strMessage = T("The {0} operation did not start.", var0:=varMethod), _
            "the message says the operation did not start" & strLabel
        TestAssert dPayload("error") = strMessage, "error carries the message" & strLabel
        TestAssert Not dPayload.Exists("runtime_error"), "no runtime error is invented" & strLabel
    Next varMethod

End Sub


Public Sub TestUnreportedOutcomeKeepsTheMethodsFields()

    Dim dPayload As Dictionary
    Dim strMessage As String

    ' A refusal keeps its pattern.
    Set dPayload = UnreportedOutcome("MergeBuild", _
        RefusalJson(ERR_OPERATION_ALREADY_RUNNING, "Another operation is running", False), strMessage)
    TestAssert dPayload("error_pattern") = ERR_OPERATION_ALREADY_RUNNING, "the pattern is kept"
    TestAssert strMessage = "Another operation is running", "the refusal text is the message"
    TestAssert Not dPayload.Exists("runtime_error"), "a refusal is not a runtime error"

    ' Fields the callback sets itself are not copied over it.
    Set dPayload = UnreportedOutcome("Export", _
        "{""success"":true,""type"":""complete"",""operation_id"":""old"",""error"":""x""}", strMessage)
    TestAssert dPayload("success") = False, "an unreported outcome is never a success"
    TestAssert Not dPayload.Exists("type"), "type is left to the callback"
    TestAssert Not dPayload.Exists("operation_id"), "operation_id is left to the callback"

    ' Plain text, such as the dispatcher's refusal, is the error.
    Set dPayload = UnreportedOutcome("Export", "VCS_API_REFUSED: busy", strMessage)
    TestAssert strMessage = "VCS_API_REFUSED: busy", "plain text is the error"

    ' A JSON result without error text says the operation did not start.
    Set dPayload = UnreportedOutcome("BuildAs", "{""success"":false}", strMessage)
    TestAssert strMessage = T("The {0} operation did not start.", var0:="BuildAs"), _
        "a result with no error text gets the did-not-start message"

End Sub

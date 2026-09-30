Attribute VB_Name = "modTestScopedSync"
'---------------------------------------------------------------------------------------
' Module    : modTestScopedSync
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Regression tests for category-scoped ExportByType / ImportByType API
'           : validation and JSON result shape.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


Public Sub TestExportByTypeRejectsUnknownType()
    Dim strJson As String
    Dim dResult As Dictionary

    strJson = VCS.ExportByType("not_a_real_type")
    Set dResult = ParseJson(strJson)
    TestAssert Not CBool(dResult("success")), "unknown type fails"
    TestAssert InStr(CStr(dResult("error")), "Unknown object type") > 0, "error mentions unknown"
End Sub


Public Sub TestImportByTypeRejectsTableData()
    Dim strJson As String
    Dim dResult As Dictionary

    strJson = VCS.ImportByType("table_data")
    Set dResult = ParseJson(strJson)
    TestAssert Not CBool(dResult("success")), "table_data import rejected"
    TestAssert InStr(CStr(dResult("error")), "Import not supported") > 0, "error mentions unsupported"
End Sub


Public Sub TestImportByTypeRejectsUnknownType()
    Dim strJson As String
    Dim dResult As Dictionary

    strJson = VCS.ImportByType(Null)
    Set dResult = ParseJson(strJson)
    TestAssert Not CBool(dResult("success")), "Null type fails"
End Sub


Public Sub TestImportByTypeOverflowLeavesRunningRootAlone()
    CheckOverflowBeforeBegin True
End Sub


Public Sub TestExportByTypeOverflowLeavesRunningRootAlone()
    CheckOverflowBeforeBegin False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckOverflowBeforeBegin
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : A type value that overflows raises in type resolution, before the call
'           : begins its own operation. The call returns that error, and the root that
'           : is running this test, if the session operation holds it, is not its to
'           : finish. Holding a root on a free session operation needs the scripted
'           : pass outside the runner (A20).
'---------------------------------------------------------------------------------------
'
Private Sub CheckOverflowBeforeBegin(blnImport As Boolean)

    Dim strJson As String
    Dim dResult As Dictionary
    Dim strToken As String
    Dim intStatus As eOperationState
    Dim dCompletion As Dictionary
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy
    Dim blnLogActive As Boolean

    strToken = Operation.CurrentRootToken
    intStatus = Operation.Status
    Set dCompletion = Operation.LastCompletion
    intMode = Operation.InteractionMode
    intPolicy = Operation.DecisionPolicy
    blnLogActive = Log.Active

    If blnImport Then
        strJson = VCS.ImportByType(1E+20)
    Else
        strJson = VCS.ExportByType(1E+20)
    End If

    Set dResult = ParseJson(strJson)
    TestAssert Not CBool(dResult("success")), "an overflowing type fails"
    TestAssert CStr(dResult("error")) = OverflowDescription, "the result carries the original error"
    TestAssert CLng(dResult("errorNumber")) = 6, "the result carries the original error number"
    TestAssert Operation.CurrentRootToken = strToken, "the running root is untouched"
    TestAssert Operation.Status = intStatus, "the running root is not finished"
    TestAssert Operation.LastCompletion Is dCompletion, "no completion is emitted"
    TestAssert Operation.InteractionMode = intMode, "the interaction mode is unchanged"
    TestAssert Operation.DecisionPolicy = intPolicy, "the decision policy is unchanged"
    TestAssert Log.Active = blnLogActive, "the running log is left active"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScopedSyncOwnershipCheck
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : The A20 checks that need a free session operation, so they cannot run
'           : inside the runner's root. Run it on the development copy:
'           : vcs_run_vba(<dev copy>, "MCP_TempFunction = ScopedSyncOwnershipCheck()").
'           : Returns "OK", or the failed checks one per line. A Function, so the
'           : runner never lists it as a test.
'---------------------------------------------------------------------------------------
'
Public Function ScopedSyncOwnershipCheck() As String

    Dim cRoot As clsRootOperationLease
    Dim cVcs As clsVersionControl
    Dim dResult As Dictionary
    Dim strToken As String
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy
    Dim strFailed As String

    If Operation.Status = eosRunning Then
        ScopedSyncOwnershipCheck = "Not run: an operation is already running."
        Exit Function
    End If

    ' An error before the call's own Begin, while another caller holds the root.
    Operation.ForceUnattended = True
    Set cRoot = Operation.TryBeginRoot(eotOther)
    If cRoot Is Nothing Then
        ScopedSyncOwnershipCheck = "Not run: the holding root was refused."
        Exit Function
    End If
    strToken = Operation.CurrentRootToken

    Set dResult = ParseJson(VCS.ImportByType(1E+20))
    CheckOverflowResult dResult, "ImportByType", strFailed
    Set dResult = ParseJson(VCS.ExportByType(1E+20))
    CheckOverflowResult dResult, "ExportByType", strFailed
    Check strFailed, Operation.Status = eosRunning, "the held root is still running"
    Check strFailed, Operation.CurrentRootToken = strToken, "the held root keeps its token"
    cRoot.Complete eorSuccess
    Set cRoot = Nothing

    ' An error after the call's own Begin. Export refuses the running database, so
    ' only import can begin here; both share the error handler.
    intMode = Operation.InteractionMode
    intPolicy = Operation.DecisionPolicy
    Operation.ForceUnattended = True
    Set cVcs = New clsVersionControl
    cVcs.FaultAfterBegin = 5
    Set dResult = ParseJson(cVcs.ImportByType("queries"))
    Check strFailed, Not CBool(dResult("success")), "the owned error fails the call"
    Check strFailed, CLng(dResult("errorNumber")) = 5, "the owned error keeps its number"
    Check strFailed, CStr(dResult("error")) = ErrorDescription(5), "the owned error keeps its description"
    Check strFailed, Len(CStr(dResult("logPath"))) > 0, "the owned error names the saved log"
    Check strFailed, Operation.Status <> eosRunning, "the owned root is finished"
    Check strFailed, Operation.Result = eorFailed, "the owned root finishes as failed"
    Check strFailed, Operation.InteractionMode = intMode, "the interaction mode is restored"
    Check strFailed, Operation.DecisionPolicy = intPolicy, "the decision policy is restored"
    Check strFailed, Operation.Begin(eotOther), "the next call can begin"
    Operation.Finish eorSuccess

    If Len(strFailed) = 0 Then strFailed = "OK"
    ScopedSyncOwnershipCheck = strFailed

End Function


Private Sub CheckOverflowResult(dResult As Dictionary, strCall As String, ByRef strFailed As String)
    Check strFailed, Not CBool(dResult("success")), strCall & " fails on an overflowing type"
    Check strFailed, CStr(dResult("error")) = OverflowDescription, strCall & " keeps the Overflow description"
    Check strFailed, CLng(dResult("errorNumber")) = 6, strCall & " keeps error number 6"
End Sub


Private Sub Check(ByRef strFailed As String, blnPassed As Boolean, strCheck As String)
    If Not blnPassed Then strFailed = strFailed & "FAIL: " & strCheck & vbCrLf
End Sub


' Err.Description for Overflow in this Access language.
Private Function OverflowDescription() As String
    Dim lngValue As Long
    On Error Resume Next
    lngValue = CLng(1E+20)
    OverflowDescription = Err.Description
    Err.Clear
End Function


' Err.Description for a raised error number in this Access language.
Private Function ErrorDescription(lngNumber As Long) As String
    On Error Resume Next
    Err.Raise lngNumber
    ErrorDescription = Err.Description
    Err.Clear
End Function

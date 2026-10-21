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
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    CheckOverflowBeforeBegin True

CleanUp:
    On Error Resume Next
    If lngTestErr <> 0 Then TestAssert False, "TestImportByTypeOverflowLeavesRunningRootAlone " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestExportByTypeOverflowLeavesRunningRootAlone()
    On Error GoTo ErrHandler
    Dim lngTestErr As Long
    Dim strTestErr As String
    CheckOverflowBeforeBegin False

CleanUp:
    On Error Resume Next
    If lngTestErr <> 0 Then TestAssert False, "TestExportByTypeOverflowLeavesRunningRootAlone " & lngTestErr & ": " & strTestErr
    Exit Sub
ErrHandler:
    lngTestErr = Err.Number
    strTestErr = Err.Description
    Resume CleanUp
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckOverflowBeforeBegin
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : A type value that overflows raises in type resolution, before the call
'           : begins its own operation. The call returns that error, and the root that
'           : is running this test, if the session operation holds it, is not its to
'           : finish. Holding a root on a free session operation needs the scripted
'           : pass outside the runner ().
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
    If intStatus = eosRunning Or intStatus = eosStaged Then
        TestAssert dResult("error_pattern") = ERR_OPERATION_ALREADY_RUNNING, "busy admission refuses before preparation"
    Else
        TestAssert CStr(dResult("error")) = OverflowDescription, "the result carries the original error"
        TestAssert CLng(dResult("errorNumber")) = 6, "the result carries the original error number"
    End If
    TestAssert Not dResult.Exists("cancelled"), "an error is not a cancel"
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
' Purpose   : The checks that need a free session operation, so they cannot run
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

    ' The session policy the MCP sets around each call. It acknowledges the logged
    ' error's message box, and an owned error has to leave it in force.
    intMode = Operation.InteractionMode
    intPolicy = Operation.DecisionPolicy

    ' An error before the call's own Begin, while another caller holds the root.
    Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
    If cRoot Is Nothing Then
        VCS.ClearOperationPolicy
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
    Set cVcs = New clsVersionControl
    cVcs.FaultAfterBegin = 5
    Set dResult = ParseJson(cVcs.ImportByType("queries", False, "block"))
    Check strFailed, Not CBool(dResult("success")), "the owned error fails the call"
    Check strFailed, CLng(dResult("errorNumber")) = 5, "the owned error keeps its number"
    Check strFailed, CStr(dResult("error")) = ErrorDescription(5), "the owned error keeps its description"
    Check strFailed, Len(CStr(dResult("logPath"))) > 0, "the owned error names the saved log"
    Check strFailed, Operation.Status <> eosRunning, "the owned root is finished"
    Check strFailed, Operation.Result = eorFailed, "the owned root finishes as failed"
    Check strFailed, Operation.InteractionMode = intMode, "the session interaction mode is back"
    Check strFailed, Operation.DecisionPolicy = intPolicy, "the session policy is back in force"
    Check strFailed, Operation.Begin(eotOther), "the next call can begin"
    Operation.Finish eorSuccess
    VCS.ClearOperationPolicy

    If Len(strFailed) = 0 Then strFailed = "OK"
    ScopedSyncOwnershipCheck = strFailed

End Function


Private Sub CheckOverflowResult(dResult As Dictionary, strCall As String, ByRef strFailed As String)
    Check strFailed, Not CBool(dResult("success")), strCall & " fails on an overflowing type"
    Check strFailed, CStr(dResult("error_pattern")) = ERR_OPERATION_ALREADY_RUNNING, strCall & " refuses before preparing another operation"
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


' run outside the test runner, which owns a root. For export, load this
' development project as a library in a disposable database (export preflight
' deliberately refuses to export the running code project).
Public Function ScopedSyncCancelCheck(Optional blnExport As Boolean = False) As String
    Dim cVcs As clsVersionControl
    Dim dResult As Dictionary
    Dim strFailed As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    If Operation.Status = eosRunning Then
        ScopedSyncCancelCheck = "Not run: an operation is already running."
        Exit Function
    End If
    Set cVcs = New clsVersionControl
    cVcs.CancelAfterBegin = True
    If blnExport Then
        Set dResult = ParseJson(cVcs.ExportByType("queries", False, "block"))
    Else
        Set dResult = ParseJson(cVcs.ImportByType("queries", False, "block"))
    End If
    Check strFailed, Not CBool(dResult("success")), "cancel fails the call"
    Check strFailed, dResult.Exists("cancelled"), "cancel carries cancelled"
    If dResult.Exists("cancelled") Then
        Check strFailed, CBool(dResult("cancelled")), "cancelled is true"
    End If
    Check strFailed, CStr(dResult("error")) = "Operation was canceled.", "cancel keeps its error"
    Check strFailed, Len(CStr(dResult("logPath"))) > 0, "cancel names its log"
    Check strFailed, Operation.Result = eorCanceled, "root finishes as cancelled"
    Check strFailed, Operation.Status <> eosRunning, "cancel releases the root"

    ' The same public entry point's owned failure must remain an ordinary error.
    Set cVcs = New clsVersionControl
    cVcs.FaultAfterBegin = 5
    If blnExport Then
        Set dResult = ParseJson(cVcs.ExportByType("queries", False, "block"))
    Else
        Set dResult = ParseJson(cVcs.ImportByType("queries", False, "block"))
    End If
    Check strFailed, Not CBool(dResult("success")), "error fails the call"
    Check strFailed, Not dResult.Exists("cancelled"), "error has no cancelled field"
    Check strFailed, CLng(dResult("errorNumber")) = 5, "error keeps its number"
    Check strFailed, Operation.Result = eorFailed, "error finishes as failed"

CleanUp:
    On Error Resume Next
    VCS.ClearOperationPolicy
    If lngErr <> 0 Then strFailed = strFailed & "ERROR: " & lngErr & ": " & strErr
    If Len(strFailed) = 0 Then strFailed = "OK"
    ScopedSyncCancelCheck = strFailed
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function





' One category call on a fresh public instance, returning the parsed result.



' State every outcome of the call must leave: the root finished as expected, its owned
' state is released, the session's mode and policy are back, and the form was left open.


Public Function ScopedSyncNativeCloseCheck(Optional blnExport As Boolean = False) As String

    Const PROBE_FORM As String = "frmA36Probe"
    Const ORIGINAL_CAPTION As String = "A36 original"
    Const SOURCE_CAPTION As String = "A36 source"

    Dim cVcs As clsVersionControl
    Dim dResult As Dictionary
    Dim frmNew As Form
    Dim strTemp As String
    Dim strSource As String
    Dim strText As String
    Dim strFailed As String
    Dim lngErr As Long
    Dim strErr As String
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy
    Dim blnPolicySet As Boolean

    On Error GoTo ErrHandler

    If Operation.Status = eosRunning Then
        ScopedSyncNativeCloseCheck = "Not run: an operation is already running."
        Exit Function
    End If
    If StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0 Then
        ScopedSyncNativeCloseCheck = "Not run: host this project as a library in a disposable database."
        Exit Function
    End If

    ' Probe form, saved closed.
    Set frmNew = Application.CreateForm
    strTemp = frmNew.Name
    frmNew.Caption = ORIGINAL_CAPTION
    DoCmd.Close acForm, strTemp, acSaveYes
    Set frmNew = Nothing
    DoCmd.Rename PROBE_FORM, acForm, strTemp

    intMode = Operation.InteractionMode
    intPolicy = Operation.DecisionPolicy

    ' Baseline export gives the form a source file to import from, or to lose.
    Set cVcs = New clsVersionControl
    Set dResult = ParseJson(cVcs.ExportByType("forms", True, "block"))
    Check strFailed, CBool(dResult("success")), "the baseline export succeeds"
    ' A form without a code module may export as .bas instead of .form.
    strSource = Options.GetExportFolder & "forms" & PathSep & PROBE_FORM & ".form"
    If Not FSO.FileExists(strSource) Then strSource = Replace(strSource, ".form", ".bas")
    If Not FSO.FileExists(strSource) Then
        ScopedSyncNativeCloseCheck = strFailed & "Not run: the baseline export wrote no " & strSource
        GoTo CleanUp
    End If
    If blnExport Then
        FSO.DeleteFile strSource, True
    Else
        strText = ReadFile(strSource)
        If InStr(1, strText, ORIGINAL_CAPTION, vbBinaryCompare) = 0 Then
            ScopedSyncNativeCloseCheck = strFailed & "Not run: the source has no caption to change."
            GoTo CleanUp
        End If
        WriteFile Replace(strText, ORIGINAL_CAPTION, SOURCE_CAPTION), strSource
    End If

    ' 1. Confirmed native Cancel with the form still open.
    DoCmd.OpenForm PROBE_FORM, acNormal, , , , acHidden
    SetCloseFaultForTest 2501
    Set dResult = RunNativeCloseCall(blnExport, False)
    SetCloseFaultForTest 0
    Check strFailed, Not CBool(dResult("success")), "native cancel fails the call"
    Check strFailed, dResult.Exists("cancelled"), "native cancel carries cancelled"
    If dResult.Exists("cancelled") Then Check strFailed, CBool(dResult("cancelled")), "cancelled is true"
    Check strFailed, CStr(dResult("error")) = "Operation was canceled.", "native cancel keeps the cancellation error"
    Check strFailed, Len(CStr(dResult("logPath"))) > 0, "native cancel names its log"
    Check strFailed, Not dResult.Exists("decision_required"), "native cancel is not decision_required"
    CheckNativeCloseState strFailed, "native cancel", eorCanceled, PROBE_FORM, intMode, intPolicy
    If blnExport Then
        Check strFailed, Not FSO.FileExists(strSource), "the open form is not exported after its close was canceled"
    Else
        Check strFailed, Forms(PROBE_FORM).Caption = ORIGINAL_CAPTION, "the open form is not imported after its close was canceled"
    End If

    ' 2. A close that fails for another reason and leaves the form open is ordinary.
    SetCloseFaultForTest 2467
    Set dResult = RunNativeCloseCall(blnExport, False)
    SetCloseFaultForTest 0
    Check strFailed, Not CBool(dResult("success")), "an ordinary close failure fails the call"
    Check strFailed, Not dResult.Exists("cancelled"), "an ordinary close failure has no cancelled field"
    Check strFailed, CStr(dResult("error")) <> "Operation was canceled.", "an ordinary close failure is not the cancellation error"
    Check strFailed, Len(CStr(dResult("logPath"))) > 0, "an ordinary close failure names its log"
    CheckNativeCloseState strFailed, "ordinary failure", eorFailed, PROBE_FORM, intMode, intPolicy
    If blnExport Then
        Check strFailed, Not FSO.FileExists(strSource), "an ordinary close failure does not export the open form"
    Else
        Check strFailed, Forms(PROBE_FORM).Caption = ORIGINAL_CAPTION, "an ordinary close failure does not import into the open form"
    End If

    ' 3. A blocked decision stays decision_required, even with a native cancel recorded.
    SetCloseFaultForTest 2501
    Set dResult = RunNativeCloseCall(blnExport, True)
    SetCloseFaultForTest 0
    Check strFailed, Not CBool(dResult("success")), "a blocked decision fails the call"
    Check strFailed, dResult.Exists("decision_required"), "a blocked decision carries decision_required"
    If dResult.Exists("decision_required") Then Check strFailed, CBool(dResult("decision_required")), "decision_required is true"
    Check strFailed, CStr(dResult("error_pattern")) = "decision_required", "a blocked decision keeps its error pattern"
    Check strFailed, Not dResult.Exists("cancelled"), "a blocked decision is not a cancellation"
    CheckNativeCloseState strFailed, "blocked decision", eorDecisionRequired, PROBE_FORM, intMode, intPolicy

    ' 4. Control: with the close working, the same call does its work.
    Set dResult = RunNativeCloseCall(blnExport, False)
    Check strFailed, CBool(dResult("success")), "the call succeeds once the close works"
    Check strFailed, Not dResult.Exists("cancelled"), "a success has no cancelled field"
    If blnExport Then
        Check strFailed, FSO.FileExists(strSource), "the form is exported once the close works"
    Else
        Check strFailed, CurrentProject.AllForms(PROBE_FORM).Name = PROBE_FORM, "the form survives the import"
        DoCmd.OpenForm PROBE_FORM, acNormal, , , , acHidden
        Check strFailed, Forms(PROBE_FORM).Caption = SOURCE_CAPTION, "the form is imported once the close works"
    End If

CleanUp:
    On Error Resume Next
    SetCloseFaultForTest 0
    If IsLoaded(acForm, PROBE_FORM) Then DoCmd.Close acForm, PROBE_FORM, acSaveNo
    DoCmd.DeleteObject acForm, PROBE_FORM
    DoCmd.DeleteObject acForm, strTemp
    Err.Clear
    If lngErr <> 0 Then strFailed = strFailed & "ERROR: " & lngErr & ": " & strErr & vbCrLf
    If Len(ScopedSyncNativeCloseCheck) = 0 Then
        If Len(strFailed) = 0 Then strFailed = "OK"
        ScopedSyncNativeCloseCheck = strFailed
    End If
    Exit Function

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function


Private Function RunNativeCloseCall(blnExport As Boolean, blnBlocked As Boolean) As Dictionary
    Dim cVcs As clsVersionControl
    Set cVcs = New clsVersionControl
    cVcs.BlockAfterBegin = blnBlocked
    If blnExport Then
        Set RunNativeCloseCall = ParseJson(cVcs.ExportByType("forms", True, "block"))
    Else
        Set RunNativeCloseCall = ParseJson(cVcs.ImportByType("forms", True, "block"))
    End If
End Function


Private Sub CheckNativeCloseState(ByRef strFailed As String, strCase As String, _
    eExpected As eOperationResult, strForm As String, intMode As eInteractionMode, _
    intPolicy As eDecisionPolicy)

    Check strFailed, Operation.Result = eExpected, strCase & " finishes the root as expected"
    Check strFailed, Operation.Status <> eosRunning, strCase & " releases the root"
    Check strFailed, Not Log.Active, strCase & " releases the log"
    Check strFailed, Operation.InteractionMode = intMode, strCase & " restores the interaction mode"
    Check strFailed, Operation.DecisionPolicy = intPolicy, strCase & " restores the caller policy"
    Check strFailed, IsLoaded(acForm, strForm), strCase & " leaves the affected form open"

End Sub


' Run outside the suite root, with this project loaded as a library in a disposable
' user database. The caller seeds frmNativeCloseProbe and its source/index first,
' then compares source/index bytes and nanosecond timestamps around this call.
Public Function IndexedSingleObjectCloseCheck(ByVal strSource As String, _
    ByVal blnExport As Boolean, ByVal blnNativeCancel As Boolean) As String
    Const PROBE_FORM As String = "frmNativeCloseProbe"
    Dim cComponent As IDbComponent
    Dim dResult As Dictionary
    Dim strLog As String
    Dim blnCanceled As Boolean
    Dim intMode As eInteractionMode
    Dim blnForce As Boolean
    Dim lngError As Long
    Dim strError As String
    On Error GoTo ErrHandler
    If Operation.IsActive Then
        IndexedSingleObjectCloseCheck = "Not run: an operation is already active."
        Exit Function
    End If
    If StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0 Then
        IndexedSingleObjectCloseCheck = "Not run: use a disposable user database."
        Exit Function
    End If
    intMode = Operation.InteractionMode
    blnForce = Operation.ForceUnattended
    Operation.ForceUnattended = True
    Operation.InteractionMode = eimSilent
    If Not Operation.Begin(IIf(blnExport, eotExport, eotMerge)) Then
        IndexedSingleObjectCloseCheck = "Not run: operation admission refused."
        GoTo CleanUp
    End If
    If blnNativeCancel Then SetCloseOutcomeForTest ecoCanceled Else SetCloseOutcomeForTest ecoFailed
    If blnExport Then
        modExport.ExportSingleObject CurrentProject.AllForms(PROBE_FORM), , False, dResult
        IndexedSingleObjectCloseCheck = FinishSingleObjectExport(Operation, Log.SavedLogFilePath, dResult)
    Else
        Set cComponent = GetComponentClass(edbForm)
        LoadSingleObject cComponent, PROBE_FORM, strSource, False, strLog, dResult, blnCanceled
        IndexedSingleObjectCloseCheck = FinishSingleObjectImport(Operation, strLog, Log.GetErrorJournalSince(0), dResult, blnCanceled)
    End If
CleanUp:
    On Error Resume Next
    SetCloseOutcomeForTest ecoReal
    If Operation.IsActive Then Operation.Finish eorFailed
    Operation.InteractionMode = intMode
    Operation.ForceUnattended = blnForce
    If lngError <> 0 Then IndexedSingleObjectCloseCheck = "ERROR: " & lngError & ": " & strError
    Exit Function
ErrHandler:
    lngError = Err.Number
    strError = Err.Description
    Resume CleanUp
End Function


' Real native-prompt driver entry point. No close fault or cancellation seam is
' armed here. The external supervisor captures the prompt, then an independent
' Win32 phase verifies and chooses its literal Cancel control.
Public Function NativeCloseLiveCall(ByVal strCase As String) As String
    If Operation.IsActive Then
        NativeCloseLiveCall = "Not run: an operation is already active."
        Exit Function
    End If
    If StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0 Then
        NativeCloseLiveCall = "Not run: use a disposable user database."
        Exit Function
    End If
    Operation.InteractionMode = eimNormal
    Operation.ForceUnattended = False
    Select Case strCase
        Case "category-export", "category-import"
            Operation.Source = eosUserInterface
            If strCase = "category-export" Then
                NativeCloseLiveCall = VCS.ExportByType("forms", True)
            Else
                NativeCloseLiveCall = VCS.ImportByType("forms", True)
            End If
        Case "single-export", "single-import"
            Operation.Source = eosExternalAPI
            If strCase = "single-export" Then
                NativeCloseLiveCall = VCS.ExportObject("form", "frmNativeCloseProbe")
            Else
                NativeCloseLiveCall = VCS.ImportObject("form", "frmNativeCloseProbe")
            End If
        Case Else
            NativeCloseLiveCall = "Not run: unknown native-close case."
    End Select
End Function

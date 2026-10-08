Attribute VB_Name = "modTestBuildAsPaths"
'---------------------------------------------------------------------------------------
' Module    : modTestBuildAsPaths
' Author    : Josh
' Date      : 10/1/2026
' Purpose   : BuildAs with a source folder and an output file (X09). An invalid pair is
'           : refused as JSON before anything starts, with no picker. The completion
'           : callback names the file a build wrote, and APICapabilities advertises
'           : the feature so a caller can refuse an older add-in.
'           :
'           : Only refusals run here: a real build would close the database the tests
'           : run in. A picker would stall the run, so a run that finishes shows none
'           : opened. BuildAsWithoutPathsCheck covers the no-argument (ribbon) form.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("unit")

' Paths that are never created.
Private Const NO_SOURCE_FOLDER As String = "C:\vcs-no-such-source-folder\"
Private Const NO_OUTPUT_FOLDER As String = "C:\vcs-no-such-output-folder\"


' The source folder of the project the tests run in. It has vcs-options.json.
Private Function ThisSourceFolder() As String
    ThisSourceFolder = CurrentProject.FullName & ".src" & PathSep
End Function


Private Sub AssertPathRefusal(ByVal strJson As String, ByVal strCase As String)

    Dim dResult As Dictionary

    Set dResult = ParseJson(strJson)
    TestAssert Not dResult Is Nothing, strCase & ": the refusal is JSON"
    If dResult Is Nothing Then Exit Sub
    TestAssert Not CBool(dResult("success")), strCase & ": success is false"
    TestAssert dNZ(dResult, "error_pattern") = ERR_INVALID_BUILD_PATH, _
        strCase & ": the pattern is invalid_build_path"
    TestAssert Len(dNZ(dResult, "error")) > 0, strCase & ": the refusal says why"

End Sub


Public Sub TestBuildAsRefusesHalfAPair()

    Dim intBefore As eOperationState

    intBefore = Operation.Status
    AssertPathRefusal VCS.BuildAs(ThisSourceFolder), "a source with no output"
    AssertPathRefusal VCS.BuildAs(, NO_OUTPUT_FOLDER & "Out.accdb"), "an output with no source"
    TestAssert Operation.Status = intBefore, "a refusal does not start or end an operation"

End Sub


Public Sub TestBuildAsRefusesSourceWithoutVcsOptions()
    AssertPathRefusal VCS.BuildAs(NO_SOURCE_FOLDER, NO_OUTPUT_FOLDER & "Out.accdb"), _
        "a folder with no source files"
End Sub


Public Sub TestBuildAsRefusesUnusableOutput()

    Dim intBefore As eOperationState

    If Not FolderHasVcsOptionsFile(ThisSourceFolder) Then
        TestAssert False, "the test project has source files in " & ThisSourceFolder
        Exit Sub
    End If

    intBefore = Operation.Status
    AssertPathRefusal VCS.BuildAs(ThisSourceFolder, "Out.accdb"), "a bare file name"
    AssertPathRefusal VCS.BuildAs(ThisSourceFolder, NO_OUTPUT_FOLDER), "a folder, not a file"
    AssertPathRefusal VCS.BuildAs(ThisSourceFolder, NO_OUTPUT_FOLDER & "Out.accdb"), _
        "an output folder that does not exist"
    AssertPathRefusal VCS.BuildAs(ThisSourceFolder, CodeProject.FullName), "the add-in itself"
    TestAssert Operation.Status = intBefore, "no refusal starts or ends an operation"

End Sub


Public Sub TestBuildAsRefusesRelativeDestinations()
    Dim varPath As Variant
    Dim dBefore As Dictionary
    Dim dCompletion As Dictionary
    Dim lngError As Long
    Dim strError As String

    On Error GoTo ErrHandler
    For Each varPath In Array(".\Out.accdb", "..\Out.accdb", _
        Left$(CurrentProject.FullName, 2) & ".\Out.accdb", Mid$(CurrentProject.FullName, 3))
        Set dBefore = BuildAsAdmissionState()
        Set dCompletion = Operation.LastCompletion
        AssertPathRefusal VCS.BuildAs(ThisSourceFolder, CStr(varPath)), CStr(varPath)
        TestAssert ConvertToJson(dBefore) = ConvertToJson(BuildAsAdmissionState()), _
            "refusal preserves database, forms, operation and session: " & CStr(varPath)
        TestAssert Operation.LastCompletion Is dCompletion, "refusal emits no completion"
    Next varPath

CleanUp:
    On Error Resume Next
    If lngError <> 0 Then TestAssert False, CStr(lngError) & ": " & strError
    Exit Sub
ErrHandler:
    lngError = Err.Number
    strError = Err.Description
    Resume CleanUp
End Sub


Public Sub TestCompletionCarriesRecordedOutputPath()

    Dim cOp As clsOperation
    Const OUTPUT_PATH As String = "C:\builds\Out.accdb"

    ' A private instance, so the session operation is never touched.
    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    TestAssert cOp.Begin(eotBuild), "the build root is granted"
    cOp.RecordOutputPath OUTPUT_PATH
    cOp.Finish eorSuccess
    TestAssert dNZ(cOp.LastCompletion, "output_path") = OUTPUT_PATH, _
        "the completion names the file the build wrote"

    ' A later record is too late, and the next root starts without the path.
    cOp.RecordOutputPath "C:\builds\Late.accdb"
    TestAssert dNZ(cOp.LastCompletion, "output_path") = OUTPUT_PATH, _
        "a record after Finish changes nothing"
    cOp.ForceUnattended = True
    TestAssert cOp.Begin(eotBuild), "a second root is granted"
    cOp.Finish eorSuccess
    TestAssert Not cOp.LastCompletion.Exists("output_path"), _
        "a root that recorded no file names none"

End Sub


Public Sub TestCapabilitiesNameBuildAsPaths()

    Dim dResult As Dictionary
    Dim varName As Variant
    Dim blnFound As Boolean

    ' APICapabilities is the probe entry point; the API method gives the same reply.
    TestAssert VCS.GetCapabilities = APICapabilities, "GetCapabilities matches APICapabilities"
    Set dResult = ParseJson(APICapabilities)
    TestAssert Not dResult Is Nothing, "the capability reply is JSON"
    If dResult Is Nothing Then Exit Sub
    TestAssert CBool(dResult("success")), "the capability query succeeds"
    For Each varName In dResult("capabilities")
        If varName = "build_as_paths" Then blnFound = True
    Next varName
    TestAssert blnFound, "build_as_paths is advertised"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildAsWithoutPathsCheck
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : BuildAs with no arguments keeps the ribbon path: no JSON refusal, and it
'           : starts through Begin. This holds a root of its own, so that Begin is
'           : refused before any picker opens; the noninteractive scope keeps the
'           : refusal a log entry. Needs a free session operation, so it cannot run
'           : inside the runner's root. Run it on the development copy:
'           : vcs_run_vba(<dev copy>, "MCP_TempFunction = BuildAsWithoutPathsCheck()").
'           : Returns "OK", or the failed checks one per line. A Function, so the
'           : runner never lists it as a test.
'---------------------------------------------------------------------------------------
'
Public Function BuildAsWithoutPathsCheck() As String

    Dim cRoot As clsRootOperationLease
    Dim strToken As String
    Dim lngScope As Long
    Dim strResult As String
    Dim strFailed As String

    If Operation.Status = eosRunning Then
        BuildAsWithoutPathsCheck = "Not run: an operation is already running."
        Exit Function
    End If

    Operation.ForceUnattended = True
    Set cRoot = Operation.TryBeginRoot(eotOther)
    If cRoot Is Nothing Then
        BuildAsWithoutPathsCheck = "Not run: the holding root was refused."
        Exit Function
    End If
    strToken = Operation.CurrentRootToken

    lngScope = Operation.PushInteractionScope(eimNonInteractive, edpDecline, False)
    strResult = VCS.BuildAs
    Operation.CloseInteractionScope lngScope

    Check strFailed, Len(strResult) = 0, "the no-argument form returns no JSON refusal: " & strResult
    Check strFailed, Operation.Status = eosRunning, "the held root is still running"
    Check strFailed, Operation.CurrentRootToken = strToken, "the held root keeps its token"
    cRoot.Complete eorSuccess
    Set cRoot = Nothing

    If Len(strFailed) = 0 Then strFailed = "OK"
    BuildAsWithoutPathsCheck = strFailed

End Function


Private Sub Check(ByRef strFailed As String, blnPassed As Boolean, strCheck As String)
    If Not blnPassed Then strFailed = strFailed & "FAIL: " & strCheck & vbCrLf
End Sub


' A38 scripted public-boundary probe. Run outside the test runner in a disposable
' database with this development project loaded as a library. A held root safely
' distinguishes path admission from operation admission on the unfixed version.
Public Function BuildAsAdmissionCheck(ByVal strSource As String, ByVal strOutput As String, _
    ByVal strExpectedPattern As String, ByVal blnHoldRoot As Boolean) As String

    Dim cRoot As clsRootOperationLease
    Dim dResult As New Dictionary
    Dim dBefore As Dictionary
    Dim dAfter As Dictionary
    Dim dRefusal As Dictionary
    Dim dCompletion As Dictionary
    Dim strReply As String
    Dim lngScope As Long
    Dim intMode As eInteractionMode
    Dim intPolicy As eDecisionPolicy
    Dim lngError As Long
    Dim strError As String

    On Error GoTo ErrHandler
    intMode = Operation.InteractionMode
    intPolicy = Operation.DecisionPolicy
    If Operation.IsActive Then
        dResult.Add "error", "Probe requires a free disposable session."
        GoTo CleanUp
    End If
    lngScope = Operation.PushInteractionScope(eimNonInteractive, edpBlock, False)
    If blnHoldRoot Then
        Set cRoot = Operation.TryBeginRoot(eotOther)
        If cRoot Is Nothing Then
            dResult.Add "error", "Probe root was refused."
            GoTo CleanUp
        End If
    End If

    Set dBefore = BuildAsAdmissionState()
    Set dCompletion = Operation.LastCompletion
    strReply = VCS.BuildAs(strSource, strOutput)
    Set dRefusal = ParseJson(strReply)
    Set dAfter = BuildAsAdmissionState()
    dResult.Add "before", dBefore
    dResult.Add "after", dAfter
    dResult.Add "refusal", dRefusal
    dResult.Add "state_unchanged", ConvertToJson(dBefore) = ConvertToJson(dAfter)
    dResult.Add "completion_unchanged", Operation.LastCompletion Is dCompletion
    dResult.Add "success", Not CBool(dRefusal("success")) And _
        dNZ(dRefusal, "error_pattern") = strExpectedPattern And _
        CBool(dResult("state_unchanged")) And CBool(dResult("completion_unchanged"))

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorSuccess
    Set cRoot = Nothing
    ' Root completion must not consume the caller-owned policy scope.
    dResult.Add "scope_survives_root", Operation.InteractionMode = eimNonInteractive And _
        Operation.DecisionPolicy = edpBlock
    Operation.CloseInteractionScope lngScope
    dResult.Add "scope_restored", Operation.InteractionMode = intMode And _
        Operation.DecisionPolicy = intPolicy
    If lngError <> 0 Then dResult.Add "error", CStr(lngError) & ": " & strError
    BuildAsAdmissionCheck = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    lngError = Err.Number
    strError = Err.Description
    Resume CleanUp
End Function


Private Function BuildAsAdmissionState() As Dictionary
    Dim dState As New Dictionary
    Dim colForms As New Collection
    Dim frm As Access.Form
    Dim dForm As Dictionary

    dState.Add "database", CurrentProject.FullName
    dState.Add "status", Operation.Status
    dState.Add "root_token", Operation.CurrentRootToken
    dState.Add "mode", Operation.InteractionMode
    dState.Add "policy", Operation.DecisionPolicy
    dState.Add "attended", Operation.Attended
    dState.Add "force_unattended", Operation.ForceUnattended
    dState.Add "operation_type", Operation.OperationType
    dState.Add "source_name", Operation.SourceName
    dState.Add "automation_source", Operation.AutomationSource
    dState.Add "blocked", Operation.DecisionBlocked
    dState.Add "error_level", Operation.ErrorLevel
    dState.Add "cancel_requested", Operation.CancelRequested
    dState.Add "native_close_cancelled", Operation.NativeCloseCanceled
    AddDecisionJournal dState, Operation.Decisions
    dState.Add "log_active", Log.Active
    dState.Add "export_folder", Options.GetExportFolder
    dState.Add "timer_pending", TimerIsPending
    dState.Add "compatibility_envelope", CurrentCompatibilityEnvelope()
    For Each frm In Forms
        Set dForm = New Dictionary
        dForm.Add "name", frm.Name
        dForm.Add "visible", frm.Visible
        dForm.Add "view", frm.CurrentView
        dForm.Add "dirty", frm.Dirty
        colForms.Add dForm
    Next frm
    dState.Add "forms", colForms
    Set BuildAsAdmissionState = dState
End Function

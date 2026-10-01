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

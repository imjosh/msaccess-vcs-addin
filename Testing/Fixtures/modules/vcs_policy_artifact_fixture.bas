Attribute VB_Name = "vcs_policy_artifact_fixture"
'---------------------------------------------------------------------------------------
' Module    : vcs_policy_artifact_fixture
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Mutate a disposable host's native command-bar picture within Access.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : F1SetCommandBarPicture
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Assign and save an IPictureDisp locally, avoiding cross-process marshalling.
'---------------------------------------------------------------------------------------
'
Public Function F1SetCommandBarPicture(ByVal strExpectedHost As String, ByVal strMenuName As String, _
    ByVal strPictureFile As String, ByVal strProofFile As String) As Boolean

    Dim objFSO As Object
    Dim objButton As Object
    Dim objPicture As Object
    Dim lngError As Long
    Dim strError As String

    ' This fixture has no add-in logging dependency. Preserve inherited diagnostics
    ' before installing its own handler; all failures propagate to the Python driver.
    If Err.Number <> 0 Then
        Debug.Print "F1SetCommandBarPicture inherited error " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo Failed
    If StrComp(CurrentProject.FullName, strExpectedHost, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2471, "F1SetCommandBarPicture", "The disposable fixture host does not match."
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(strPictureFile) Then
        Err.Raise vbObjectError + 2472, "F1SetCommandBarPicture", "The picture fixture does not exist."
    End If
    If StrComp(Left$(strProofFile, Len(CurrentProject.Path) + 1), _
        CurrentProject.Path & "\", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2473, "F1SetCommandBarPicture", "The picture proof must stay beside the disposable host."
    End If
    If objFSO.FileExists(strProofFile) Then
        Err.Raise vbObjectError + 2474, "F1SetCommandBarPicture", "The picture proof file already exists."
    End If
    Set objPicture = stdole.LoadPicture(strPictureFile)
    Set objButton = Application.CommandBars(strMenuName).Controls(1)
    objButton.Picture = objPicture
    objButton.BuiltInFace = False
    stdole.SavePicture objButton.Picture, strProofFile
    F1SetCommandBarPicture = True
    Exit Function

Failed:
    lngError = Err.Number
    strError = Err.Description
    Err.Raise lngError, "vcs_policy_artifact_fixture.F1SetCommandBarPicture", strError
End Function

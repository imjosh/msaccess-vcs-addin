Attribute VB_Name = "vcs_policy_artifact_fixture"
'---------------------------------------------------------------------------------------
' Module    : vcs_policy_artifact_fixture
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Mutate a disposable host's native command-bar picture within Access.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private m_strLastError As String


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

    ' This fixture has no add-in logging dependency. Failures remain observable
    ' through its Boolean result and retained diagnostic, without a modal rethrow.
    m_strLastError = vbNullString
    If Err.Number <> 0 Then
        m_strLastError = "Inherited error " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo Failed
    If StrComp(CurrentProject.FullName, strExpectedHost, vbTextCompare) <> 0 Then
        m_strLastError = CStr(vbObjectError + 2471) & ": The disposable fixture host does not match."
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(strPictureFile) Then
        m_strLastError = CStr(vbObjectError + 2472) & ": The picture fixture does not exist."
        Exit Function
    End If
    If StrComp(Left$(strProofFile, Len(CurrentProject.Path) + 1), _
        CurrentProject.Path & "\", vbTextCompare) <> 0 Then
        m_strLastError = CStr(vbObjectError + 2473) & ": The picture proof must stay beside the disposable host."
        Exit Function
    End If
    If objFSO.FileExists(strProofFile) Then
        m_strLastError = CStr(vbObjectError + 2474) & ": The picture proof file already exists."
        Exit Function
    End If
    Set objPicture = stdole.LoadPicture(strPictureFile)
    Set objButton = Application.CommandBars(strMenuName).Controls(1)
    objButton.Picture = objPicture
    stdole.SavePicture objButton.Picture, strProofFile
    F1SetCommandBarPicture = True
    Exit Function

Failed:
    m_strLastError = CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    F1SetCommandBarPicture = False
End Function


'---------------------------------------------------------------------------------------
' Procedure : F1CanonicalizePicture
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Normalize a requested picture to an owned proof without changing source.
'---------------------------------------------------------------------------------------
'
Public Function F1CanonicalizePicture(ByVal strExpectedHost As String, _
    ByVal strPictureFile As String, ByVal strProofFile As String) As Boolean

    Dim objFSO As Object
    Dim objPicture As Object

    m_strLastError = vbNullString
    If Err.Number <> 0 Then
        m_strLastError = "Inherited error " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo Failed
    If StrComp(CurrentProject.FullName, strExpectedHost, vbTextCompare) <> 0 Then
        m_strLastError = CStr(vbObjectError + 2471) & ": The disposable fixture host does not match."
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(strPictureFile) Then
        m_strLastError = CStr(vbObjectError + 2472) & ": The picture fixture does not exist."
        Exit Function
    End If
    If StrComp(Left$(strProofFile, Len(CurrentProject.Path) + 1), _
        CurrentProject.Path & "\", vbTextCompare) <> 0 Then
        m_strLastError = CStr(vbObjectError + 2473) & ": The picture proof must stay beside the disposable host."
        Exit Function
    End If
    If objFSO.FileExists(strProofFile) Then
        m_strLastError = CStr(vbObjectError + 2474) & ": The picture proof file already exists."
        Exit Function
    End If
    Set objPicture = stdole.LoadPicture(strPictureFile)
    stdole.SavePicture objPicture, strProofFile
    F1CanonicalizePicture = True
    Exit Function

Failed:
    m_strLastError = CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    F1CanonicalizePicture = False
End Function


'---------------------------------------------------------------------------------------
' Procedure : F1PictureFixtureLastError
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Return the native fixture failure to its Python caller without modal UI.
'---------------------------------------------------------------------------------------
'
Public Function F1PictureFixtureLastError() As String
    F1PictureFixtureLastError = m_strLastError
End Function

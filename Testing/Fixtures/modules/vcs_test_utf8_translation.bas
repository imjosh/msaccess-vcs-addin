Attribute VB_Name = "vcs_test_utf8_translation"
Option Compare Database
Option Explicit
Option Private Module

' Imported once before running code in a fresh disposable development host.
' Retain this fixture through owned close and fresh reopen; dispose of the host.
Public Function Utf8TranslationProbe(ByVal blnExport As Boolean) As String
    Dim cTranslation As clsTranslation
    Dim blnSavedContribute As Boolean
    Dim blnSaved As Boolean
    Dim strExpected As String
    Dim lngErr As Long
    Dim strErr As String
    LogUnhandledErrors "vcs_test_utf8_translation.Utf8TranslationProbe"
    On Error GoTo ErrHandler
    Set cTranslation = New clsTranslation
    blnSavedContribute = cTranslation.Contribute
    blnSaved = True
    cTranslation.Contribute = False
    strExpected = FSO.BuildPath(CodeProject.Path, "Translation")
    If StrComp(cTranslation.TranslationsPath, strExpected, vbTextCompare) <> 0 Then
        Utf8TranslationProbe = "Translation path is outside the disposable checkout"
        GoTo CleanUp
    End If
    If blnExport Then
        cTranslation.ExportTemplate
        cTranslation.ExportTranslations
    End If
    Utf8TranslationProbe = "OK"
CleanUp:
    On Error Resume Next
    If blnSaved Then cTranslation.Contribute = blnSavedContribute
    Set cTranslation = Nothing
    If lngErr <> 0 Then Utf8TranslationProbe = "Translation caller " & lngErr & ": " & strErr
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function

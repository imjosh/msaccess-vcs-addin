VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestCallerErrorTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestCallerErrorTranslation
' Purpose   : Decode localized import results against real translation-table entries.
'           : Each check rolls back its fixture and restores the language cache. A private
'           : operation keeps errors and completion away from the active test runner.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Private Const TEST_LANGUAGE As String = "A26_TEST"
Private Const FALLBACK_MESSAGE As String = "Import completed with errors. Check the log for details."
Private Const MULTIPLE_MESSAGE As String = " (and {0} more errors in the log)"

Public Sub TestLocalizedImportWithoutJournal()
    CheckImportResult 0, True
End Sub

Public Sub TestLocalizedImportWithWarningOnly()
    CheckImportResult 0, False
End Sub

Public Sub TestSingleImportErrorIsPreserved()
    CheckImportResult 1, False
End Sub

Public Sub TestLocalizedImportCountsErrorAndCriticalEntries()
    CheckImportResult 3, False
End Sub

Private Sub CheckImportResult(ByVal lngErrors As Long, ByVal blnNoJournal As Boolean)

    Dim dbs As DAO.Database
    Dim wsp As DAO.Workspace
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim colErrors As Collection
    Dim dResult As Dictionary
    Dim strLanguage As String
    Dim strFallback As String
    Dim strMultiple As String
    Dim strFirst As String
    Dim strExpected As String
    Dim blnTransaction As Boolean
    Dim blnLanguageSaved As Boolean
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    strLanguage = Translation.GetCurrentLanguage
    blnLanguageSaved = True
    Set dbs = CodeDb
    Set wsp = DBEngine.Workspaces(0)
    wsp.BeginTrans
    blnTransaction = True
    dbs.Execute "DELETE FROM tblTranslation WHERE Language='" & TEST_LANGUAGE & "'", dbFailOnError
    dbs.Execute "DELETE FROM tblLanguages WHERE ID='" & TEST_LANGUAGE & "'", dbFailOnError
    dbs.Execute "INSERT INTO tblLanguages (ID) VALUES ('" & TEST_LANGUAGE & "')", dbFailOnError

    strFallback = "Erreur ""import"" C:\source" & vbCrLf & "Voir le journal."
    strMultiple = vbCrLf & "Encore ""{0}"" erreurs \journal"
    strFirst = "Premiere erreur ""originale"" C:\source {0} {1}" & vbCrLf & "details"
    AddTranslation dbs, FALLBACK_MESSAGE, strFallback
    AddTranslation dbs, MULTIPLE_MESSAGE, strMultiple
    Translation.SetLanguage TEST_LANGUAGE
    TestAssert T(FALLBACK_MESSAGE) = strFallback, "the fixture language is loaded"

    If Not blnNoJournal Then
        Set colErrors = New Collection
        AddEntry colErrors, "warning", "Not an error"
        If lngErrors > 0 Then AddEntry colErrors, "error", strFirst
        If lngErrors > 1 Then AddEntry colErrors, "critical", "Critical error"
        AddEntry colErrors, "warning", "Still not an error"
        If lngErrors > 2 Then AddEntry colErrors, "error", "Another error"
    End If

    Set cOp = New clsOperation
    cOp.ForceUnattended = True
    Set cRoot = cOp.TryBeginRoot(eotMerge, edpDecline)
    TestAssert Not cRoot Is Nothing, "the private import begins"
    cOp.ErrorLevel = eelError
    Set dResult = ParseJson(FinishSingleObjectImport(cOp, "C:\source\Merge.log", colErrors))
    Select Case lngErrors
        Case 0: strExpected = strFallback
        Case 1: strExpected = strFirst
        Case Else
            strExpected = strFirst & Replace(strMultiple, "{0}", CStr(lngErrors - 1))
    End Select
    TestAssert dResult("error") = strExpected, "decoded error matches the localized message"
    TestAssert Not CBool(dResult("success")), "an import with errors remains unsuccessful"
    TestAssert dResult("logPath") = "C:\source\Merge.log", "the log path is preserved"
    TestAssert dResult.Count = 3, "the response fields remain compatible"
    TestAssert cOp.Status = eosReady, "completion releases the import root"

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    If blnTransaction Then wsp.Rollback
    If blnLanguageSaved Then Translation.SetLanguage strLanguage
    If Err.Number <> 0 Then TestAssert False, "translation fixture cleanup: " & Err.Description
    If lngErr <> 0 Then TestAssert False, "localized import result: " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub

Private Sub AddTranslation(dbs As DAO.Database, ByVal strMessage As String, ByVal strTranslated As String)

    Dim rst As DAO.Recordset
    Dim lngId As Long

    Set rst = dbs.OpenRecordset("SELECT * FROM tblStrings WHERE msgid='" & strMessage & "'", dbOpenDynaset)
    If rst.EOF Then
        rst.AddNew
        rst!msgid = strMessage
        rst.Update
        rst.Bookmark = rst.LastModified
    End If
    lngId = rst!ID
    rst.Close
    Set rst = dbs.OpenRecordset("tblTranslation", dbOpenDynaset)
    rst.AddNew
    rst!Language = TEST_LANGUAGE
    rst!StringID = lngId
    rst!Translation = strTranslated
    rst.Update
    rst.Close

End Sub

Private Sub AddEntry(colErrors As Collection, ByVal strLevel As String, ByVal strMessage As String)
    Dim dEntry As Dictionary
    Set dEntry = New Dictionary
    dEntry.Add "level", strLevel
    dEntry.Add "message", strMessage
    colErrors.Add dEntry
End Sub

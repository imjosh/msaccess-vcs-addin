Attribute VB_Name = "modTestQuery"
'---------------------------------------------------------------------------------------
' Module    : modTestQuery
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Legacy query import tests. Paired .sql overlays QueryDefs.SQL only when
'           : ForceImportOriginalQuerySQL is enabled (issue #769).
'           : Run: ?VCS.RunTests("modTestQuery")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("integration")


Private Const TEST_QUERY As String = "vcs_test_qry_sql_override"
Private Const TEST_ALT_QUERY As String = "vcs_test_qry_alt_export"


'---------------------------------------------------------------------------------------
' Procedure : TestLegacyQueryForceSqlOverride
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : A divergent .bas + .sql pair imports the .bas definition unless Force
'           : original SQL is on, in which case the .sql definition wins. Also covers
'           : the Load Selected missing-source contract: format 4.1.2 resolves to .bas,
'           : so a leftover .qdef path is the one that used to fail silently.
'---------------------------------------------------------------------------------------
'
Public Sub TestLegacyQueryForceSqlOverride()

    Dim cQuery As IDbComponent
    Dim strFolder As String
    Dim strBas As String
    Dim strSql As String
    Dim blnSavedForce As Boolean
    Dim blnIndexDisabled As Boolean
    Dim strImported As String
    Dim eelSavedLevel As eErrorLevel
    Dim lngErr As Long
    Dim strErr As String

    blnSavedForce = Options.ForceImportOriginalQuerySQL
    blnIndexDisabled = VCSIndex.Disabled
    VCSIndex.Disabled = True
    ' No operation begins in this project during a test run, so an error level left
    ' by an earlier test persists, and at eelCritical LoadComponentFromText reports
    ' failure even after the query loads.
    eelSavedLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError
    On Error GoTo ErrHandler

    DeleteObjectIfExists acQuery, TEST_QUERY
    strFolder = GetTempFolder("VCS") & PathSep
    strBas = strFolder & TEST_QUERY & ".bas"
    strSql = strFolder & TEST_QUERY & ".sql"

    WriteFile LegacyBasDefinition("FromBas"), strBas
    WriteFile "SELECT 2 AS FromSql;" & vbCrLf, strSql

    Set cQuery = New clsDbQuery

    Options.ForceImportOriginalQuerySQL = False
    DeleteObjectIfExists acQuery, TEST_QUERY
    cQuery.Import strBas
    strImported = CurrentDb.QueryDefs(TEST_QUERY).SQL
    TestAssert InStr(1, strImported, "FromBas", vbTextCompare) > 0, _
        "Force SQL off imports the .bas definition"
    TestAssert InStr(1, strImported, "FromSql", vbTextCompare) = 0, _
        "Force SQL off does not apply the paired .sql"

    Options.ForceImportOriginalQuerySQL = True
    DeleteObjectIfExists acQuery, TEST_QUERY
    Set cQuery = New clsDbQuery
    cQuery.Import strBas
    strImported = CurrentDb.QueryDefs(TEST_QUERY).SQL
    TestAssert InStr(1, strImported, "FromSql", vbTextCompare) > 0, _
        "Force SQL on overlays the paired .sql"
    TestAssert InStr(1, strImported, "FromBas", vbTextCompare) = 0, _
        "Force SQL on does not keep the .bas SQL"

    TestAssert SourceFileIsMissing(strFolder & TEST_QUERY & ".qdef"), _
        "missing .qdef is reported instead of treated as present"

CleanUp:
    On Error Resume Next
    Options.ForceImportOriginalQuerySQL = blnSavedForce
    VCSIndex.Disabled = blnIndexDisabled
    Operation.ErrorLevel = eelSavedLevel
    DeleteObjectIfExists acQuery, TEST_QUERY
    If Len(strFolder) > 0 Then
        If FSO.FolderExists(StripSlash(strFolder)) Then FSO.DeleteFolder StripSlash(strFolder), True
    End If
    If lngErr <> 0 Then TestAssert False, _
        "unexpected legacy query import error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestLoadSingleObjectMissingSource
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : The Load Selected missing-file decision must be true for an empty path
'           : and a nonexistent file, and false when the file exists. LoadSingleObject
'           : itself logs eelError on this path, so the harness asserts the decision
'           : rather than invoking that side effect.
'---------------------------------------------------------------------------------------
'
Public Sub TestLoadSingleObjectMissingSource()

    Dim strFolder As String
    Dim strFile As String

    strFolder = GetTempFolder("VCS") & PathSep
    strFile = strFolder & "present.bas"

    TestAssert SourceFileIsMissing(vbNullString), "empty path is missing"
    TestAssert SourceFileIsMissing(strFolder & "no-such-object.bas"), _
        "nonexistent file is missing"

    WriteFile "x", strFile
    TestAssert Not SourceFileIsMissing(strFile), "existing file is not missing"

    If FSO.FolderExists(StripSlash(strFolder)) Then FSO.DeleteFolder StripSlash(strFolder), True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestLegacyQueryAlternateExportIncludesSql
' Author    : Josh
' Date      : 10/20/2026
' Purpose   : A legacy alternate-path (temp) export writes the paired .sql, so a fresh
'           : comparison export hashes the same complete content as a normal export.
'           : Without it, a one-sided source edit after an export looked like an
'           : independent database change and was blocked.
'---------------------------------------------------------------------------------------
'
Public Sub TestLegacyQueryAlternateExportIncludesSql()

    Dim cQuery As IDbComponent
    Dim strFolder As String
    Dim strPriorExport As String
    Dim lngPriorFormat As Long
    Dim blnPriorSaveSql As Boolean
    Dim blnIndexDisabled As Boolean
    Dim blnCaptured As Boolean
    Dim strNormal As String
    Dim strAlternate As String
    Dim strBaseline As String
    Dim eelSavedLevel As eErrorLevel
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    strPriorExport = Options.ExportFolder
    lngPriorFormat = Options.ExportFormatVersion
    blnPriorSaveSql = Options.SaveQuerySQL
    blnIndexDisabled = VCSIndex.Disabled
    eelSavedLevel = Operation.ErrorLevel
    blnCaptured = True
    VCSIndex.Disabled = True
    Operation.ErrorLevel = eelNoError

    strFolder = GetTempFolder("VCSLegacyAltQuery")
    Options.ExportFolder = strFolder
    Options.ExportFormatVersion = EFV_4_1_2
    Options.SaveQuerySQL = True
    DeleteObjectIfExists acQuery, TEST_ALT_QUERY
    CurrentDb.CreateQueryDef TEST_ALT_QUERY, "SELECT 1 AS AltExport;"
    ' The export reads SQL through the cached SharedDb handle.
    SharedDb.QueryDefs.Refresh

    Set cQuery = New clsDbQuery
    Set cQuery.DbObject = CurrentData.AllQueries(TEST_ALT_QUERY)
    strNormal = cQuery.SourceFile
    TestAssert LCase$(FSO.GetExtensionName(strNormal)) = "bas", "format 4.1.2 exports the legacy .bas"
    cQuery.Export
    strAlternate = strFolder & PathSep & "alt" & PathSep & "queries" & PathSep & FSO.GetFileName(strNormal)
    cQuery.Export strAlternate

    TestAssert FSO.FileExists(SwapExtension(strNormal, "sql")), "a normal legacy export writes the paired .sql"
    TestAssert FSO.FileExists(SwapExtension(strAlternate, "sql")), "an alternate legacy export writes the paired .sql"
    strBaseline = VCSIndex.SingleObjectComparisonHash(cQuery, strNormal)
    TestAssert VCSIndex.SingleObjectComparisonHash(cQuery, strAlternate) = strBaseline, _
        "a fresh comparison export matches the normal export's complete content"

    WriteFile LegacyBasDefinition("SourceEdit"), strNormal
    TestAssert Not VCSIndex.SingleObjectContentConflict(VCSIndex.SingleObjectComparisonHash(cQuery, strNormal), _
        VCSIndex.SingleObjectComparisonHash(cQuery, strAlternate), strBaseline, eatImport), _
        "a one-sided source edit after an export is not a conflict"

CleanUp:
    On Error Resume Next
    If blnCaptured Then
        Options.ExportFolder = strPriorExport
        Options.ExportFormatVersion = lngPriorFormat
        Options.SaveQuerySQL = blnPriorSaveSql
        VCSIndex.Disabled = blnIndexDisabled
        Operation.ErrorLevel = eelSavedLevel
    End If
    DeleteObjectIfExists acQuery, TEST_ALT_QUERY
    If Len(strFolder) > 0 Then
        If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    End If
    If lngErr <> 0 Then TestAssert False, _
        "unexpected legacy alternate export error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Function  : LegacyBasDefinition
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Minimal SQL-view SaveAsText body with a distinctive output alias.
'---------------------------------------------------------------------------------------
'
Private Function LegacyBasDefinition(strAlias As String) As String
    LegacyBasDefinition = _
        "dbMemo ""SQL"" =""SELECT 1 AS " & strAlias & ";""" & vbCrLf & _
        "dbBoolean ""ReturnsRecords"" =""-1""" & vbCrLf & _
        "dbInteger ""ODBCTimeout"" =""60""" & vbCrLf & _
        "Begin" & vbCrLf & _
        "End" & vbCrLf
End Function

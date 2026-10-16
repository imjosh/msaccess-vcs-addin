Attribute VB_Name = "modTestIndex"
'---------------------------------------------------------------------------------------
' Module    : modTestIndex
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Integration tests for clsVCSIndex: update, item retrieval, exists check.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("integration")


Public Sub TestIndexExists()
    Dim cModule As IDbComponent
    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub
    If Not FSO.FileExists(cModule.SourceFile) Then Exit Sub

    ' The add-in's own modules should be in the index (they were exported)
    TestAssert VCSIndex.Exists(cModule, cModule.SourceFile), _
        "exported module exists in index"
End Sub


Public Sub TestIndexItem_ReturnsObject()
    Dim cModule As IDbComponent
    Dim cIdx As Object
    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub
    If Not FSO.FileExists(cModule.SourceFile) Then Exit Sub
    If Not VCSIndex.Exists(cModule, cModule.SourceFile) Then Exit Sub

    Set cIdx = VCSIndex.Item(cModule)
    TestAssert Not cIdx Is Nothing, "Item returns non-Nothing"
End Sub


Public Sub TestIndexExists_NonexistentItem()
    Dim cModule As IDbComponent
    Set cModule = New clsDbModule

    ' A component with no DbObject set shouldn't be in the index
    TestAssert Not VCSIndex.Exists(cModule, "C:\nonexistent\fake_module.bas"), _
        "nonexistent item not in index"
End Sub


Public Sub TestIndexConflictsObject()
    TestAssert Not VCSIndex.Conflicts Is Nothing, "Conflicts returns object"
    TestAssert VCSIndex.Conflicts.Count >= 0, "Conflicts count is non-negative"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestIndexSchemaStateRoundTrip
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : The SchemaState section drives re-export of external database schemas, so
'           : it has to survive a save/load cycle of the binary index format.
'---------------------------------------------------------------------------------------
'
Public Sub TestIndexSchemaStateRoundTrip()

    Dim cSave As clsVCSIndex
    Dim cLoad As clsVCSIndex
    Dim strFolder As String

    strFolder = GetTempFolder("VCSIdx")

    Set cSave = New clsVCSIndex
    cSave.SchemaState("Sample Connection") = "rev=1;ddl=sp_GetDDL"
    cSave.Save strFolder

    Set cLoad = New clsVCSIndex
    cLoad.LoadFromFile StripSlash(strFolder) & PathSep & "vcs-index.idx"

    TestAssert cLoad.SchemaState("Sample Connection") = "rev=1;ddl=sp_GetDDL", _
        "schema state survives index round trip"

    ' Schema names are matched without regard to case
    TestAssert cLoad.SchemaState("sample connection") = "rev=1;ddl=sp_GetDDL", _
        "schema state lookup is case insensitive"

    ' An unrecorded schema returns an empty string, which forces a baseline re-export
    TestAssert cLoad.SchemaState("Never Exported") = vbNullString, _
        "unrecorded schema returns empty state"

    LogUnhandledErrors
    On Error Resume Next
    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder StripSlash(strFolder), True
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetTestComponent
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Helper to get a real module IDbComponent.
'---------------------------------------------------------------------------------------
'
Private Function GetTestComponent() As IDbComponent
    If CurrentProject.AllModules.Count = 0 Then Exit Function
    Dim cModule As IDbComponent
    Set cModule = New clsDbModule
    Set cModule.DbObject = CurrentProject.AllModules(0)
    Set GetTestComponent = cModule
End Function


Public Sub TestSingleObjectContentHistoryRules()
    Dim cIndex As New clsVCSIndex
    TestAssert Not cIndex.SingleObjectContentConflict("same", "same", "", eatImport), "identical contents are safe without history"
    TestAssert Not cIndex.SingleObjectContentConflict("new source", "normalized database", "normalized database", eatImport), "ordinary source edit preserves unchanged normalized database baseline"
    TestAssert Not cIndex.SingleObjectContentConflict("old source", "new database", "old source", eatExport), "ordinary database edit preserves unchanged source baseline"
    TestAssert cIndex.SingleObjectContentConflict("source edit", "database edit", "old database", eatImport), "independent database edit needs a decision"
    TestAssert cIndex.SingleObjectContentConflict("source edit", "database edit", "old source", eatExport), "independent source edit needs a decision"
    TestAssert cIndex.SingleObjectContentConflict("source edit", "database", "", eatImport), "missing database history is conservative"
    TestAssert cIndex.SingleObjectContentConflict("source", "database edit", "", eatExport), "old disabled-index history is conservative"
    TestAssert Not cIndex.SingleObjectContentConflict("source", "", "", eatImport, False), "new database object has no destination collision"
    TestAssert Not cIndex.SingleObjectContentConflict("", "database", "", eatExport, False), "new source file has no destination collision"
End Sub


Public Sub TestIndexDatabaseAndSourceBaselinesSurviveReopen()
    On Error GoTo ErrHandler
    Dim cSave As New clsVCSIndex
    Dim cLoad As New clsVCSIndex
    Dim cComponent As IDbComponent
    Dim cEntry As clsVCSIndexItem
    Dim strFolder As String
    Dim lngErr As Long
    Dim strErr As String

    strFolder = GetTempFolder("VCSF1Idx")
    Set cComponent = New clsDbQuery
    Set cEntry = cSave.Item(cComponent, "qryHistory.sql")
    cEntry.AllFilesHash = "source bytes"
    cEntry.DatabaseFilesHash = "normalized database bytes"
    cEntry.SourceFilesHash = "complete source bytes"
    cEntry.ImportDate = #10/16/2026 12:00:00 PM#
    cSave.Item(cComponent, "qryUnrelated.sql").DatabaseFilesHash = "unrelated baseline"
    cSave.Save strFolder
    cLoad.LoadFromFile strFolder & PathSep & "vcs-index.idx"
    TestAssert cLoad.Exists(cComponent, "qryHistory.sql"), "history entry survives index reopen"
    Set cEntry = cLoad.Item(cComponent, "qryHistory.sql")
    TestAssert cEntry.AllFilesHash = "source bytes", "source baseline remains separate"
    TestAssert cEntry.DatabaseFilesHash = "normalized database bytes", "actual database baseline survives reopen"
    TestAssert cEntry.SourceFilesHash = "complete source bytes", "complete source baseline survives reopen"
    TestAssert cLoad.Item(cComponent, "qryUnrelated.sql").DatabaseFilesHash = "unrelated baseline", "unrelated object history is retained"

CleanUp:
    On Error Resume Next
    If Len(strFolder) > 0 Then If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    If lngErr <> 0 Then TestAssert False, "index baseline round trip " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestIndexVersionsThreeAndFourKeepMissingDatabaseHistoryConservative()
    On Error GoTo ErrHandler
    Dim cSave As clsVCSIndex
    Dim cLoad As clsVCSIndex
    Dim cComponent As IDbComponent
    Dim strFolder As String
    Dim strFile As String
    Dim bData() As Byte
    Dim varVersion As Variant
    Dim lngErr As Long
    Dim strErr As String

    strFolder = GetTempFolder("VCSF1Legacy")
    strFile = strFolder & PathSep & "vcs-index.idx"
    Set cComponent = New clsDbQuery
    For Each varVersion In Array(3, 4, 5)
        Set cSave = New clsVCSIndex
        cSave.Item(cComponent, "qryLegacy.sql").AllFilesHash = "legacy source baseline"
        If varVersion = 5 Then cSave.Item(cComponent, "qryLegacy.sql").DatabaseFilesHash = "old incomplete database hash"
        cSave.Save strFolder
        ' Disposable format fixture only. No production index is modified.
        ' Without optional database hashes, v4 has the same component layout.
        bData = GetFileBytes(strFile)
        bData(4) = CByte(varVersion)
        If varVersion = 3 Then ReDim Preserve bData(0 To UBound(bData) - 4)
        WriteBinaryFile strFile, bData
        Set cLoad = New clsVCSIndex
        cLoad.LoadFromFile strFile
        TestAssert cLoad.Exists(cComponent, "qryLegacy.sql"), "legacy index is readable"
        TestAssert cLoad.Item(cComponent, "qryLegacy.sql").AllFilesHash = "legacy source baseline", "legacy source baseline is retained"
        TestAssert cLoad.SingleObjectDestinationBaseline(cLoad.Item(cComponent, "qryLegacy.sql"), eatImport) = "", "legacy and unpublished v5 history do not invent a complete database baseline"
        TestAssert cLoad.SingleObjectContentConflict("source edit", "database", _
            cLoad.SingleObjectDestinationBaseline(cLoad.Item(cComponent, "qryLegacy.sql"), eatImport), eatImport), "missing legacy database baseline requires review"
    Next varVersion

CleanUp:
    On Error Resume Next
    If Len(strFolder) > 0 Then If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    If lngErr <> 0 Then TestAssert False, "legacy index migration " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestCompleteComparisonIncludesCommandBarImagesAndThemeDirectories()
    On Error GoTo ErrHandler
    Dim cIndex As New clsVCSIndex
    Dim cComponent As IDbComponent
    Dim strFolder As String
    Dim strFirst As String
    Dim strSecond As String
    Dim lngErr As Long
    Dim strErr As String
    strFolder = GetTempFolder("VCSF1Artifacts")
    strFirst = strFolder & PathSep & "first" & PathSep
    strSecond = strFolder & PathSep & "second" & PathSep
    WriteFile "same JSON", strFirst & "Menu.json"
    WriteFile "same JSON", strSecond & "Menu.json"
    WriteFile "first image", strFirst & "Menu_Images" & PathSep & "Button.bmp"
    WriteFile "first image", strSecond & "Menu_Images" & PathSep & "button.bmp"
    Set cComponent = New clsDbCommandBar
    TestAssert cIndex.SingleObjectComparisonHash(cComponent, strFirst & "Menu.json") = _
        cIndex.SingleObjectComparisonHash(cComponent, strSecond & "Menu.json"), "Windows-relative image names are case insensitive"
    WriteFile "independent image", strSecond & "Menu_Images" & PathSep & "button.bmp"
    TestAssert cIndex.SingleObjectComparisonHash(cComponent, strFirst & "Menu.json") <> _
        cIndex.SingleObjectComparisonHash(cComponent, strSecond & "Menu.json"), "image-only changes alter the complete content hash"
    Set cComponent = New clsDbTheme
    WriteFile "theme XML", strFirst & "Theme" & PathSep & "theme" & PathSep & "theme1.xml"
    WriteFile "theme XML", strSecond & "Theme" & PathSep & "theme" & PathSep & "theme1.xml"
    TestAssert cIndex.SingleObjectComparisonHash(cComponent, strFirst & "Theme.thmx") = _
        cIndex.SingleObjectComparisonHash(cComponent, strSecond & "Theme.thmx"), "extracted theme content compares without a thmx file"
    WriteFile "independent theme XML", strSecond & "Theme" & PathSep & "theme" & PathSep & "theme1.xml"
    TestAssert cIndex.SingleObjectComparisonHash(cComponent, strFirst & "Theme.thmx") <> _
        cIndex.SingleObjectComparisonHash(cComponent, strSecond & "Theme.thmx"), "nested theme edits alter the complete content hash"
CleanUp:
    On Error Resume Next
    If Len(strFolder) > 0 Then If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    If lngErr <> 0 Then TestAssert False, "complete artifact comparison " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub

Attribute VB_Name = "modOrphaned"
'---------------------------------------------------------------------------------------
' Module    : modOrphaned
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions relating to detecting and removing orphaned items in relation
'           : to the source code files and existing database objects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName = "modOrphaned"


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 2/23/2021
' Purpose   : Clears existing source files that don't have a matching object in the
'           : database.
'           : Note that this function is integrated with the index, so deleted files
'           : are removed from the index, and potential conflicts are logged as well.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFiles(cType As IDbComponent)

    Dim dBaseNames As Dictionary
    Dim dExtensions As Dictionary
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim varExt As Variant

    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub

    ' Set up dictionary objects for case-insensitive comparison
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    Set dExtensions = New Dictionary
    dExtensions.CompareMode = TextCompare

    ' Cache a list of base source file names for actual database objects
    Perf.OperationStart "Clear Orphaned Files"
    Set dItems = cType.GetAllFromDB(False)
    For Each varKey In dItems.Keys
        dBaseNames.Add FSO.GetBaseName(varKey), vbNullString
        If cType.SingleFile Then Exit For
    Next varKey

    ' Build dictionary of included extensions
    For Each varExt In cType.FileExtensions
        dExtensions.Add varExt, vbNullString
    Next varExt

    ' Check for single-file exports with no items
    If cType.SingleFile Then
        If dItems.Count = 0 Then
            ' No more of these items exist in the database.
            ' (For example, IMEX specs)
            If FSO.FileExists(cType.SourceFile) Then
                ' Compare to index to check for any source changes.
                CompareToIndex cType, cType.SourceFile, dExtensions, dBaseNames
            End If
        End If
    Else
        ' Remove sub-artifacts (sidecars, per-object folders) before the generic scan.
        ' File sidecars are data-driven from FileExtensions(efesAll); folder artifacts
        ' (command-bar _Images, extracted theme folders) are dispatched by type.
        If Not InteractionIsNonInteractive() Then
            ClearOrphanedComponentArtifacts cType, dBaseNames
            ClearOrphanedComponentFolders cType, dBaseNames
        End If
        ' Loop through files in folder (and subfolders for @Folder support)
        ScanFolderForOrphans cType, StripSlash(cType.BaseFolder), dExtensions, dBaseNames

        ' Remove base folder if we don't have any files in it
        If FSO.GetFolder(cType.BaseFolder).Files.Count = 0 _
            And FSO.GetFolder(cType.BaseFolder).SubFolders.Count = 0 Then
            LogUnhandledErrors
            On Error Resume Next
            FSO.DeleteFolder StripSlash(cType.BaseFolder), True
            CatchAny eelWarning, "Unable to delete empty folder: " & cType.BaseFolder, ModuleName & ".ClearOrphanedSourceFiles"
        End If
    End If

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScanFolderForOrphans
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Recursively scan a folder and its subfolders for orphaned source files.
'           : Removes empty subfolders after processing. Uses Win32 API via
'           : ScanFolderContents for fast enumeration without FSO COM overhead.
'---------------------------------------------------------------------------------------
'
Private Sub ScanFolderForOrphans(cType As IDbComponent, strFolder As String, _
    dExtensions As Dictionary, dBaseNames As Dictionary)

    Dim colFiles As New Collection
    Dim colSubFolders As New Collection
    Dim varItem As Variant

    ' A prior group cleanup can remove a folder captured by its parent's scan.
    If Not FSO.FolderExists(strFolder) Then Exit Sub
    ' Single-pass Win32 API scan for files and subfolders
    ScanFolderContents strFolder, colFiles, colSubFolders

    ' Process files in this folder. This disk scan is near-instant and is intentionally
    ' excluded from the scan progress bar (see modExport - the bar is sized to the object
    ' count only), so we do not increment progress here.
    For Each varItem In colFiles
        ' A complete unchanged group may already have removed this companion.
        If FSO.FileExists(CStr(varItem)) Then
            CompareToIndex cType, CStr(varItem), dExtensions, dBaseNames
        End If
    Next varItem

    ' Recurse into subfolders
    For Each varItem In colSubFolders
        ' Removing an unchanged primary also removes its nested artifact directory.
        If FSO.FolderExists(CStr(varItem)) Then
            ScanFolderForOrphans cType, CStr(varItem), dExtensions, dBaseNames
            ' Recursion may itself have removed the captured directory.
            If FSO.FolderExists(CStr(varItem)) Then
                If FSO.GetFolder(CStr(varItem)).Files.Count = 0 _
                    And FSO.GetFolder(CStr(varItem)).SubFolders.Count = 0 Then
                    LogUnhandledErrors
                    On Error Resume Next
                    FSO.DeleteFolder CStr(varItem), True
                    CatchAny eelWarning, "Unable to delete empty folder: " & CStr(varItem), ModuleName & ".ScanFolderForOrphans"
                End If
            End If
        End If
    Next varItem

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompareToIndex
' Author    : Adam Waller
' Date      : 12/5/2023
' Purpose   : Compare the potential orphaned file to the index to determine if we need
'           : to list this as a possible conflict item.
'---------------------------------------------------------------------------------------
'
Private Sub CompareToIndex(cType As IDbComponent, strFilePath As String, dExtensions As Dictionary, dBaseNames As Dictionary)

    Dim strFileName As String
    Dim strBaseName As String
    Dim strExt As String
    Dim strHash As String
    Dim blnUnchanged As Boolean
    Dim cEntry As clsVCSIndexItem
    Dim colFileExtensions As Collection
    Dim strPrimaryExtension As String
    Dim strIndexedPrimary As String
    Dim strIndexFile As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    ' A captured companion may have been deleted with its complete primary group.
    If Not FSO.FileExists(strFilePath) Then Exit Sub
    strFileName = FSO.GetFileName(strFilePath)
    strBaseName = FSO.GetBaseName(strFileName)
    strExt = Mid$(strFileName, Len(strBaseName) + 2)
    If dBaseNames.Exists(strBaseName) Then Exit Sub
    If InteractionIsNonInteractive() Then
        ' Actual history owns group identity, including .cls modules and primaries
        ' retained from an earlier query, layout, table or macro source format.
        strIndexedPrimary = VCSIndex.IndexedSourceGroupFile(cType, strFilePath)
        If Len(strIndexedPrimary) > 0 Then
            strPrimaryExtension = FSO.GetExtensionName(strIndexedPrimary)
        Else
            Set colFileExtensions = cType.FileExtensions
            If colFileExtensions.Count = 0 Then Exit Sub
            strPrimaryExtension = CStr(colFileExtensions(1))
        End If
    Else
        ' Keep the established interactive cleanup path outside this policy repair.
        strPrimaryExtension = CStr(dExtensions(0))
    End If
    strIndexFile = Nz2(strIndexedPrimary, strFileName)


    ' See if extension exists in cached list
    If dExtensions.Exists(strExt) Or Len(strIndexedPrimary) > 0 Then

        ' See if base file name exists in list of database objects
        If Not dBaseNames.Exists(strBaseName) Then

            ' See if this is the primary file extension for this component type
            If StrComp(strExt, strPrimaryExtension, vbTextCompare) = 0 Then

                ' Object not found in database. Check the index
                If VCSIndex.Exists(cType, strIndexFile) Then

                    ' If file is unchanged from the index, we can go ahead and delete it.
                    ' (The source file matches the last version imported or exported)
                    strHash = GetSourceFilesPropertyHash(cType, strFilePath)
                    Set cEntry = VCSIndex.Item(cType, strIndexFile)
                    blnUnchanged = (cEntry.FilePropertiesHash = strHash)
                    If InteractionIsNonInteractive() Then
                        blnUnchanged = (Len(cEntry.SourceFilesHash) > 0)
                        If blnUnchanged Then blnUnchanged = (cEntry.SourceFilesHash = VCSIndex.SingleObjectComparisonHash(cType, strFilePath))
                    End If
                    If blnUnchanged Then

                        ' Remove file and index entry
                        Log.Add "  Removing orphaned file: " & cType.BaseFolder & strFileName, Options.ShowDebug
                        If InteractionIsNonInteractive() Then
                            DeleteUnchangedSourceGroup cType, strFilePath
                        Else
                            DeleteFile strFilePath, True
                        End If
                        If Operation.ErrorLevel < eelCritical Then
                            VCSIndex.Remove cType, strIndexFile
                            ' Do not classify captured companion paths as new orphans.
                            If InteractionIsNonInteractive() And Not dBaseNames.Exists(strBaseName) Then
                                dBaseNames.Add strBaseName, vbNullString
                            End If
                        End If
                    Else
                        ' File properties different from index. Add as a conflict to resolve.
                        ' (This can happen when the last export was during a different daylight savings time
                        ' setting, as the past file modified date returned by FSO is not adjusted for DST.)
                        Log.Add "  Orphaned source file does not match last export: " & strFilePath, Options.ShowDebug
                        VCSIndex.Conflicts.Add cType, strFilePath, 0, GetSourceModifiedDate(cType, strFilePath), ercDelete, strFilePath, ercDelete, _
                            strClassification:="tracked_changed_orphan", blnTracked:=True
                    End If
                Else
                    ' Object does not exist in the index. It might be a new file added
                    ' by another developer. Don't delete it, as it may need to be merged
                    ' into the database. (Defaults to skip deleting the file)
                    Log.Add "  Found new source file: " & strFilePath, Options.ShowDebug
                    VCSIndex.Conflicts.Add cType, strFilePath, 0, GetSourceModifiedDate(cType, strFilePath), ercDelete, strFilePath, ercSkip, _
                        strClassification:="untracked_orphan", blnTracked:=False
                End If

            Else
                ' Not the primary extension for this component type.
                ' If the primary source file exists, we will let that file handle evaluate any conflicts
                If Not FSO.FileExists(SwapExtension(strFilePath, strPrimaryExtension)) Then
                    If InteractionIsNonInteractive() Then
                        CompareOrphanedCompanion cType, strFilePath, strPrimaryExtension, strIndexedPrimary
                        ' One logical orphan group may have several companion files.
                        ' Its complete-content decision owns the remaining members.
                        If Not dBaseNames.Exists(strBaseName) Then dBaseNames.Add strBaseName, vbNullString
                    Else
                        ' Preserve the legacy direct cleanup path without an explicit policy.
                        Log.Add "  Removing orphaned file: " & cType.BaseFolder & strFileName, Options.ShowDebug
                        DeleteFile strFilePath, True
                    End If
                End If
            End If
        End If
    End If

    CatchAny eelCritical, T("Unable to classify the orphan source content safely."), _
        ModuleName & ".CompareToIndex", True, True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedComponentArtifacts
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Remove object-named sidecar files declared in FileExtensions(efesAll) but
'           : not in FileExtensions(efesIndexed). Called for every component type during
'           : orphan cleanup; a no-op wherever efesAll equals efesIndexed.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedComponentArtifacts(cmp As IDbComponent, dObjectBaseNames As Dictionary)

    Dim dArtifactExts As Dictionary

    Set dArtifactExts = GetArtifactOnlyExtensions(cmp)
    If dArtifactExts.Count = 0 Then Exit Sub
    If Not FSO.FolderExists(cmp.BaseFolder) Then Exit Sub
    ClearOrphanedArtifactFilesInFolder cmp, StripSlash(cmp.BaseFolder), dArtifactExts, dObjectBaseNames

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedComponentFolders
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Remove per-object artifact folders that the generic file scan cannot see.
'           : Only two component types produce such folders, so they are dispatched here
'           : explicitly rather than via a mostly no-op interface method on all 29
'           : component classes. A new folder-producing component adds one branch here.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedComponentFolders(cmp As IDbComponent, dObjectBaseNames As Dictionary)

    If TypeOf cmp Is clsDbCommandBar Then
        ' Command-bar images live in "<Bar>_Images" subfolders.
        ClearOrphanedArtifactFolders cmp, dObjectBaseNames, "_Images"
    ElseIf TypeOf cmp Is clsDbTheme Then
        ' Extracted themes live in a subfolder named after the theme.
        ClearOrphanedArtifactFolders cmp, dObjectBaseNames
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetArtifactOnlyExtensions
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Return extensions present in FileExtensions(efesAll) but not in
'           : FileExtensions(efesIndexed) as a case-insensitive dictionary.
'---------------------------------------------------------------------------------------
'
Private Function GetArtifactOnlyExtensions(cmp As IDbComponent) As Dictionary

    Dim dResult As Dictionary
    Dim colIndexed As Collection
    Dim colAll As Collection
    Dim varExt As Variant

    Set dResult = New Dictionary
    dResult.CompareMode = TextCompare
    Set colIndexed = cmp.FileExtensions(efesIndexed)
    Set colAll = cmp.FileExtensions(efesAll)

    For Each varExt In colAll
        If Not ExtensionInCollection(CStr(varExt), colIndexed) Then
            If Not dResult.Exists(CStr(varExt)) Then dResult.Add CStr(varExt), vbNullString
        End If
    Next varExt

    Set GetArtifactOnlyExtensions = dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtensionInCollection
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Case-insensitive membership test for a file extension collection.
'---------------------------------------------------------------------------------------
'
Private Function ExtensionInCollection(strExt As String, colExts As Collection) As Boolean

    Dim varItem As Variant

    For Each varItem In colExts
        If StrComp(CStr(varItem), strExt, vbTextCompare) = 0 Then
            ExtensionInCollection = True
            Exit Function
        End If
    Next varItem

End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedArtifactFiles
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Recursively delete files in a component's BaseFolder whose extension is
'           : listed in Extensions and whose base name is not a current database object.
'           : Used for sidecar files (e.g. form/report .json and .svg) not declared in
'           : FileExtensions and therefore invisible to CompareToIndex.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedArtifactFiles(cType As IDbComponent, dObjectBaseNames As Dictionary, _
    ParamArray Extensions())

    Dim dExtensions As Dictionary
    Dim varExt As Variant

    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub

    Set dExtensions = New Dictionary
    dExtensions.CompareMode = TextCompare
    For Each varExt In Extensions
        dExtensions.Add CStr(varExt), vbNullString
    Next varExt

    ClearOrphanedArtifactFilesInFolder cType, StripSlash(cType.BaseFolder), dExtensions, dObjectBaseNames

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedArtifactFilesInFolder
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Recursive helper for ClearOrphanedArtifactFiles.
'---------------------------------------------------------------------------------------
'
Private Sub ClearOrphanedArtifactFilesInFolder(cType As IDbComponent, strFolder As String, _
    dExtensions As Dictionary, dObjectBaseNames As Dictionary)

    Dim colFiles As New Collection
    Dim colSubFolders As New Collection
    Dim varItem As Variant
    Dim strFileName As String
    Dim strBaseName As String
    Dim strExt As String

    ScanFolderContents strFolder, colFiles, colSubFolders

    For Each varItem In colFiles
        strFileName = FSO.GetFileName(CStr(varItem))
        strBaseName = FSO.GetBaseName(strFileName)
        strExt = Mid$(strFileName, Len(strBaseName) + 2)
        If dExtensions.Exists(strExt) Then
            If Not dObjectBaseNames.Exists(strBaseName) Then
                Log.Add "  Removing orphaned artifact file: " & CStr(varItem), Options.ShowDebug
                DeleteFile CStr(varItem), True
            End If
        End If
    Next varItem

    For Each varItem In colSubFolders
        ClearOrphanedArtifactFilesInFolder cType, CStr(varItem), dExtensions, dObjectBaseNames
    Next varItem

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedArtifactFolders
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Delete immediate subfolders of a component's BaseFolder when the folder
'           : name (minus an optional suffix such as "_Images") is not a current
'           : database object. Used for per-object artifact folders (command-bar images,
'           : extracted theme folders).
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedArtifactFolders(cType As IDbComponent, dObjectBaseNames As Dictionary, _
    Optional strSuffix As String = vbNullString)

    Dim colFiles As New Collection
    Dim colSubFolders As New Collection
    Dim varItem As Variant
    Dim strFolderName As String
    Dim strBaseName As String

    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub

    ScanFolderContents StripSlash(cType.BaseFolder), colFiles, colSubFolders

    For Each varItem In colSubFolders
        strFolderName = FSO.GetFileName(CStr(varItem))
        If Len(strSuffix) > 0 Then
            If Right$(strFolderName, Len(strSuffix)) = strSuffix Then
                strBaseName = Left$(strFolderName, Len(strFolderName) - Len(strSuffix))
            Else
                strBaseName = vbNullString
            End If
        Else
            strBaseName = strFolderName
        End If
        If Len(strBaseName) > 0 Then
            If Not dObjectBaseNames.Exists(strBaseName) Then
                Log.Add "  Removing orphaned artifact folder: " & CStr(varItem), Options.ShowDebug
                LogUnhandledErrors
                On Error Resume Next
                FSO.DeleteFolder CStr(varItem), True
                CatchAny eelWarning, "Unable to delete orphaned artifact folder: " & CStr(varItem), ModuleName & ".ClearOrphanedArtifactFolders"
            End If
        End If
    Next varItem

End Sub


' Called only after complete content history proves this indexed orphan unchanged.
'---------------------------------------------------------------------------------------
' Procedure : DeleteUnchangedSourceGroup
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Delete complete orphan artifacts only after history proves the source group unchanged.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteUnchangedSourceGroup(ByVal cType As IDbComponent, ByVal strPrimary As String)
    Dim varExt As Variant
    Dim strBase As String
    Dim strFile As String
    Dim strFolder As String
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    strBase = FSO.GetParentFolderName(strPrimary) & PathSep & FSO.GetBaseName(strPrimary)
    For Each varExt In cType.FileExtensions(efesAll)
        strFile = strBase & "." & CStr(varExt)
        If FSO.FileExists(strFile) Then DeleteFile strFile, True
    Next varExt
    If FSO.FileExists(strPrimary) Then DeleteFile strPrimary, True
    Select Case cType.ComponentType
        Case edbCommandBar: strFolder = strBase & "_Images"
        Case edbTheme: strFolder = strBase
    End Select
    If Len(strFolder) > 0 Then
        If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    End If
    CatchAny eelCritical, T("Unable to remove the unchanged orphan source group."), ModuleName & ".DeleteUnchangedSourceGroup", True, True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompareOrphanedCompanion
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Protect a complete orphan group when only metadata companions remain.
'---------------------------------------------------------------------------------------
'
Private Sub CompareOrphanedCompanion(ByVal cType As IDbComponent, ByVal strCompanion As String, _
    ByVal strPrimaryExtension As String, Optional ByVal strIndexedPrimary As String)

    Dim strPrimary As String
    Dim cEntry As clsVCSIndexItem
    Dim blnUnchanged As Boolean
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    strPrimary = Nz2(strIndexedPrimary, SwapExtension(strCompanion, strPrimaryExtension))
    If VCSIndex.Exists(cType, strPrimary) Then
        Set cEntry = VCSIndex.Item(cType, strPrimary)
        blnUnchanged = (Len(cEntry.SourceFilesHash) > 0)
        If blnUnchanged Then blnUnchanged = (cEntry.SourceFilesHash = VCSIndex.SingleObjectComparisonHash(cType, strPrimary))
        If Operation.ErrorLevel >= eelCritical Then Exit Sub
        If blnUnchanged Then
            DeleteUnchangedSourceGroup cType, strPrimary
            If Operation.ErrorLevel < eelCritical Then VCSIndex.Remove cType, strPrimary
        Else
            VCSIndex.Conflicts.Add cType, strCompanion, cEntry.SyncDate, _
                GetSourceModifiedDate(cType, strPrimary), ercDelete, strCompanion, ercDelete, _
                strClassification:="tracked_changed_orphan", blnTracked:=True
        End If
    Else
        VCSIndex.Conflicts.Add cType, strCompanion, 0, GetSourceModifiedDate(cType, strPrimary), _
            ercDelete, strCompanion, ercSkip, strClassification:="untracked_orphan", blnTracked:=False
    End If
    CatchAny eelCritical, T("Unable to classify the complete orphan source group safely."), _
        ModuleName & ".CompareOrphanedCompanion", True, True
End Sub

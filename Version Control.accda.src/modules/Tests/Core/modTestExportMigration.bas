Attribute VB_Name = "modTestExportMigration"
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")

' Public-entry regression driver. Load the development project as a library in a
' fresh A37_*.accdb disposable host; never mutate a development or user database.
' A Function with an argument is intentionally outside runner discovery.
Public Function ExportMigrationCheck(ByVal strCase As String) As String
    Dim dReport As Dictionary
    Dim colFailures As Collection
    Dim dBefore As Dictionary
    Dim dAfter As Dictionary
    Dim dResult As Dictionary
    Dim dCurrent As Dictionary
    Dim frmNew As Form
    Dim strTemp As String
    Dim strFolder As String
    Dim strQuery As String
    Dim strOriginalQuery As String
    Dim lngErr As Long
    Dim strErr As String
    Dim lngCancelCheck As Long
    Dim cCategory As IDbComponent
    Dim cObject As IDbComponent
    Dim varKey As Variant
    Dim dObjects As Dictionary
    Dim blnPolicy As Boolean

    On Error GoTo ErrHandler
    Set dReport = New Dictionary
    Set colFailures = New Collection
    dReport.Add "case", strCase
    Set dReport("failures") = colFailures
    If Left$(CurrentProject.Name, 4) <> "A37_" _
        Or CurrentProject.FullName = CodeProject.FullName _
        Or CurrentProject.AllForms.Count <> 0 _
        Or CurrentData.AllQueries.Count <> 0 Then
        colFailures.Add "Requires an empty, disposable A37_*.accdb host"
        GoTo CleanUp
    End If

    Set frmNew = Application.CreateForm
    strTemp = frmNew.Name
    frmNew.Caption = "A37 unchanged form"
    DoCmd.Close acForm, strTemp, acSaveYes
    Set frmNew = Nothing
    DoCmd.Rename "frmA37First", acForm, strTemp
    DoCmd.CopyObject , "frmA37Second", acForm, "frmA37First"
    CurrentDb.CreateQueryDef "qryA37Conflict", "SELECT 1 AS [A37Value];"

    Set Options = Nothing
    Options.ExportFormatVersion = EFV_5_1_0
    Options.ExportLayoutSvg = False
    Options.UseGitIntegration = False
    Options.SaveOptionsForProject
    VCS.SetOperationPolicy "block"
    blnPolicy = True
    VCS.FullExport
    Set dResult = Operation.LastCompletion
    Set dReport("baselineOutcome") = dResult
    CheckMigration colFailures, OutcomeSucceeded(dResult), "baseline full export succeeds"
    strFolder = Options.GetExportFolder
    strQuery = strFolder & "queries" & PathSep & "qryA37Conflict.sql"
    strOriginalQuery = ReadFile(strQuery)
    VCSIndex.ExportDate = #1/1/2000#
    VCSIndex.FullExportDate = #1/1/2000#
    Set cCategory = New clsDbForm
    Set dObjects = cCategory.GetAllFromDB(False)
    For Each varKey In dObjects.Keys
        Set cObject = dObjects(varKey)
        VCSIndex.Item(cObject).ExportDate = #1/1/2000#
    Next varKey
    VCSIndex.Save
    Set dBefore = MigrationSnapshot
    Set dReport("before") = dBefore
    CheckMigration colFailures, Not FSO.FileExists(strFolder & "forms\frmA37First.svg"), "baseline has no SVG"

    Options.ExportLayoutSvg = True
    If strCase = "partial_cancel" Or strCase = "skip" Then Options.FormatSQL = Not Options.FormatSQL
    If strCase = "global" Or strCase = "filtered" Then Options.ExportFormatVersion = EFV_5_0_0
    Options.SaveOptionsForProject
    Select Case strCase
        Case "blocked", "skip"
            WriteFile "SELECT 2 AS [A37Value];", strQuery
            SetFileDate strQuery, DateAdd("s", 2, Now), True
            If strCase = "skip" Then
                VCS.ClearOperationPolicy
                VCS.SetOperationPolicy "skip"
            End If
        Case "scan_cancel", "global"
            Operation.SetCancelPollForTest 1
        Case "partial_cancel"
            ' One checkpoint per scan category, then one per selected object.
            ' Stop after the first of two forms has been written.
            lngCancelCheck = GetContainers.Count
            For Each cCategory In GetContainers
                If cCategory.ComponentType = edbForm Then Exit For
                Set dObjects = cCategory.GetAllFromDB(False)
                If cCategory.SingleFile Then
                    If dObjects.Count > 0 Then lngCancelCheck = lngCancelCheck + 1
                Else
                    lngCancelCheck = lngCancelCheck + dObjects.Count
                End If
            Next cCategory
            Operation.SetCancelPollForTest lngCancelCheck + 1
        Case "success", "filtered"
        Case Else
            colFailures.Add "Unknown regression case"
            GoTo CleanUp
    End Select
    If strCase = "filtered" Then VCS.ExportVBA Else VCS.FullExport
    Operation.SetCancelPollForTest 0
    Set dResult = Operation.LastCompletion
    Set dReport("attemptOutcome") = dResult
    Set dAfter = MigrationSnapshot
    Set dCurrent = Options.GetCategoryHashes
    Set dReport("afterAttempt") = dAfter
    If strCase = "success" Or strCase = "skip" Or strCase = "filtered" Then
        CheckMigration colFailures, OutcomeSucceeded(dResult), "completed export succeeds"
        CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmA37First.svg"), "completed form migration writes SVG"
        CheckMigration colFailures, dAfter("CategoryHashes")("Forms") = dCurrent("Forms"), "completed forms record option hash"
        If strCase = "skip" Then
            CheckMigration colFailures, dAfter("CategoryHashes")("Queries") = dBefore("CategoryHashes")("Queries"), "skipped query migration stays pending"
            CheckMigration colFailures, dAfter("FullExportDate") = dBefore("FullExportDate"), "skipped migration does not claim full completion"
        End If
        If strCase = "filtered" Then
            CheckMigration colFailures, dAfter("FullExportDate") = dBefore("FullExportDate"), "VBA export does not claim full project completion"
            CheckMigration colFailures, dAfter("CategoryHashes")("_Global") = dBefore("CategoryHashes")("_Global"), "VBA export leaves global migration pending"
        ElseIf strCase = "success" Then
            CheckMigration colFailures, dAfter("FullExportDate") <> dBefore("FullExportDate"), "successful full export advances completion date"
            CheckMigration colFailures, dAfter("CategoryHashes")("_Global") = dCurrent("_Global"), "successful full export records global hash"
        End If
    Else
        CheckMigration colFailures, Not OutcomeSucceeded(dResult), "aborted export fails publicly"
        If strCase = "blocked" Then
            CheckMigration colFailures, CBool(dResult("decision_required")), "conflict remains decision_required"
            CheckMigration colFailures, Not dResult.Exists("cancelled"), "blocked conflict is not cancellation"
            CheckMigration colFailures, dResult("decisions").Count > 0, "blocked conflict preserves decision journal"
        Else
            CheckMigration colFailures, Operation.Result = eorCanceled, "checkpoint cancellation remains cancelled"
        End If
        CheckMigration colFailures, Len(CStr(dResult("log_path"))) > 0, "abort carries its operation log"
        CheckMigration colFailures, dAfter("FullExportDate") = dBefore("FullExportDate"), "abort preserves full completion date"
        CheckMigration colFailures, dAfter("ExportDate") = dBefore("ExportDate"), "abort preserves export completion date"
        CheckMigration colFailures, dAfter("CategoryHashes")("Forms") = dBefore("CategoryHashes")("Forms"), "unfinished form migration remains pending"
        If strCase = "global" Then
            CheckMigration colFailures, dAfter("CategoryHashes")("_Global") = dBefore("CategoryHashes")("_Global"), "aborted global migration stays pending"
        End If
        If strCase = "partial_cancel" Then
            CheckMigration colFailures, dAfter("CategoryHashes")("Queries") = dCurrent("Queries"), "completed query migration records its new hash"
            CheckMigration colFailures, dAfter("CategoryHashes")("Reports") = dBefore("CategoryHashes")("Reports"), "unvisited report migration stays pending"
            CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmA37First.svg"), "partial progress includes first SVG"
            CheckMigration colFailures, Not FSO.FileExists(strFolder & "forms\frmA37Second.svg"), "second SVG is still pending"
            CheckMigration colFailures, dAfter("Components")("Forms")("frmA37First.form")("ExportDate") <> _
                dBefore("Components")("Forms")("frmA37First.form")("ExportDate"), "partial component progress is persisted"
        Else
            CheckMigration colFailures, Not FSO.FileExists(strFolder & "forms\frmA37First.svg"), "abort did not export unchanged form"
        End If
    End If

    ' Resolve the external conflict by restoring the original source, then fast
    ' export through the same public entry point. Artifacts are the primary oracle.
    If strCase = "blocked" Or strCase = "skip" Then WriteFile strOriginalQuery, strQuery
    VCS.Export
    Set dResult = Operation.LastCompletion
    Set dReport("retryOutcome") = dResult
    Set dAfter = MigrationSnapshot
    Set dReport("afterRetry") = dAfter
    CheckMigration colFailures, dAfter("CategoryHashes")("Forms") = dCurrent("Forms"), "retry completes form option migration"
    CheckMigration colFailures, dAfter("CategoryHashes")("Queries") = dCurrent("Queries"), "retry completes query option migration"
    CheckMigration colFailures, dAfter("CategoryHashes")("_Global") = dCurrent("_Global"), "retry completes global option migration"
    CheckMigration colFailures, OutcomeSucceeded(dResult), "later fast export succeeds publicly"
    CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmA37First.svg"), "later fast export writes first SVG"
    CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmA37Second.svg"), "later fast export writes second SVG"
    CheckMigration colFailures, Operation.Status = eosReady, "root is released"
    CheckMigration colFailures, Operation.DecisionPolicy <> edpAsk, "caller policy survives export"

CleanUp:
    On Error Resume Next
    Operation.SetCancelPollForTest 0
    If blnPolicy Then VCS.ClearOperationPolicy
    If IsLoaded(acForm, "frmVCSMain", False) Then DoCmd.Close acForm, "frmVCSMain", acSaveNo
    Set frmNew = Nothing
    If lngErr <> 0 Then colFailures.Add "ERROR " & lngErr & ": " & strErr
    dReport("success") = (colFailures.Count = 0)
    ExportMigrationCheck = ConvertToJson(dReport)
    Exit Function
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Function

' The layout exporter uses this shared file writer. Verify actual UTF-8 bytes,
' including non-ASCII text and its absence of a BOM, in a disposable temp file.
Public Sub TestLayoutSvgFileWriter()
    Dim strPath As String
    Dim strText As String
    Dim bData() As Byte
    Dim lngErr As Long
    Dim strErr As String
    On Error GoTo ErrHandler
    strPath = GetTempFolder & "A37_" & FSO.GetTempName
    strText = "<svg>" & ChrW$(&H3A9) & "</svg>" & vbCrLf
    WriteFileNoBom strText, strPath
    TestAssert FSO.FileExists(strPath), "SVG writer creates its file"
    bData = GetFileBytes(strPath)
    TestAssert bData(0) = Asc("<"), "UTF-8 output starts with content, without a BOM"
    TestAssert UBound(bData) + 1 = 15, "UTF-8 encodes omega as two bytes"
    TestAssert ReadFile(strPath) = strText, "SVG text survives UTF-8 write/read"
CleanUp:
    On Error Resume Next
    If Len(strPath) > 0 Then
        If FSO.FileExists(strPath) Then FSO.DeleteFile strPath, True
    End If
    If lngErr <> 0 Then TestAssert False, "SVG file writer " & lngErr & ": " & strErr
    Exit Sub
ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub

Private Function MigrationSnapshot() As Dictionary
    Dim strPath As String
    Dim dDump As Dictionary
    ' Reload the persisted binary index, so in-memory state cannot hide a defect.
    Set VCSIndex = Nothing
    strPath = Options.GetExportFolder & "A37-index.json"
    VCSIndex.DumpToJson strPath
    Set dDump = ParseJson(ReadFile(strPath))
    Set MigrationSnapshot = dDump("Items")
    FSO.DeleteFile strPath, True
End Function

Private Sub CheckMigration(colFailures As Collection, blnCondition As Boolean, strContext As String)
    If Not blnCondition Then colFailures.Add strContext
End Sub

Private Function OutcomeSucceeded(dResult As Dictionary) As Boolean
    OutcomeSucceeded = (CStr(dResult("type")) = "complete")
    If dResult.Exists("success") Then OutcomeSucceeded = OutcomeSucceeded And CBool(dResult("success"))
End Function

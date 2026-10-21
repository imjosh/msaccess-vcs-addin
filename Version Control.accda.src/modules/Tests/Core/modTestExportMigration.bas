Attribute VB_Name = "modTestExportMigration"
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")

' Public-entry regression driver. Load the development project as a library in a
' fresh UP008_*.accdb disposable host; never mutate a development or user database.
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
    Dim dAttemptArtifacts As Dictionary
    Dim dRetryArtifacts As Dictionary
    Dim strDecisionPolicy As String

    On Error GoTo ErrHandler
    Set dReport = New Dictionary
    Set colFailures = New Collection
    dReport.Add "case", strCase
    Set dReport("failures") = colFailures
    If Left$(CurrentProject.Name, 6) <> "UP008_" _
        Or CurrentProject.FullName = CodeProject.FullName _
        Or CurrentProject.AllForms.Count <> 0 _
        Or CurrentData.AllQueries.Count <> 0 Then
        colFailures.Add "Requires an empty, disposable UP008_*.accdb host"
        GoTo CleanUp
    End If

    Set frmNew = Application.CreateForm
    strTemp = frmNew.Name
    frmNew.Caption = "UP008 unchanged form"
    DoCmd.Close acForm, strTemp, acSaveYes
    Set frmNew = Nothing
    DoCmd.Rename "frmUP008First", acForm, strTemp
    DoCmd.CopyObject , "frmUP008Second", acForm, "frmUP008First"
    CurrentDb.CreateQueryDef "qryUP008Conflict", "SELECT 1 AS [UP008Value];"

    Set Options = Nothing
    Options.ExportFormatVersion = EFV_5_1_0
    Options.ExportLayoutSvg = False
    Options.UseGitIntegration = False
    Options.SaveOptionsForProject
    VCS.FullExport "block"
    Set dResult = Operation.LastCompletion
    Set dReport("baselineOutcome") = dResult
    CheckMigration colFailures, OutcomeSucceeded(dResult), "baseline full export succeeds"
    strFolder = Options.GetExportFolder
    strQuery = strFolder & "queries" & PathSep & "qryUP008Conflict.sql"
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
    CheckMigration colFailures, Not FSO.FileExists(strFolder & "forms\frmUP008First.svg"), "baseline has no SVG"

    Options.ExportLayoutSvg = True
    If strCase = "partial_cancel" Or strCase = "skip" Then Options.FormatSQL = Not Options.FormatSQL
    If strCase = "global" Or strCase = "filtered" Then Options.ExportFormatVersion = EFV_5_0_0
    Options.SaveOptionsForProject
    strDecisionPolicy = "block"
    Select Case strCase
        Case "blocked", "skip"
            WriteFile "SELECT 2 AS [UP008Value];", strQuery
            SetFileDate strQuery, DateAdd("s", 2, Now), True
            If strCase = "skip" Then strDecisionPolicy = "skip"
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
        Case "warning_complete"
            SetMigrationWarningForTest "complete"
        Case "warning_incomplete"
            SetMigrationWarningForTest "incomplete"
        Case Else
            colFailures.Add "Unknown regression case"
            GoTo CleanUp
    End Select
    If strCase = "filtered" Then
        VCS.ExportVBA strDecisionPolicy
    Else
        VCS.FullExport strDecisionPolicy
    End If
    Operation.SetCancelPollForTest 0
    Set dResult = Operation.LastCompletion
    Set dReport("attemptOutcome") = dResult
    Set dAfter = MigrationSnapshot
    Set dCurrent = Options.GetCategoryHashes
    Set dReport("afterAttempt") = dAfter
    Set dReport("attemptArtifacts") = MigrationArtifacts(strFolder)
    If strCase = "success" Or strCase = "skip" Or strCase = "filtered" _
        Or strCase = "warning_complete" Then
        CheckMigration colFailures, OutcomeSucceeded(dResult), "completed export succeeds"
        CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmUP008First.svg"), "completed form migration writes SVG"
        CheckMigration colFailures, dAfter("CategoryHashes")("Forms") = dCurrent("Forms"), "completed forms record option hash"
        If strCase = "skip" Then
            CheckMigration colFailures, dAfter("CategoryHashes")("Queries") = dBefore("CategoryHashes")("Queries"), "skipped query migration stays pending"
            CheckMigration colFailures, dAfter("FullExportDate") = dBefore("FullExportDate"), "skipped migration does not claim full completion"
        End If
        If strCase = "filtered" Then
            CheckMigration colFailures, dAfter("FullExportDate") = dBefore("FullExportDate"), "VBA export does not claim full project completion"
            CheckMigration colFailures, dAfter("CategoryHashes")("_Global") = dBefore("CategoryHashes")("_Global"), "VBA export leaves global migration pending"
        ElseIf strCase = "success" Or strCase = "warning_complete" Then
            CheckMigration colFailures, dAfter("FullExportDate") <> dBefore("FullExportDate"), "successful full export advances completion date"
            CheckMigration colFailures, dAfter("CategoryHashes")("_Global") = dCurrent("_Global"), "successful full export records global hash"
        End If
        If strCase = "warning_complete" Then
            CheckMigration colFailures, CLng(dResult("error_level")) = eelWarning, "complete output emitted a warning"
        End If
    ElseIf strCase = "warning_incomplete" Then
        CheckMigration colFailures, OutcomeSucceeded(dResult), "incomplete output warning is publicly recorded"
        CheckMigration colFailures, CLng(dResult("error_level")) = eelWarning, "incomplete output emitted a warning"
        CheckMigration colFailures, Not FSO.FileExists(strFolder & "forms\frmUP008First.svg"), "warning left first SVG missing"
        CheckMigration colFailures, dAfter("CategoryHashes")("Forms") = dBefore("CategoryHashes")("Forms"), "missing SVG leaves form migration pending"
        CheckMigration colFailures, dAfter("FullExportDate") = dBefore("FullExportDate"), "missing SVG does not claim full completion"
        CheckMigration colFailures, dAfter("ExportDate") = dBefore("ExportDate"), "missing SVG does not claim export completion"
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
            CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmUP008First.svg"), "partial progress includes first SVG"
            CheckMigration colFailures, Not FSO.FileExists(strFolder & "forms\frmUP008Second.svg"), "second SVG is still pending"
            CheckMigration colFailures, dAfter("Components")("Forms")("frmUP008First.form")("ExportDate") <> _
                dBefore("Components")("Forms")("frmUP008First.form")("ExportDate"), "partial component progress is persisted"
        Else
            CheckMigration colFailures, Not FSO.FileExists(strFolder & "forms\frmUP008First.svg"), "abort did not export unchanged form"
        End If
    End If

    ' Resolve the external conflict by restoring the original source, then fast
    ' export through the same public entry point. Artifacts are the primary oracle.
    If strCase = "blocked" Or strCase = "skip" Then WriteFile strOriginalQuery, strQuery
    ' Separate timestamps enough to detect an unnecessary repeat of completed SVGs.
    Pause 1.2
    VCS.Export "block"
    Set dResult = Operation.LastCompletion
    Set dReport("retryOutcome") = dResult
    Set dAfter = MigrationSnapshot
    Set dReport("afterRetry") = dAfter
    Set dReport("retryArtifacts") = MigrationArtifacts(strFolder)
    If strCase = "warning_complete" Then
        Set dAttemptArtifacts = dReport("attemptArtifacts")
        Set dRetryArtifacts = dReport("retryArtifacts")
        CheckMigration colFailures, _
            dAttemptArtifacts("frmUP008First")("sha256") = dRetryArtifacts("frmUP008First")("sha256"), _
            "complete warning SVG bytes unchanged on retry"
        CheckMigration colFailures, _
            dAttemptArtifacts("frmUP008First")("modified") = dRetryArtifacts("frmUP008First")("modified"), _
            "complete warning SVG was not rewritten on retry"
    End If
    CheckMigration colFailures, dAfter("CategoryHashes")("Forms") = dCurrent("Forms"), "retry completes form option migration"
    CheckMigration colFailures, dAfter("CategoryHashes")("Queries") = dCurrent("Queries"), "retry completes query option migration"
    CheckMigration colFailures, dAfter("CategoryHashes")("_Global") = dCurrent("_Global"), "retry completes global option migration"
    CheckMigration colFailures, OutcomeSucceeded(dResult), "later fast export succeeds publicly"
    CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmUP008First.svg"), "later fast export writes first SVG"
    CheckMigration colFailures, FSO.FileExists(strFolder & "forms\frmUP008Second.svg"), "later fast export writes second SVG"
    CheckMigration colFailures, Operation.Status = eosReady, "root is released"

CleanUp:
    On Error Resume Next
    Operation.SetCancelPollForTest 0
    SetMigrationWarningForTest vbNullString
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

Private Function MigrationArtifacts(strFolder As String) As Dictionary
    Dim dArtifacts As Dictionary
    Dim dFile As Dictionary
    Dim varName As Variant
    Dim strPath As String
    Dim blnShort As Boolean
    Set dArtifacts = New Dictionary
    blnShort = Options.UseShortHash
    Options.UseShortHash = False
    For Each varName In Array("frmUP008First", "frmUP008Second")
        Set dFile = New Dictionary
        strPath = strFolder & "forms" & PathSep & CStr(varName) & ".svg"
        dFile.Add "path", strPath
        dFile.Add "exists", FSO.FileExists(strPath)
        If FSO.FileExists(strPath) Then
            dFile.Add "bytes", FSO.GetFile(strPath).Size
            dFile.Add "sha256", GetFileHash(strPath)
            dFile.Add "modified", CStr(FSO.GetFile(strPath).DateLastModified)
        End If
        dArtifacts.Add CStr(varName), dFile
    Next varName
    Options.UseShortHash = blnShort
    Set MigrationArtifacts = dArtifacts
End Function

Private Function MigrationSnapshot() As Dictionary
    Dim strPath As String
    Dim dDump As Dictionary
    ' Reload the persisted binary index, so in-memory state cannot hide a defect.
    Set VCSIndex = Nothing
    strPath = Options.GetExportFolder & "UP008-index.json"
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

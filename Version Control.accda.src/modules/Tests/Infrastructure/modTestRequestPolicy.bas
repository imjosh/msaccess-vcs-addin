Attribute VB_Name = "modTestRequestPolicy"
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Public Sub TestPublicIdSurvivesCompletionAndIsSeparateFromToken()
    Dim cOp As New clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strId As String
    cOp.Source = eosExternalAPI
    Set cRoot = cOp.TryBeginRoot(eotOther, edpSkip)
    strId = cOp.PublicId
    TestAssert Len(strId) > 0, "direct work generates a public ID"
    TestAssert strId <> cRoot.Token, "public ID is separate from continuation ownership"
    cRoot.Complete eorSuccess
    TestAssert cOp.PublicId = strId, "callback lifetime does not define the public ID"
    TestAssert cOp.LastCompletion("operation_id") = strId, "final result retains the same public ID"
    TestAssert cOp.LastCompletion("decision_policy") = "skip", "completion preserves the effective request policy"
    TestAssert cOp.LastCompletion("initiating_source") = "api", "completion preserves initiating source"
End Sub

Public Sub TestRequestSourceCannotReplaceRootSource()
    Dim cOp As New clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strId As String
    cOp.Source = eosMCPTool
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    strId = cOp.PublicId
    cOp.Source = eosUserInterface
    TestAssert cOp.Source = eosMCPTool, "an active root keeps its initiating source"
    TestAssert cOp.AutomationSource, "source-sensitive behavior uses captured source"
    TestAssert cOp.PublicId = strId, "source refusal preserves public identity"
    cRoot.Complete eorSuccess
End Sub

Public Sub TestPublicIdsNeverReuseAcceptedCallerIdentity()
    Dim cFirst As clsOperationRequest
    Dim cDuplicate As clsOperationRequest
    Dim strCallback As String
    Dim strRefusal As String
    strCallback = "{""operation_id"":""" & NewOperationIdentity() & """,""callback_url"":""http://127.0.0.1:1/callback""}"
    Set cFirst = CreateOperationRequest(eosMCPTool, strCallback)
    strRefusal = ClaimOperationRequest(cFirst)
    TestAssert Len(strRefusal) = 0, "a fresh caller identity can be admitted"
    Set cDuplicate = CreateOperationRequest(eosMCPTool, strCallback)
    strRefusal = ClaimOperationRequest(cDuplicate)
    TestAssert InStr(strRefusal, "operation_id_reused") > 0, "a new request cannot reuse an admitted identity"
End Sub

Public Sub TestTypedArgumentsRoundTripThroughQueuedJson()
    Dim colArgs As New Collection
    Dim colRestored As Collection
    colArgs.Add ""
    colArgs.Add False
    colArgs.Add "block|with-pipe"
    Set colRestored = ParseJson(ConvertToJson(colArgs))
    TestAssert colRestored.Count = 3, "all positions are retained"
    TestAssert VarType(colRestored(2)) = vbBoolean, "booleans keep their type"
    TestAssert colRestored(1) = "", "an empty argument keeps its position"
    TestAssert colRestored(3) = "block|with-pipe", "pipes are ordinary argument content"
End Sub

Public Sub TestCompletionKeepsFinalErrorLevelAfterScopeRelease()
    Dim cOp As New clsOperation
    Dim cRoot As clsRootOperationLease
    Set cRoot = cOp.TryBeginRoot(eotExport, edpBlock)
    cOp.ErrorLevel = eelError
    cRoot.Complete eorFailed
    TestAssert cOp.ErrorLevel = eelNoError, "scope restoration releases temporary state"
    TestAssert cOp.LastCompletion("error_level") = eelError, "final diagnostics survive restoration"
    TestAssert Not cOp.LastCompletion("success"), "the final result remains unsuccessful"
End Sub

Public Sub TestExpiryCompletesOnlyItsOwnRequest()
    Dim cOp As New clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strId As String
    Set cRoot = cOp.TryBeginRoot(eotOther, edpBlock)
    strId = cOp.PublicId
    cOp.Heartbeat = DateAdd("n", -11, Now)
    TestAssert cOp.Status = eosReady, "already-due heartbeat expiry releases ownership"
    TestAssert cOp.Result = eorTimeout, "expiry records a timeout"
    TestAssert cOp.LastCompletion("operation_id") = strId, "expiry identifies the expired request"
    TestAssert cOp.InteractionMode = eimNormal, "expiry releases the request policy mode"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorTimeout, "a stale lease cannot replace the expired result"
End Sub

Public Function RequestPolicyProbe(ByVal strCase As String, Optional ByVal strCallbackInfo As String) As String
    Dim cRoot As clsRootOperationLease
    Dim cPrior As clsOperationRequest
    Dim cRequest As clsOperationRequest
    Dim cRecreated As clsOperation
    Dim cEmitter As clsMCP
    Dim d As New Dictionary
    Dim strToken As String
    Dim strQueue As String
    Dim strError As String
    Dim lngError As Long
    Dim cVcs As clsVersionControl
    Dim cLauncher As clsRebuildLauncherProbe
    Dim dParsed As Dictionary
    Dim strOldId As String
    Dim datDeadline As Date
    Dim lngResetBefore As Long
    Dim blnFixtureExit As Boolean
    If Operation.IsActive Or TimerIsPending Then
        RequestPolicyProbe = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
        Exit Function
    End If
    Set cPrior = IncomingOperationRequest
    On Error GoTo Failed
    Select Case strCase
        Case "retired"
            d("setter") = VCS.SetOperationPolicy("block")
            d("mode") = VCS.SetInteractionMode(2)
            d("clear") = VCS.ClearOperationPolicy()
            Set d("after") = ParseJson(ModeOnlyScopeProbe("snapshot"))
        Case "active_isolation", "recreation", "stale_continuation", "cancel_target"
            Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
            Set IncomingOperationRequest = cRequest
            Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
            If cRoot Is Nothing Then Err.Raise vbObjectError + 2450, "RequestPolicyProbe", Operation.AdmissionRefusal
            Set cEmitter = MCP
            Set d("before") = ParseJson(ModeOnlyScopeProbe("snapshot"))
            Select Case strCase
                Case "active_isolation"
                    d("inspection") = API("ListAddinDialogs")
                    d("setter") = API("SetOperationPolicy", "decline")
                    d("mode") = API("SetInteractionMode", 2)
                    d("test") = API("RunFilteredTests", "skip")
                    d("async") = APIAsync(strCallbackInfo, "Export", "skip")
                    d("manual") = HandleRibbonCommand("btnShowOptions")
                    d("dynamic") = RunInAddIn("RequestPolicyNoop", True)
                Case "cancel_target"
                    d("missing") = API("DismissAddinDialog", "frmVCSMain", "cancel")
                    d("mismatch") = API("DismissAddinDialog", "frmVCSMain", "cancel", "foreign-operation")
                Case "stale_continuation"
                    RunBuildFromContinuation "foreign-token", "C:\missing\", False
                Case "recreation"
                    strToken = cRoot.DetachForContinuation()
                    Set cRecreated = New clsOperation
                    d("restored") = cRecreated.RestorePersistedRootForTest()
                    d("operation_id") = cRecreated.PublicId
                    d("policy") = cRecreated.PolicyName
                    d("source") = cRecreated.SourceLabel
                    d("token") = cRecreated.CurrentRootToken
                    d("attended") = cRecreated.Attended
                    d("resume_wrong") = Not cRecreated.ResumeRoot("foreign-token") Is Nothing
                    Operation.CompleteRootLease strToken, eorSuccess
            End Select
            Set d("after") = ParseJson(ModeOnlyScopeProbe("snapshot"))
            d("attempts_before_completion") = cEmitter.GetDiagnosticData()("terminalPostAttempts")
            If Not cRoot Is Nothing Then cRoot.Complete eorSuccess
            d("attempts_after_completion") = cEmitter.GetDiagnosticData()("terminalPostAttempts")
        Case "timer_refusal"
            strQueue = SetTimer("RequestPolicyNoop", "first", "second", blnReserveOnly:=True)
            d("queue") = strQueue
            Set d("before") = RequestTimerSnapshot()
            d("refusal") = SetTimer("RequestPolicyNoop", "foreign", "changed")
            d("async_refusal") = APIAsync(strCallbackInfo, "Export", "block")
            Set d("after") = RequestTimerSnapshot()
            DispatchReservedTimerForTest
        Case "stale_timer_event"
            strQueue = SetTimer("RequestPolicyNoop", "owned", "arguments", sngSeconds:=60)
            d("queue") = strQueue
            Set d("before") = RequestTimerSnapshot()
            WinAPITimerCallback 0, 0, -1, 0
            Set d("after") = RequestTimerSnapshot()
            DispatchReservedTimerForTest
        Case "failed_timer"
            FailNextTimerForTest
            d("queue") = SetTimer("RequestPolicyNoop")
            d("pending") = TimerIsPending
        Case "queue_recheck"
            d("queue") = APIAsync(strCallbackInfo, "Export", "block")
            Set cRequest = CreateOperationRequest(eosExternalAPI)
            Set IncomingOperationRequest = cRequest
            Set cRoot = Operation.TryBeginRoot(eotOther, edpSkip)
            If cRoot Is Nothing Then Err.Raise vbObjectError + 2450, "RequestPolicyProbe", Operation.AdmissionRefusal
            Set d("before") = ParseJson(ModeOnlyScopeProbe("snapshot"))
            DispatchReservedTimerForTest
            Set d("after") = ParseJson(ModeOnlyScopeProbe("snapshot"))
            cRoot.Complete eorSuccess
        Case "dynamic_failure"
            FailNextTimerReservationForTest
            d("result") = RunInAddIn("RequestPolicyNoop", True)
            d("pending") = TimerIsPending
        Case "inspect_expiry"
            Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
            Set IncomingOperationRequest = cRequest
            Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
            Operation.Heartbeat = DateAdd("n", -11, Now)
            d("inspection") = API("ListAddinDialogs")
            d("result") = Operation.Result
            Set d("completion") = Operation.LastCompletion
            d("active") = Operation.IsActive
        Case "disconnect"
            Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
            Set IncomingOperationRequest = cRequest
            Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
            strOldId = Operation.PublicId
            MCP.RegisterCallback vbNullString
            d("owned_policy_after_disconnect") = Operation.PolicyName
            cRoot.Complete eorSuccess
            d("completed_id") = Operation.PublicId
            d("expected_id") = strOldId
            Set IncomingOperationRequest = Nothing
            Set cRoot = Operation.TryBeginRoot(eotOther)
            d("next_policy") = Operation.PolicyName
            d("next_source") = Operation.SourceLabel
            d("next_mode") = Operation.InteractionMode
            cRoot.Complete eorSuccess
        Case "cancel_reuse"
            Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
            Set IncomingOperationRequest = cRequest
            Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
            strOldId = Operation.PublicId
            cRoot.Complete eorSuccess
            Set cRequest = CreateOperationRequest(eosExternalAPI)
            Set IncomingOperationRequest = cRequest
            Set cRoot = Operation.TryBeginRoot(eotOther, edpSkip)
            Set d("before") = ParseJson(ModeOnlyScopeProbe("snapshot"))
            d("stale_cancel") = API("DismissAddinDialog", "frmVCSMain", "cancel", strOldId)
            Set d("after_stale") = ParseJson(ModeOnlyScopeProbe("snapshot"))
            d("matching_cancel") = API("DismissAddinDialog", "frmVCSMain", "cancel", Operation.PublicId)
            d("matching_result") = Operation.Result
            cRoot.Complete eorCanceled
        Case "ribbon_queue_failure"
            FailNamedTimerForTest "Build"
            d("accepted") = HandleRibbonCommand("btnBuild", Options.GetExportFolder)
            d("active") = Operation.IsActive
            d("pending") = TimerIsPending
            Set d("completion") = Operation.LastCompletion
        Case "failed_build_queue", "failed_merge_reset", "failed_merge_resume"
            Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
            Set IncomingOperationRequest = cRequest
            lngResetBefore = ProjectResetCallsForTest
            If strCase = "failed_build_queue" Then
                FailNamedTimerForTest "Build"
                d("admission") = VCS.Build(Options.GetExportFolder, "block")
            Else
                If strCase = "failed_merge_reset" Then
                    FailNamedTimerForTest "MergeReset"
                Else
                    FailNamedTimerForTest "MergeResume"
                End If
                d("admission") = VCS.MergeBuild("block")
            End If
            datDeadline = DateAdd("s", 30, Now)
            Do While Operation.IsActive And Now < datDeadline
                DoEvents
                Pause 0.05
            Loop
            d("active") = Operation.IsActive
            d("pending") = TimerIsPending
            d("reset_calls") = ProjectResetCallsForTest - lngResetBefore
            d("result") = Operation.Result
            Set d("completion") = Operation.LastCompletion
        Case "rebuild_reservation", "rebuild_activation", "rebuild_child_identity"
            Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
            Set IncomingOperationRequest = cRequest
            Set cLauncher = New clsRebuildLauncherProbe
            Set cVcs = New clsVersionControl
            Set cVcs.RebuildLauncherForTest = cLauncher
            If strCase = "rebuild_reservation" Then
                FailNextTimerReservationForTest
            ElseIf strCase = "rebuild_activation" Then
                FailNextTimerForTest
            End If
            Set dParsed = ParseJson(strCallbackInfo)
            If strCase <> "rebuild_child_identity" Then dParsed("operation_id") = NewOperationIdentity()
            d("expected_build_operation_id") = dParsed("operation_id")
            d("outcome") = cVcs.RebuildAddIn(Options.GetExportFolder, ConvertToJson(dParsed))
            d("launch_count") = cLauncher.LaunchCount
            d("fixture_helper_pid") = cLauncher.ProcessId
            d("fixture_helper_creation_time") = cLauncher.CreationTime
            d("reservation_held") = Operation.IsActive
            If cLauncher.LaunchCount > 0 Then
                Operation.Heartbeat = DateAdd("n", -11, Now)
                Operation.RetainRebuildHandoff Operation.CurrentRootToken
                d("active_after_due_heartbeat") = Operation.IsActive
                d("handoff_hold") = Operation.RebuildHandoffHeld
                d("timer_reservation_after_due_heartbeat") = GetSetting(PROJECT_NAME, "Timer", "Reservation")
                Set cRecreated = New clsOperation
                d("held_root_restored") = cRecreated.RestorePersistedRootForTest()
                d("restored_handoff_hold") = cRecreated.RebuildHandoffHeld
                d("restored_active_after_due_heartbeat") = cRecreated.IsActive
                blnFixtureExit = cLauncher.ConfirmExit()
                d("fixture_helper_original_handle_exit") = blnFixtureExit
                If blnFixtureExit Then
                    ' Only this inert fixture can establish that its helper can no
                    ' longer act. Real rebuild recovery remains a separate contract.
                    If Operation.PublicId = cRequest.PublicId Then
                        strToken = Operation.CurrentRootToken
                        ReleaseTimerReservation strToken
                        Operation.CompleteRootLease strToken, eorFailed
                        d("fixture_only_reservation_released") = True
                    End If
                End If
            End If
            d("active_after_fixture_cleanup") = Operation.IsActive
            d("pending_after_fixture_cleanup") = TimerIsPending
        Case Else
            d("error") = "Unknown request-policy probe case"
    End Select
CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    FailNamedTimerForTest vbNullString
    Set IncomingOperationRequest = cPrior
    If lngError <> 0 Then
        d("probe_error") = strError
        d("probe_error_number") = lngError
    End If
    RequestPolicyProbe = ConvertToJson(d)
    Exit Function
Failed:
    strError = Err.Description
    lngError = Err.Number
    Resume CleanUp
End Function

Public Function RequestTimerSnapshot() As Dictionary
    Dim d As New Dictionary
    Dim varKey As Variant
    For Each varKey In Array("Operation", "Param1", "Param2", "OpToken", "CallbackInfo", "RequestJson", "Reservation", "TimerID")
        d(CStr(varKey)) = GetSetting(PROJECT_NAME, "Timer", CStr(varKey))
    Next varKey
    Set RequestTimerSnapshot = d
End Function

Public Sub RequestPolicyNoop(Optional ByVal strFirst As String, Optional ByVal strSecond As String)
End Sub

Public Function RequestPolicyStatementProbe() As String
    Dim d As New Dictionary
    VCS.Export "block"
    Set d("direct") = Operation.LastCompletion
    API "Export", "block"
    Set d("api") = Operation.LastCompletion
    FailNamedTimerForTest "Build"
    VCS.Build Options.GetExportFolder, "block"
    Set d("direct_build") = Operation.LastCompletion
    FailNamedTimerForTest "Build"
    API "Build", Options.GetExportFolder, "block"
    Set d("api_build") = Operation.LastCompletion
    d("active") = Operation.IsActive
    d("pending") = TimerIsPending
    RequestPolicyStatementProbe = ConvertToJson(d)
End Function

Public Function ClaimRequestIdentityProbe(ByVal strOperationId As String) As String
    Dim cRequest As clsOperationRequest
    Dim strRefusal As String
    Set cRequest = CreateOperationRequest(eosMCPTool)
    cRequest.PublicId = strOperationId
    strRefusal = ClaimOperationRequest(cRequest)
    If Len(strRefusal) > 0 Then
        ClaimRequestIdentityProbe = strRefusal
    Else
        ClaimRequestIdentityProbe = "{""success"":true}"
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : RequestPolicyF4Probe
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Exercise actual form preflight in a disposable user database. An owned
'           : request supplies its policy and token; retired activation is never used.
'---------------------------------------------------------------------------------------
'
Public Function RequestPolicyF4Probe(ByVal strCase As String, ByVal strCallbackInfo As String, _
    Optional ByVal strPolicy As String = "block") As String

    Dim d As New Dictionary
    Dim cPrior As clsOperationRequest
    Dim cRequest As clsOperationRequest
    Dim cRoot As clsRootOperationLease
    Dim cEmitter As clsMCP
    Dim intPolicy As eDecisionPolicy
    Dim strRefusal As String
    Dim strToken As String
    Dim strError As String
    Dim lngError As Long
    Dim blnHadForm As Boolean
    Dim blnFormCaptured As Boolean
    Dim intPriorPolicy As eDecisionPolicy
    Dim blnPriorFullBuild As Boolean
    Dim strPriorSource As String
    Dim strPriorToken As String
    Dim strPriorResult As String

    Set cPrior = IncomingOperationRequest
    InitializeRequestErrors "RequestPolicyF4Probe"
    On Error GoTo Failed
    If strCase = "inspect" Then
        Set d = F4OwnerSnapshot()
        GoTo CleanUp
    End If
    If Operation.IsActive Or TimerIsPending Then
        RequestPolicyF4Probe = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
        Exit Function
    End If
    strRefusal = ParsePolicyOrRefuse(strPolicy, intPolicy)
    If Len(strRefusal) > 0 Then Err.Raise vbObjectError + 2470, "RequestPolicyF4Probe", strRefusal
    If strCase = "previous_journal" Then
        Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
        Set IncomingOperationRequest = cRequest
        Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
        If cRoot Is Nothing Then Err.Raise vbObjectError + 2471, "RequestPolicyF4Probe", Operation.AdmissionRefusal
        Operation.RecordDecision NewDecision(DECISION_REQUIRED, "Previous F4 operation", _
            "This prior operation owns a different decision.", DECISION_BLOCKED)
        cRoot.Complete eorSuccess
        Set d("completion") = Operation.LastCompletion
        Set d("after") = F4OwnerSnapshot()
        GoTo CleanUp
    End If

    blnHadForm = IsLoaded(acForm, "frmVCSMain", False)
    If strCase = "active_form" Then
        ' Configure the attempted form request before the competing owner starts.
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    ElseIf strCase <> "full_build_confirmation" Then
        Err.Raise vbObjectError + 2472, "RequestPolicyF4Probe", "Unknown F4 probe case"
    End If

    If strCase = "full_build_confirmation" Then
        Set cRequest = CreateOperationRequest(eosUserInterface, strCallbackInfo)
        Set IncomingOperationRequest = cRequest
        Set cRoot = Operation.TryBeginRoot(eotBuild, intPolicy)
        If cRoot Is Nothing Then Err.Raise vbObjectError + 2471, "RequestPolicyF4Probe", Operation.AdmissionRefusal
        strToken = cRoot.DetachForContinuation()
        ' Match the public Build wrapper's acquisition-before-Form_Load handoff.
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    End If

    With Form_frmVCSMain
        intPriorPolicy = .intDecisionPolicy
        blnPriorFullBuild = .chkFullBuild
        strPriorSource = .strSourcePath
        strPriorToken = .strRequestRootToken
        strPriorResult = .strStartResult
        blnFormCaptured = True
        .chkFullBuild = True
        .strSourcePath = vbNullString
        .intDecisionPolicy = intPolicy
        .strRequestRootToken = vbNullString
        If strCase = "active_form" Then
            Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
            Set IncomingOperationRequest = cRequest
            Set cRoot = Operation.TryBeginRoot(eotOther, edpBlock)
            If cRoot Is Nothing Then Err.Raise vbObjectError + 2471, "RequestPolicyF4Probe", Operation.AdmissionRefusal
            Operation.RecordDecision NewDecision(DECISION_REQUIRED, "Active F4 owner", _
                "The refused requests must not replace this owner's journal.", DECISION_BLOCKED)
            Set cEmitter = MCP
            Set d("before") = F4OwnerSnapshot()
            Set IncomingOperationRequest = CreateOperationRequest(eosUserInterface)
            d("public_refusal") = VCS.MergeBuild("decline")
            d("ribbon_accepted") = HandleRibbonCommand("btnBuild", Options.GetExportFolder)
            .cmdBuild_Click
            Set d("after") = F4OwnerSnapshot()
            d("attempts_before_completion") = cEmitter.GetDiagnosticData()("terminalPostAttempts")
            cRoot.Complete eorSuccess
            d("attempts_after_completion") = cEmitter.GetDiagnosticData()("terminalPostAttempts")
        Else
            .strRequestRootToken = strToken
            d("source_before_form_entry") = Operation.SourceLabel
            d("policy_before_form_entry") = Operation.PolicyName
            d("valid_source_options") = FolderHasVcsOptionsFile(Options.GetExportFolder)
            ' If a regression bypasses the confirmation, this control stops the real
            ' timer handoff. A 2410 result fails qualification; it is not coverage.
            FailNamedTimerForTest "Build"
            .cmdBuild_Click
            d("outcome") = .strStartResult
            Set d("completion") = Operation.LastCompletion
            Set d("after") = F4OwnerSnapshot()
        End If
    End With

CleanUp:
    On Error Resume Next
    If Not cRoot Is Nothing Then cRoot.Complete eorFailed
    FailNamedTimerForTest vbNullString
    If blnFormCaptured And IsLoaded(acForm, "frmVCSMain", False) Then
        With Form_frmVCSMain
            .intDecisionPolicy = intPriorPolicy
            .chkFullBuild = blnPriorFullBuild
            .strSourcePath = strPriorSource
            .strRequestRootToken = strPriorToken
            .strStartResult = strPriorResult
        End With
        If Not blnHadForm Then DoCmd.Close acForm, "frmVCSMain", acSaveNo
    End If
    Set IncomingOperationRequest = cPrior
    If lngError <> 0 Then
        d("probe_error") = strError
        d("probe_error_number") = lngError
    End If
    RequestPolicyF4Probe = ConvertToJson(d)
    Exit Function
Failed:
    strError = Err.Description
    lngError = Err.Number
    Resume CleanUp
End Function

Private Function F4OwnerSnapshot() As Dictionary
    Dim d As Dictionary
    Set d = ParseJson(ModeOnlyScopeProbe("snapshot"))
    d("active") = Operation.IsActive
    d("pending") = TimerIsPending
    d("host") = CurrentProject.FullName
    Set d("timer") = RequestTimerSnapshot()
    If IsLoaded(acForm, "frmVCSMain", False) Then
        With Form_frmVCSMain
            d("form_policy") = .intDecisionPolicy
            d("form_full_build") = CBool(.chkFullBuild)
            d("form_source_path") = .strSourcePath
            d("form_request_token") = .strRequestRootToken
            d("form_start_result") = .strStartResult
            d("form_merge_enabled") = .chkFullBuild.Enabled
        End With
    End If
    Set F4OwnerSnapshot = d
End Function

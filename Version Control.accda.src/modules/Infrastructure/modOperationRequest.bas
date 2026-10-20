Attribute VB_Name = "modOperationRequest"
'---------------------------------------------------------------------------------------
' Module    : modOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Admits request identities and isolates request diagnostics from active work.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Infrastructure")

Private Type udtRequestGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef pguid As udtRequestGuid) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef rguid As udtRequestGuid, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long

Private Declare PtrSafe Function CreateMutexW Lib "kernel32" (ByVal lpMutexAttributes As LongPtr, ByVal bInitialOwner As Long, ByVal lpName As LongPtr) As LongPtr
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function ReleaseMutex Lib "kernel32" (ByVal hMutex As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetUserNameExW Lib "secur32" (ByVal NameFormat As Long, ByVal lpNameBuffer As LongPtr, ByRef nSize As Long) As Long
Private Const NAME_SAM_COMPATIBLE As Long = 2
Private Const WAIT_OBJECT_0 As Long = 0
Private Const WAIT_ABANDONED As Long = &H80
Private Const IDENTITY_LOCK_TIMEOUT_MS As Long = 5000

Private m_Request As clsOperationRequest
Private m_blnDispatching As Boolean
Private m_PrivateClaims As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : NewOperationIdentity
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Create a random identity distinct from internal continuation ownership.
'---------------------------------------------------------------------------------------
'
Public Function NewOperationIdentity() As String
    Dim guid As udtRequestGuid
    Dim strGuid As String
    InitializeRequestErrors "NewOperationIdentity"
    On Error GoTo Failed
    If CoCreateGuid(guid) <> 0 Then Err.Raise vbObjectError + 2401, "NewOperationIdentity", T("Could not generate an operation identity.")
    strGuid = String$(39, vbNullChar)
    If StringFromGUID2(guid, StrPtr(strGuid), 39) = 0 Then Err.Raise vbObjectError + 2402, "NewOperationIdentity", T("Could not format an operation identity.")
    NewOperationIdentity = Mid$(strGuid, 2, 36)
    Exit Function
Failed:
    Err.Raise Err.Number, "modOperationRequest.NewOperationIdentity", Err.Description
End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Capture incoming source and callback identity before operation admission.
'---------------------------------------------------------------------------------------
'
Public Function CreateOperationRequest(ByVal intSource As eOperationSource, Optional ByVal strCallback As String) As clsOperationRequest
    Dim c As New clsOperationRequest
    Dim d As Dictionary
    InitializeRequestErrors "CreateOperationRequest"
    On Error GoTo Failed
    c.Source = intSource
    c.CallbackInfo = strCallback
    If Len(strCallback) > 0 Then
        Set d = ParseJson(strCallback)
        If Not d.Exists("operation_id") Or Not d.Exists("callback_url") Then
            Err.Raise vbObjectError + 2403, "CreateOperationRequest", T("Callback information must include operation_id and callback_url.")
        End If
        c.PublicId = CStr(d("operation_id"))
        If Len(c.PublicId) = 0 Or Len(CStr(d("callback_url"))) = 0 Then
            Err.Raise vbObjectError + 2403, "CreateOperationRequest", T("Callback information must include operation_id and callback_url.")
        End If
    Else
        c.PublicId = NewOperationIdentity()
    End If
    c.ReservationToken = NewOperationIdentity()
    Set CreateOperationRequest = c
    Exit Function
Failed:
    Err.Raise Err.Number, "modOperationRequest.CreateOperationRequest", Err.Description
End Function


'---------------------------------------------------------------------------------------
' Procedure : IncomingOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Expose incoming request context without replacing the active owner.
'---------------------------------------------------------------------------------------
'
Public Property Get IncomingOperationRequest() As clsOperationRequest
    Set IncomingOperationRequest = m_Request
End Property


'---------------------------------------------------------------------------------------
' Procedure : IncomingOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Read or restore incoming request context separately from active ownership.
'---------------------------------------------------------------------------------------
'
Public Property Set IncomingOperationRequest(ByVal c As clsOperationRequest)
    Set m_Request = c
End Property


'---------------------------------------------------------------------------------------
' Procedure : RequestSource
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Read incoming caller source before a root has captured it.
'---------------------------------------------------------------------------------------
'
Public Property Get RequestSource() As eOperationSource
    If m_Request Is Nothing Then
        RequestSource = eosUserInterface
    Else
        RequestSource = m_Request.Source
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : APIDispatchBusy
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Track the dispatch stack separately from active root ownership.
'---------------------------------------------------------------------------------------
'
Public Property Get APIDispatchBusy() As Boolean
    APIDispatchBusy = m_blnDispatching
End Property


'---------------------------------------------------------------------------------------
' Procedure : APIDispatchBusy
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Read or update the dispatch-stack guard independently of root state.
'---------------------------------------------------------------------------------------
'
Public Property Let APIDispatchBusy(ByVal blnBusy As Boolean)
    m_blnDispatching = blnBusy
End Property

' Exact, case-sensitive UTF-16 value names avoid hashing/case-folding collisions.
' Accepted IDs are at most 128 UTF-16 code units. Admitted IDs are never reused,
' including after completion, callback release, VBA reset, or Access restart.

'---------------------------------------------------------------------------------------
' Procedure : IdentityKey
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Encode exact UTF-16 identity units into case-insensitive registry-safe keys.
'---------------------------------------------------------------------------------------
'
Private Function IdentityKey(ByVal strId As String) As String
    Dim lngIdx As Long
    For lngIdx = 1 To Len(strId)
        IdentityKey = IdentityKey & Right$("0000" & Hex$(AscW(Mid$(strId, lngIdx, 1)) And &HFFFF&), 4)
    Next lngIdx
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClaimOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Refuse reused IDs and claim a fresh or matching queued request.
'---------------------------------------------------------------------------------------
'
Public Function ClaimOperationRequest(ByVal c As clsOperationRequest, Optional ByVal blnPersistent As Boolean = True) As String
    Dim strKey As String
    InitializeRequestErrors "ClaimOperationRequest"
    On Error GoTo Failed
    If c.Consumed Then
        ClaimOperationRequest = RefusalJson("operation_id_reused", T("This request already owns an operation."), False)
        Exit Function
    End If
    If Len(c.PublicId) = 0 Or Len(c.PublicId) > 128 Or InStr(c.PublicId, vbNullChar) > 0 Then
        ClaimOperationRequest = RefusalJson("invalid_operation_id", T("Operation IDs must contain 1 to 128 characters and no null characters."), False)
        Exit Function
    End If
    strKey = IdentityKey(c.PublicId)
    If Not blnPersistent Then
        If m_PrivateClaims Is Nothing Then Set m_PrivateClaims = New Dictionary
        If Not m_PrivateClaims.Exists(strKey) Then
            m_PrivateClaims.Add strKey, c.ReservationToken
            c.Claimed = True
            Exit Function
        ElseIf c.Claimed And m_PrivateClaims(strKey) = c.ReservationToken Then
            Exit Function
        End If
        ClaimOperationRequest = RefusalJson("operation_id_reused", T("That operation ID has already been admitted. Use a new operation ID."), False)
        Exit Function
    End If
    ClaimOperationRequest = ClaimPersistentOperationRequest(c, strKey)

    Exit Function
Failed:
    ClaimOperationRequest = RefusalJson("operation_id_store_unavailable", _
        T("The operation identity could not be reserved. No operation started. {0}", var0:=Err.Description), False)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RequestAdmissionRefusal
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Refuse new work before changing active owner state.
'---------------------------------------------------------------------------------------
'
Public Function RequestAdmissionRefusal(Optional ByVal blnAllowCurrentTimer As Boolean = False, _
    Optional ByVal blnQuiet As Boolean = False) As String
    If Operation.IsActive Then
        RequestAdmissionRefusal = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
        If blnQuiet Or Operation.AutomationSource Or RequestSource <> eosUserInterface Then
            SysCmd acSysCmdSetStatus, OperationRunningMessage()
        Else
            MsgBox2 T("Another Operation Already Running"), OperationRunningMessage(), , vbExclamation
        End If
    ElseIf Not blnAllowCurrentTimer And TimerIsPending Then
        RequestAdmissionRefusal = RefusalJson("timer_busy", T("Another request is waiting for its callback timer."), False)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ManualCommandRefused
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Guard manual entry and report busy state without automated modal UI.
'---------------------------------------------------------------------------------------
'
Public Function ManualCommandRefused() As Boolean
    ManualCommandRefused = Operation.IsActive Or TimerIsPending
    If Not ManualCommandRefused Then Exit Function
    If Operation.AutomationSource Or APIDispatchBusy Or TimerIsPending Then
        SysCmd acSysCmdSetStatus, OperationRunningMessage()
    Else
        MsgBox2 T("Another Operation Already Running"), OperationRunningMessage(), , vbExclamation
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : OperationPolicyName
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Format the captured policy for public result metadata.
'---------------------------------------------------------------------------------------
'
Public Function OperationPolicyName(ByVal intPolicy As eDecisionPolicy) As String
    Select Case intPolicy
        Case edpBlock: OperationPolicyName = "block"
        Case edpDecline: OperationPolicyName = "decline"
        Case edpPreferSource: OperationPolicyName = "prefer_source"
        Case edpPreferDatabase: OperationPolicyName = "prefer_database"
        Case edpSkip: OperationPolicyName = "skip"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddOperationIdentity
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Copy stable owner identity into a request result.
'---------------------------------------------------------------------------------------
'
Public Sub AddOperationIdentity(ByVal d As Object, ByVal c As clsOperation)
    d("operation_id") = c.PublicId
    d("initiating_source") = c.SourceLabel
    d("decision_policy") = c.PolicyName
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompletedOperationJson
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Serialize the retained final result of an admitted operation.
'---------------------------------------------------------------------------------------
'
Public Function CompletedOperationJson(ByVal c As clsOperation) As String
    Dim d As Dictionary
    InitializeRequestErrors "CompletedOperationJson"
    On Error GoTo Failed
    Set d = c.LastCompletion
    If d Is Nothing Then
        CompletedOperationJson = RefusalJson("operation_not_started", T("The operation did not start."), False)
    Else
        CompletedOperationJson = ConvertToJson(d)
    End If
    Exit Function
Failed:
    Err.Raise Err.Number, "modOperationRequest.CompletedOperationJson", Err.Description
End Function


'---------------------------------------------------------------------------------------
' Procedure : LogOperationIdentity
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Identify the admitted owner in its own operation log.
'---------------------------------------------------------------------------------------
'
Public Sub LogOperationIdentity()
    Log.Add T("Operation ID: {0}", var0:=Operation.PublicId), False
    Log.Add T("Decision policy: {0}", var0:=IIf(Len(Operation.PolicyName) > 0, Operation.PolicyName, T("none"))), False
    Log.Add T("Initiating source: {0}", var0:=Operation.SourceLabel), False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OverlayOperationIdentity
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Preserve an existing result while adding its admitted identity.
'---------------------------------------------------------------------------------------
'
Public Function OverlayOperationIdentity(ByVal strJson As String, ByVal cOp As clsOperation) As String
    Dim d As Object
    OverlayOperationIdentity = strJson
    If Len(strJson) = 0 Then Exit Function
    InitializeRequestErrors "OverlayOperationIdentity"
    On Error GoTo Unchanged
    Set d = ParseJson(strJson)
    AddOperationIdentity d, cOp
    If Not cOp.LastCompletion Is Nothing Then d("error_level") = cOp.LastCompletion("error_level")
    OverlayOperationIdentity = ConvertToJson(d)
Unchanged:
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClaimPersistentOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Claim a never-reused ID under the cross-process per-user mutex.
'---------------------------------------------------------------------------------------
'
Private Function ClaimPersistentOperationRequest(ByVal c As clsOperationRequest, ByVal strKey As String) As String
    Dim hMutex As LongPtr
    Dim blnOwned As Boolean
    Dim lngWait As Long
    Dim strUser As String
    Dim lngUserLength As Long
    Dim strMutexName As String
    Dim strStored As String
    Dim strError As String
    InitializeRequestErrors "ClaimPersistentOperationRequest"
    On Error GoTo StoreFailed
    strUser = String$(1024, vbNullChar)
    lngUserLength = Len(strUser)
    If GetUserNameExW(NAME_SAM_COMPATIBLE, StrPtr(strUser), lngUserLength) = 0 Then GoTo StoreFailed
    strUser = Left$(strUser, lngUserLength)
    strMutexName = "Global\MSAccessVCS.OperationIds." & Replace(strUser, "\", "_")
    hMutex = CreateMutexW(0, 0, StrPtr(strMutexName))
    If hMutex = 0 Then GoTo StoreFailed
    lngWait = WaitForSingleObject(hMutex, IDENTITY_LOCK_TIMEOUT_MS)
    blnOwned = (lngWait = WAIT_OBJECT_0 Or lngWait = WAIT_ABANDONED)
    If Not blnOwned Then GoTo StoreFailed
    strStored = GetSetting(PROJECT_NAME, "OperationIds", strKey)
    If c.Claimed Then
        If strStored = c.ReservationToken Then GoTo CleanUp
    ElseIf Len(strStored) = 0 Then
        SaveSetting PROJECT_NAME, "OperationIds", strKey, c.ReservationToken
        If GetSetting(PROJECT_NAME, "OperationIds", strKey) <> c.ReservationToken Then GoTo StoreFailed
        c.Claimed = True
        GoTo CleanUp
    End If
    ClaimPersistentOperationRequest = RefusalJson("operation_id_reused", T("That operation ID has already been admitted. Use a new operation ID."), False)
CleanUp:
    InitializeRequestErrors "ClaimPersistentOperationRequest.CleanUp"
    On Error Resume Next
    If blnOwned Then ReleaseMutex hMutex
    If hMutex <> 0 Then CloseHandle hMutex
    Exit Function
StoreFailed:
    strError = Err.Description
    ClaimPersistentOperationRequest = RefusalJson("operation_id_store_unavailable", T("The operation identity could not be reserved. No operation started. {0}", var0:=strError), False)
    GoTo CleanUp
End Function

' Request admission precedes ownership. Global error logging here would append to
' another request's journal/log. Preserve an inherited diagnostic without loading
' Log, then run the standard debug check before a request-local On Error handler.

'---------------------------------------------------------------------------------------
' Procedure : InitializeRequestErrors
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Keep inherited diagnostics outside another root before standard debug setup.
'---------------------------------------------------------------------------------------
'
Public Sub InitializeRequestErrors(ByVal strProcedure As String)
    If Err.Number <> 0 Then
        Debug.Print strProcedure & ": inherited error " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    DebugMode True
End Sub

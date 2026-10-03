Attribute VB_Name = "modCompatibilitySession"
Option Compare Database
Option Explicit
'@Folder("API")

' X17: admission never reads or changes Operation, MCP, Options or a root lease.
Public Const COMMAND_PROTOCOL As String = "msaccess-vcs.session/1"
Public Const SERVER_MINIMUM As String = "0.3.0"
Public Const SERVER_MAXIMUM As String = "0.4.0"
Public Const SERVER_DEVELOPMENT As String = "0.3.0-dev.17"
Private Const MAX_SESSIONS As Long = 128
Private Const SESSION_IDLE_MINUTES As Long = 30
Private m_strInstance As String
Private m_strLoadedFile As String
Private m_strLoadedSnapshot As String
Private m_strEnvelope As String
Private m_strPolicyOwner As String
Private m_dSessions As Dictionary
Private m_lngHandshakes As Long
Private m_lngValidations As Long
Private m_lngDispatches As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Type FILE_INFORMATION
    Attributes As Long
    CreatedLow As Long
    CreatedHigh As Long
    AccessedLow As Long
    AccessedHigh As Long
    WrittenLow As Long
    WrittenHigh As Long
    Volume As Long
    SizeHigh As Long
    SizeLow As Long
    Links As Long
    IndexHigh As Long
    IndexLow As Long
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As GUID, ByVal buffer As LongPtr, ByVal capacity As Long) As Long
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal path As LongPtr, ByVal access As Long, ByVal share As Long, ByVal security As LongPtr, ByVal disposition As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileInformationByHandle Lib "kernel32" (ByVal handle As LongPtr, ByRef information As FILE_INFORMATION) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long

Private Function NewInstance() As String
    Dim value As GUID
    Dim buffer As String
    If CoCreateGuid(value) <> 0 Then Exit Function
    buffer = String$(39, vbNullChar)
    If StringFromGUID2(value, StrPtr(buffer), 39) = 0 Then Exit Function
    NewInstance = Mid$(buffer, 2, 36)
End Function

' Tables such as translation contributions are mutable runtime data; their writes
' do not replace the loaded VBA library. Track the physical file generation here.
Private Function LoadedFileIdentity(Optional ByRef snapshot As String = vbNullString) As String
    Dim handle As LongPtr
    Dim info As FILE_INFORMATION
    Dim path As String
    path = CodeProject.FullName
    ' OPEN_EXISTING, FILE_READ_ATTRIBUTES, sharing read/write/delete. No target open.
    handle = CreateFileW(StrPtr(path), &H80, 7, 0, 3, 0, 0)
    If handle = -1 Then Exit Function
    If GetFileInformationByHandle(handle, info) <> 0 Then
        LoadedFileIdentity = LCase$(path) & ":" & Hex$(info.Volume) & ":" & _
            Hex$(info.IndexHigh) & ":" & Hex$(info.IndexLow) & ":" & _
            Hex$(info.CreatedHigh) & ":" & Hex$(info.CreatedLow)
        snapshot = LoadedFileIdentity & ":" & Hex$(info.WrittenHigh) & ":" & _
            Hex$(info.WrittenLow) & ":" & Hex$(info.SizeHigh) & ":" & Hex$(info.SizeLow)
    End If
    CloseHandle handle
End Function

Private Function InstanceReady() As Boolean
    Dim identity As String
    Dim snapshot As String
    identity = LoadedFileIdentity(snapshot)
    If Len(identity) = 0 Then Exit Function
    If Len(m_strInstance) = 0 Then
        m_strInstance = NewInstance()
        m_strLoadedFile = identity
        m_strLoadedSnapshot = snapshot
        Set m_dSessions = New Dictionary
    End If
    ' Replacement cannot relabel an older loaded project as the new installation.
    InstanceReady = (Len(m_strInstance) > 0 And identity = m_strLoadedFile And snapshot = m_strLoadedSnapshot)
End Function

' Trusted internal data writers account for their own table writes. Unknown
' size/timestamp changes still fail closed, including same-file overwrites.
' These helpers are not automation commands and do not grant a session.
Public Function BeginCompatibilityDataWrite() As String
    If InstanceReady Then BeginCompatibilityDataWrite = m_strLoadedSnapshot
End Function

Public Sub EndCompatibilityDataWrite(ByVal token As String)
    Dim identity As String
    Dim snapshot As String
    On Error GoTo Finished
    If Len(token) = 0 Or token <> m_strLoadedSnapshot Then Exit Sub
    DBEngine.Idle dbRefreshCache
    identity = LoadedFileIdentity(snapshot)
    If identity = m_strLoadedFile And Len(snapshot) > 0 Then m_strLoadedSnapshot = snapshot
Finished:
End Sub

Public Function APIIdentity() As String
    Dim result As New Dictionary
    On Error GoTo Failed
    If Not InstanceReady Then GoTo Failed
    result.Add "success", True
    result.Add "addin_instance", m_strInstance
    result.Add "addin_version", CStr(CodeDb.Properties("AppVersion"))
    result.Add "loaded_path", CodeProject.FullName
    result.Add "installation_identity", m_strLoadedFile
    result.Add "protocol", COMMAND_PROTOCOL
    result.Add "supported_server_range", SupportedServerRange
    APIIdentity = ConvertToJson(result)
    Exit Function
Failed:
    APIIdentity = SessionFailure("loaded_identity_unconfirmed")
End Function

Public Function SupportedServerRange() As String
    SupportedServerRange = ">=" & SERVER_MINIMUM & " <" & SERVER_MAXIMUM & _
        "; prerelease=" & SERVER_DEVELOPMENT
End Function

' Strict SemVer; numeric components are compared as strings to avoid overflow.
Private Function ParseVersion(ByVal value As String) As Variant
    Dim re As Object
    Dim found As Object
    Dim parts As Variant
    Dim item As Variant
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-([0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*))?(\+([0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*))?$"
    If Not re.Test(value) Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then Exit Function
    Set found = re.Execute(value)(0)
    parts = Array(found.SubMatches(0), found.SubMatches(1), found.SubMatches(2), found.SubMatches(4))
    For Each item In Split(parts(3), ".")
        If NumericIdentifier(CStr(item)) And Len(item) > 1 Then
            If Left$(item, 1) = "0" Then Exit Function
        End If
    Next
    ParseVersion = parts
End Function

Private Function NumericIdentifier(ByVal value As String) As Boolean
    Dim i As Long
    If Len(value) = 0 Then Exit Function
    For i = 1 To Len(value)
        If Mid$(value, i, 1) < "0" Or Mid$(value, i, 1) > "9" Then Exit Function
    Next
    NumericIdentifier = True
End Function

Private Function CompareNumeric(ByVal left As String, ByVal right As String) As Long
    If Len(left) <> Len(right) Then
        CompareNumeric = Sgn(Len(left) - Len(right))
    Else
        CompareNumeric = Sgn(StrComp(left, right, vbBinaryCompare))
    End If
End Function

Public Function CompareSemanticVersions(ByVal left As String, ByVal right As String) As Variant
    Dim a As Variant, b As Variant, ap As Variant, bp As Variant
    Dim i As Long, compared As Long
    a = ParseVersion(left): b = ParseVersion(right)
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    For i = 0 To 2
        compared = CompareNumeric(CStr(a(i)), CStr(b(i)))
        If compared <> 0 Then CompareSemanticVersions = compared: Exit Function
    Next
    If a(3) = b(3) Then CompareSemanticVersions = 0: Exit Function
    If Len(a(3)) = 0 Then CompareSemanticVersions = 1: Exit Function
    If Len(b(3)) = 0 Then CompareSemanticVersions = -1: Exit Function
    ap = Split(a(3), "."): bp = Split(b(3), ".")
    For i = 0 To UBound(ap)
        If i > UBound(bp) Then CompareSemanticVersions = 1: Exit Function
        If NumericIdentifier(ap(i)) And NumericIdentifier(bp(i)) Then
            compared = CompareNumeric(CStr(ap(i)), CStr(bp(i)))
        ElseIf NumericIdentifier(ap(i)) <> NumericIdentifier(bp(i)) Then
            If NumericIdentifier(ap(i)) Then compared = -1 Else compared = 1
        Else
            compared = Sgn(StrComp(ap(i), bp(i), vbBinaryCompare))
        End If
        If compared <> 0 Then CompareSemanticVersions = compared: Exit Function
    Next
    CompareSemanticVersions = Sgn(UBound(ap) - UBound(bp))
End Function

Public Function ServerVersionReason(ByVal value As String) As String
    Dim parts As Variant
    Dim maximum As Variant
    Dim i As Long
    Dim compared As Long
    If Len(value) = 0 Then ServerVersionReason = "unknown_version": Exit Function
    parts = ParseVersion(value)
    If IsEmpty(parts) Then ServerVersionReason = "invalid_version": Exit Function
    maximum = ParseVersion(SERVER_MAXIMUM)
    For i = 0 To 2
        compared = CompareNumeric(CStr(parts(i)), CStr(maximum(i)))
        If compared < 0 Then Exit For
        If compared > 0 Or i = 2 Then
            ServerVersionReason = "unsupported_boundary": Exit Function
        End If
    Next
    If Len(parts(3)) > 0 Then
        If CompareSemanticVersions(value, SERVER_DEVELOPMENT) <> 0 Then ServerVersionReason = "prerelease_not_admitted"
    ElseIf CompareSemanticVersions(value, SERVER_MINIMUM) < 0 Then
        ServerVersionReason = "below_minimum"
    End If
End Function

Public Function SessionFailure(ByVal reason As String, Optional ByVal version As String, _
    Optional ByVal pattern As String = "compatibility_session_invalid") As String
    Dim result As New Dictionary
    Dim action As String
    action = T("Reconnect with a server satisfying {0} and complete the mutual handshake before dispatch. Do not replay an uncertain operation.", var0:=SupportedServerRange)
    If reason = "unsupported_boundary" Then action = T("Use a server/add-in combination explicitly supporting each other's versions and protocol. Updating the server again is not the remedy.")
    result.Add "success", False
    result.Add "started", False
    result.Add "error_pattern", pattern
    result.Add "component", "server"
    If Len(version) > 0 Then result.Add "installed_version", version Else result.Add "installed_version", Null
    result.Add "required_range", SupportedServerRange
    result.Add "minimum_version", SERVER_MINIMUM
    result.Add "requirement_status", "assigned_unpublished"
    result.Add "compatibility_reason", reason
    result.Add "recovery_action", action
    result.Add "error", T("Compatibility admission refused: {0}. {1}", var0:=reason, var1:=action)
    SessionFailure = ConvertToJson(result)
End Function

Public Function CompatibilitySessionExpired(ByVal touched As Date, ByVal checkedAt As Date, _
    Optional ByVal ownsPolicy As Boolean = False) As Boolean
    ' Idle owner policies require explicit owner cleanup (developer decision).
    CompatibilitySessionExpired = (Not ownsPolicy And checkedAt >= DateAdd("n", SESSION_IDLE_MINUTES, touched))
End Function

Private Sub PruneSessions()
    Dim key As Variant
    Dim entry As Dictionary
    Dim ownsPolicy As Boolean
    For Each key In m_dSessions.Keys
        Set entry = m_dSessions(key)
        ownsPolicy = (m_strPolicyOwner = entry("server_instance") & ":" & entry("connection_id"))
        If CompatibilitySessionExpired(entry("touched"), Now, ownsPolicy) Then m_dSessions.Remove key
    Next
End Sub

Public Function APIHandshake(ByVal request As String) As String
    Dim incoming As Dictionary
    Dim result As New Dictionary
    Dim reason As String
    Dim version As String
    Dim key As Variant
    On Error GoTo Failed
    If Not InstanceReady Then GoTo Failed
    Set incoming = ParseJson(request)
    If incoming.Exists("server_version") Then version = CStr(incoming("server_version"))
    m_lngHandshakes = m_lngHandshakes + 1
    reason = ServerVersionReason(version)
    If Len(reason) > 0 Then
        APIHandshake = SessionFailure(reason, version, IIf(reason = "invalid_version" Or reason = "unknown_version", "version_unconfirmed", "version_incompatible"))
        Exit Function
    End If
    If incoming("protocol") <> COMMAND_PROTOCOL Then reason = "protocol_mismatch": GoTo Refused
    If incoming("addin_instance") <> m_strInstance Then reason = "addin_instance_changed": GoTo Refused
    For Each key In Array("session_id", "server_instance", "connection_id")
        If Len(CStr(incoming(key))) <> 36 Then reason = "caller_identity_unconfirmed": GoTo Refused
    Next
    PruneSessions
    If m_dSessions.Exists(incoming("session_id")) Then reason = "session_id_already_used": GoTo Refused
    If m_dSessions.Count >= MAX_SESSIONS Then reason = "session_capacity": GoTo Refused
    incoming.Add "touched", Now
    incoming.Add "requirement", SupportedServerRange
    m_dSessions.Add incoming("session_id"), incoming
    result.Add "success", True
    For Each key In Array("session_id", "server_instance", "connection_id", "addin_instance", "protocol")
        result.Add key, incoming(key)
    Next
    APIHandshake = ConvertToJson(result)
    Exit Function
Refused:
    APIHandshake = SessionFailure(reason, version)
    Exit Function
Failed:
    APIHandshake = SessionFailure("handshake_unconfirmed", version)
End Function

Public Function APIValidateSession(ByVal envelope As String) As String
    Dim incoming As Dictionary
    Dim entry As Dictionary
    Dim key As Variant
    On Error GoTo Failed
    m_lngValidations = m_lngValidations + 1
    If Not InstanceReady Then GoTo Failed
    Set incoming = ParseJson(envelope)
    PruneSessions
    If Not m_dSessions.Exists(incoming("session_id")) Then GoTo Failed
    Set entry = m_dSessions(incoming("session_id"))
    For Each key In Array("server_instance", "connection_id", "addin_instance")
        If incoming(key) <> entry(key) Then GoTo Failed
    Next
    If entry("addin_instance") <> m_strInstance Or entry("requirement") <> SupportedServerRange Or entry("protocol") <> COMMAND_PROTOCOL Then GoTo Failed
    entry("touched") = Now
    APIValidateSession = "{""success"":true}"
    Exit Function
Failed:
    APIValidateSession = SessionFailure("missing_stale_or_foreign_session")
End Function

Public Function EnterCompatibilitySession(ByVal envelope As String) As String
    Dim checked As String
    If Len(m_strEnvelope) > 0 Then
        EnterCompatibilitySession = SessionFailure("nested_session_dispatch")
        Exit Function
    End If
    checked = APIValidateSession(envelope)
    If ParseJson(checked)("success") Then
        m_strEnvelope = envelope
    Else
        EnterCompatibilitySession = checked
    End If
End Function

Public Sub LeaveCompatibilitySession()
    m_strEnvelope = vbNullString
End Sub

Public Function CurrentCompatibilityEnvelope() As String
    CurrentCompatibilityEnvelope = m_strEnvelope
End Function

Public Function CompatibilityDispatchAllowed() As Boolean
    If Len(m_strEnvelope) = 0 Then Exit Function
    CompatibilityDispatchAllowed = CBool(ParseJson(APIValidateSession(m_strEnvelope))("success"))
End Function

Public Sub CountCompatibilityDispatch()
    m_lngDispatches = m_lngDispatches + 1
End Sub

Public Function APICompatibilityCounters() As String
    ' Read-only diagnostics; no session is established by inspecting these counters.
    Dim count As Long
    If Not (m_dSessions Is Nothing) Then count = m_dSessions.Count
    APICompatibilityCounters = "{""handshakes"":" & m_lngHandshakes & _
        ",""validations"":" & m_lngValidations & ",""dispatches"":" & m_lngDispatches & ",""sessions"":" & count & "}"
End Function

Public Function APIDisconnectSession(ByVal envelope As String) As String
    Dim incoming As Dictionary
    Dim checked As String
    checked = APIValidateSession(envelope)
    If CBool(ParseJson(checked)("success")) Then
        Set incoming = ParseJson(envelope)
        If m_strPolicyOwner = incoming("server_instance") & ":" & incoming("connection_id") Then
            If Operation.IsActive Or TimerIsPending Then
                APIDisconnectSession = SessionFailure("session_in_use")
                Exit Function
            End If
            ClearSessionPolicy Operation
            m_strPolicyOwner = vbNullString
        End If
        m_dSessions.Remove incoming("session_id")
    End If
    APIDisconnectSession = checked
End Function

Public Function APIExecute(ByVal envelope As String, ByVal method As String, _
    Optional arg1 As Variant, Optional arg2 As Variant, Optional arg3 As Variant) As Variant
    Dim refused As String
    On Error GoTo Failed
    refused = EnterCompatibilitySession(envelope)
    If Len(refused) > 0 Then APIExecute = refused: Exit Function
    refused = CompatibilityCommandRefusal(method)
    If Len(refused) > 0 Then
        APIExecute = refused
        LeaveCompatibilitySession
        Exit Function
    End If
    If Not IsMissing(arg3) Then
        APIExecute = API(method, arg1, arg2, arg3, strSessionEnvelope:=envelope)
    ElseIf Not IsMissing(arg2) Then
        APIExecute = API(method, arg1, arg2, strSessionEnvelope:=envelope)
    ElseIf Not IsMissing(arg1) Then
        APIExecute = API(method, arg1, strSessionEnvelope:=envelope)
    Else
        APIExecute = API(method, strSessionEnvelope:=envelope)
    End If
    RememberCompatibilityPolicy method, APIExecute
    LeaveCompatibilitySession
    Exit Function
Failed:
    LeaveCompatibilitySession
    APIExecute = AsyncRequestFailure(vbNullString, Err.Description, lngNumber:=Err.Number)
End Function

Public Function APIExecuteAsync(ByVal envelope As String, ByVal callback As String, ByVal method As String, _
    Optional arg1 As Variant, Optional arg2 As Variant) As String
    Dim refused As String
    On Error GoTo Failed
    refused = EnterCompatibilitySession(envelope)
    If Len(refused) > 0 Then
        APIExecuteAsync = PostCompatibilityRefusal(callback, refused)
        Exit Function
    End If
    refused = CompatibilityCommandRefusal(method)
    If Len(refused) > 0 Then
        APIExecuteAsync = PostCompatibilityRefusal(callback, refused)
        LeaveCompatibilitySession
        Exit Function
    End If
    If Not IsMissing(arg2) Then
        APIExecuteAsync = APIAsync(callback, method, arg1, arg2, strSessionEnvelope:=envelope)
    ElseIf Not IsMissing(arg1) Then
        APIExecuteAsync = APIAsync(callback, method, arg1, strSessionEnvelope:=envelope)
    Else
        APIExecuteAsync = APIAsync(callback, method, strSessionEnvelope:=envelope)
    End If
    LeaveCompatibilitySession
    Exit Function
Failed:
    LeaveCompatibilitySession
    APIExecuteAsync = AsyncRequestFailure(callback, Err.Description, lngNumber:=Err.Number)
End Function

Public Function PostCompatibilityRefusal(ByVal callback As String, ByVal refused As String) As String
    Dim emitter As New clsMCP
    Dim result As Dictionary
    Set result = ParseJson(refused)
    emitter.RegisterCallback callback
    If emitter.IsActive Then emitter.PostCallback "error", -1, -1, CStr(result("error")), result
    PostCompatibilityRefusal = refused
End Function

' An admitted rebuild delegates to a fresh connection in a fresh Access instance.
' Carry versions only on this new handshake, never on ordinary commands.
Public Function WorkerCompatibilityArguments() As Variant
    Dim incoming As Dictionary
    Dim entry As Dictionary
    If Not CompatibilityDispatchAllowed Then Exit Function
    Set incoming = ParseJson(m_strEnvelope)
    Set entry = m_dSessions(incoming("session_id"))
    WorkerCompatibilityArguments = Array(CStr(entry("server_version")), _
        CStr(entry("server_instance")), NewInstance(), NewInstance(), _
        CStr(CodeDb.Properties("AppVersion")))
End Function

Private Function CurrentCaller() As String
    Dim request As Dictionary
    Set request = ParseJson(m_strEnvelope)
    CurrentCaller = request("server_instance") & ":" & request("connection_id")
End Function

Public Function CompatibilityCommandRefusal(ByVal method As String) As String
    If Len(m_strPolicyOwner) = 0 Then Exit Function
    If m_strPolicyOwner = CurrentCaller Then Exit Function
    Select Case method
        Case "GetOption", "GetLogContent", "GetExportFolder", "GetProjectName", "IsVBACompiled", "IsDatabaseOpen"
            Exit Function
    End Select
    CompatibilityCommandRefusal = SessionFailure("interaction_policy_owned_by_another_caller")
End Function

Public Sub RememberCompatibilityPolicy(ByVal method As String, ByVal returned As Variant)
    Dim result As Dictionary
    If method <> "SetOperationPolicy" And method <> "ClearOperationPolicy" Then Exit Sub
    On Error GoTo Finished
    Set result = ParseJson(CStr(returned))
    If result("success") = True Then
        If method = "SetOperationPolicy" Then m_strPolicyOwner = CurrentCaller Else m_strPolicyOwner = vbNullString
    End If
Finished:
End Sub

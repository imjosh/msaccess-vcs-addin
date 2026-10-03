Attribute VB_Name = "modTestCompatibilitySession"
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.API")
Private m_lngRequest As Long

Public Sub StrictServerVersions()
    TestAssert ServerVersionReason("0.2.99") = "below_minimum"
    TestAssert ServerVersionReason("0.3.0") = vbNullString
    TestAssert ServerVersionReason("0.3.10") = vbNullString
    TestAssert ServerVersionReason("0.4.0") = "unsupported_boundary"
    TestAssert ServerVersionReason("1.0.0") = "unsupported_boundary"
    TestAssert ServerVersionReason("0.3.0-dev.17+local.9") = vbNullString
    TestAssert ServerVersionReason("0.3.0-dev.16") = "prerelease_not_admitted"
    TestAssert ServerVersionReason("0.3.0-01") = "invalid_version"
    TestAssert ServerVersionReason("0.03.0") = "invalid_version"
    TestAssert ServerVersionReason("v0.3.0") = "invalid_version"
    TestAssert ServerVersionReason("0.3.0" & vbLf) = "invalid_version"
    TestAssert ServerVersionReason(vbNullString) = "unknown_version"
End Sub

Public Sub SemanticOrdering()
    TestAssert CompareSemanticVersions("1.9.0", "1.10.0") = -1
    TestAssert CompareSemanticVersions("1.0.0-alpha.2", "1.0.0-alpha.11") = -1
    TestAssert CompareSemanticVersions("1.0.0-2", "1.0.0-alpha") = -1
    TestAssert CompareSemanticVersions("1.0.0-rc.1", "1.0.0") = -1
    TestAssert CompareSemanticVersions("1.0.0+x", "1.0.0+y") = 0
    TestAssert CompareSemanticVersions("99999999999999999999.0.0", "99999999999999999998.0.0") = 1
    TestAssert IsEmpty(CompareSemanticVersions("1.0", "1.0.0"))
End Sub

Public Sub LegacyAdmissionChangesNoOperationState()
    Dim before As String
    Dim counters As Dictionary
    Dim result As Dictionary
    before = OperationState
    Set counters = ParseJson(APICompatibilityCounters())
    Set result = ParseJson(CStr(API("SetInteractionMode", eimNonInteractive)))
    TestAssert result("success") = False
    TestAssert result("error_pattern") = "compatibility_session_invalid"
    TestAssert OperationState = before
    TestAssert ParseJson(APICompatibilityCounters())("dispatches") = counters("dispatches")
    Set result = ParseJson(APIAsync(vbNullString, "Export"))
    TestAssert result("success") = False
    TestAssert OperationState = before
    TestAssert Not TimerIsPending
End Sub

Public Sub RefusedHandshakeChangesNoOperationState()
    Dim request As Dictionary
    Dim result As Dictionary
    Dim before As String
    before = OperationState
    Set request = NewRequest("0.2.99")
    Set result = ParseJson(APIHandshake(ConvertToJson(request)))
    TestAssert result("success") = False
    TestAssert result("compatibility_reason") = "below_minimum"
    TestAssert OperationState = before
    Set result = ParseJson(APIValidateSession(ConvertToJson(request)))
    TestAssert result("success") = False
End Sub

Public Sub SessionsAreCachedAndCallerBound()
    Dim request As Dictionary
    Dim result As Dictionary
    Dim counters As Dictionary
    Dim i As Long
    Dim before As String
    before = OperationState
    Set request = NewRequest(SERVER_DEVELOPMENT)
    Set result = ParseJson(APIHandshake(ConvertToJson(request)))
    TestAssert result("success") = True
    request.Remove "server_version"
    request.Remove "protocol"
    Set counters = ParseJson(APICompatibilityCounters())
    For i = 1 To 10
        Set result = ParseJson(APIValidateSession(ConvertToJson(request)))
        TestAssert result("success") = True
    Next
    TestAssert ParseJson(APICompatibilityCounters())("handshakes") = counters("handshakes")
    request("connection_id") = "00000000-0000-0000-0000-000000000002"
    Set result = ParseJson(APIValidateSession(ConvertToJson(request)))
    TestAssert result("success") = False
    TestAssert OperationState = before
    request("connection_id") = "00000000-0000-0000-0000-000000000001"
    Set result = ParseJson(APIDisconnectSession(ConvertToJson(request)))
    TestAssert result("success") = True
    Set result = ParseJson(APIValidateSession(ConvertToJson(request)))
    TestAssert result("success") = False
End Sub

Private Function OperationState() As String
    OperationState = CStr(Operation.Status) & ":" & CStr(Operation.Source) & ":" & _
        CStr(Operation.InteractionMode) & ":" & CStr(Operation.DecisionPolicy) & ":" & Operation.CurrentRootToken
End Function

Private Function NewRequest(ByVal version As String) As Dictionary
    Dim request As Dictionary
    Dim identity As Dictionary
    Set identity = ParseJson(APIIdentity())
    Set request = New Dictionary
    m_lngRequest = m_lngRequest + 1
    request.Add "session_id", Left$(identity("addin_instance"), 32) & Format$(m_lngRequest, "0000")
    request.Add "server_instance", "00000000-0000-0000-0000-000000000000"
    request.Add "connection_id", "00000000-0000-0000-0000-000000000001"
    request.Add "server_version", version
    request.Add "protocol", COMMAND_PROTOCOL
    request.Add "addin_instance", identity("addin_instance")
    Set NewRequest = request
End Function

Public Sub ForeignSessionCannotClearCallerPolicy()
    Dim first As Dictionary, second As Dictionary, result As Dictionary
    Dim before As String
    Set first = NewRequest(SERVER_DEVELOPMENT)
    Set second = NewRequest(SERVER_DEVELOPMENT)
    second("connection_id") = "00000000-0000-0000-0000-000000000002"
    Set result = ParseJson(APIHandshake(ConvertToJson(first)))
    TestAssert result("success") = True
    Set result = ParseJson(APIHandshake(ConvertToJson(second)))
    TestAssert result("success") = True
    Set result = ParseJson(CStr(APIExecute(ConvertToJson(first), "SetOperationPolicy", "block")))
    TestAssert result("success") = True
    before = OperationState
    Set result = ParseJson(CStr(APIExecute(ConvertToJson(second), "ClearOperationPolicy")))
    TestAssert result("success") = False
    TestAssert result("compatibility_reason") = "interaction_policy_owned_by_another_caller"
    TestAssert OperationState = before
    Set result = ParseJson(CStr(APIExecute(ConvertToJson(first), "ClearOperationPolicy")))
    TestAssert result("success") = True
    APIDisconnectSession ConvertToJson(first)
    APIDisconnectSession ConvertToJson(second)
End Sub

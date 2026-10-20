VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperationRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Carries incoming request identity independently of active operation state.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")

' Incoming information is independent of the operation which may already be active.
Public Source As eOperationSource
Public PublicId As String
Public CallbackInfo As String
Public ReservationToken As String
Public Claimed As Boolean
Public Consumed As Boolean


'---------------------------------------------------------------------------------------
' Procedure : ToJson
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Serialize queued request ownership without changing an active operation.
'---------------------------------------------------------------------------------------
'
Public Function ToJson() As String
    Dim d As New Dictionary
    InitializeRequestErrors "ToJson"
    On Error GoTo Failed
    d("source") = Source
    d("operation_id") = PublicId
    d("callback_info") = CallbackInfo
    d("reservation_token") = ReservationToken
    d("claimed") = Claimed
    ToJson = ConvertToJson(d)
    Exit Function
Failed:
    Err.Raise Err.Number, "clsOperationRequest.ToJson", Err.Description
End Function


'---------------------------------------------------------------------------------------
' Procedure : FromJson
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Restore an incoming request. Parsing errors reach its owning dispatcher.
'---------------------------------------------------------------------------------------
'
Public Sub FromJson(ByVal strJson As String)
    Dim d As Dictionary
    InitializeRequestErrors "FromJson"
    On Error GoTo Failed
    Set d = ParseJson(strJson)
    Source = CLng(d("source"))
    PublicId = CStr(d("operation_id"))
    CallbackInfo = CStr(d("callback_info"))
    ReservationToken = CStr(d("reservation_token"))
    Claimed = CBool(d("claimed"))
    Exit Sub
Failed:
    Err.Raise Err.Number, "clsOperationRequest.FromJson", Err.Description
End Sub

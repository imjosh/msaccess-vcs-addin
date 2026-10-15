Attribute VB_Name = "vcs_diagnostic_write_probe"
Option Compare Database
Option Explicit

' Run in a disposable source-built development host before any root starts.
' The callback endpoint must be owned by the caller. No HTTP request is sent here.
Public Function DiagnosticWriteProbe(ByVal strCallback As String, ByVal strMessage As String) As String
    Dim d As New Dictionary
    Dim lngJournal As Long
    Dim lngErrors As Long
    Dim eLevel As eErrorLevel
    Dim eState As eOperationState
    Dim eMode As eInteractionMode
    Dim eSavedMode As eInteractionMode
    On Error GoTo Failed
    eSavedMode = Operation.InteractionMode
    Operation.InteractionMode = eimNormal
    lngJournal = Log.ErrorJournalCount
    lngErrors = Log.ErrorCount
    eLevel = Operation.ErrorLevel
    eState = Operation.Status
    eMode = Operation.InteractionMode
    d.Add "export_folder", Options.GetExportFolder
    d.Add "status_before", CLng(eState)
    d.Add "mode_before", CLng(eMode)
    d.Add "original_mode", CLng(eSavedMode)
    MCP.RegisterCallback strCallback
    d.Add "active", MCP.IsActive
    MCPDebugLog strMessage
    MCP.RegisterCallback "{invalid callback"
    d.Add "registration_error_consumed", Err.Number = 0
    d.Add "failed_registration_released", Not MCP.IsActive
    d.Add "failed_registration_fields_cleared", Len(MCP.CallbackUrl) = 0 And Len(MCP.OperationId) = 0
    d.Add "journal_delta", Log.ErrorJournalCount - lngJournal
    d.Add "error_delta", Log.ErrorCount - lngErrors
    d.Add "level_unchanged", Operation.ErrorLevel = eLevel
    d.Add "status_unchanged", Operation.Status = eState
    d.Add "mode_unchanged", Operation.InteractionMode = eMode
    MCP.RegisterCallback vbNullString
    d.Add "callback_released", Not MCP.IsActive
    Operation.InteractionMode = eSavedMode
    d.Add "original_mode_restored", Operation.InteractionMode = eSavedMode
    DiagnosticWriteProbe = modJsonConverter.ConvertToJson(d)
    Exit Function
Failed:
    d.Add "probe_error", CStr(Err.Number) & ": " & Err.Description
    MCP.RegisterCallback vbNullString
    Operation.InteractionMode = eSavedMode
    DiagnosticWriteProbe = modJsonConverter.ConvertToJson(d)
End Function

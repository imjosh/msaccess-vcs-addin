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
    On Error GoTo Failed
    lngJournal = Log.ErrorJournalCount
    lngErrors = Log.ErrorCount
    eLevel = Operation.ErrorLevel
    eState = Operation.Status
    eMode = Operation.InteractionMode
    d.Add "export_folder", Options.GetExportFolder
    d.Add "status_before", CLng(eState)
    d.Add "mode_before", CLng(eMode)
    MCP.RegisterCallback strCallback
    d.Add "active", MCP.IsActive
    MCPDebugLog strMessage
    d.Add "journal_delta", Log.ErrorJournalCount - lngJournal
    d.Add "error_delta", Log.ErrorCount - lngErrors
    d.Add "level_unchanged", Operation.ErrorLevel = eLevel
    d.Add "status_unchanged", Operation.Status = eState
    d.Add "mode_unchanged", Operation.InteractionMode = eMode
    MCP.RegisterCallback vbNullString
    d.Add "callback_released", Not MCP.IsActive
    DiagnosticWriteProbe = modJsonConverter.ConvertToJson(d)
    Exit Function
Failed:
    d.Add "probe_error", CStr(Err.Number) & ": " & Err.Description
    MCP.RegisterCallback vbNullString
    DiagnosticWriteProbe = modJsonConverter.ConvertToJson(d)
End Function

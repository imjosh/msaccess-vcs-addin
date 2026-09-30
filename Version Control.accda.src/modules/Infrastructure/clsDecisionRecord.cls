VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDecisionRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDecisionRecord
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : One entry in the decision journal: the five fields that always travel
'           : together when a prompt or a conflict is recorded. Build one with
'           : NewDecision. The fields become the kind, title, message, resolution and
'           : detail keys of the decisions the MCP reports, so their values are wire
'           : text and are not translated here.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Public Kind As String
Public Title As String
Public Message As String
Public Resolution As String
Public Detail As String

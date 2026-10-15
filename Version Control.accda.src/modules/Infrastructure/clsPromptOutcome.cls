VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPromptOutcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsPromptOutcome
' Author    : Josh
' Date      : 09/29/2026
' Purpose   : How a decision policy answered one message box without showing it: the
'           : button to hand back to the caller, the resolution word (acknowledged,
'           : declined or decision_required) and whether the prompt is left blocked.
'           : Built by ResolveNonInteractivePrompt.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Public Result As VbMsgBoxResult
Public Resolution As String
Public Blocked As Boolean

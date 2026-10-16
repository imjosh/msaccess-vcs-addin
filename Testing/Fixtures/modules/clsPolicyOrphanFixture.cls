VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPolicyOrphanFixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsPolicyOrphanFixture
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Reproducible disposable class-module source for canonical orphan history.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : FixtureValue
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Supply an actual editable class-module definition.
'---------------------------------------------------------------------------------------
'
Public Function FixtureValue() As Long
    FixtureValue = 1
End Function

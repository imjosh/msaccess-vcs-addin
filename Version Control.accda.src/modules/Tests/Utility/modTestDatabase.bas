Attribute VB_Name = "modTestDatabase"
'---------------------------------------------------------------------------------------
' Module    : modTestDatabase
' Author    : Adam Waller
' Date      : 6/19/2026
' Purpose   : Unit tests for modDatabase utility functions (engine-managed property
'           : handling for table/object property round-tripping).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")
'@Tag("unit")


Public Sub TestIsEngineManagedProperty()
    ' The FCMin* feature-compatibility version stamps are engine-managed.
    TestAssert IsEngineManagedProperty("FCMinDesignVer"), "FCMinDesignVer"
    TestAssert IsEngineManagedProperty("FCMinReadVer"), "FCMinReadVer"
    TestAssert IsEngineManagedProperty("FCMinWriteVer"), "FCMinWriteVer"
    TestAssert IsEngineManagedProperty("fcminwritever"), "case-insensitive match"
    ' Ordinary display/custom properties are not.
    TestAssert Not IsEngineManagedProperty("Description"), "Description is settable"
    TestAssert Not IsEngineManagedProperty("ColumnWidth"), "ColumnWidth is settable"
    TestAssert Not IsEngineManagedProperty("FC"), "shorter than FCMin prefix"
    TestAssert Not IsEngineManagedProperty(vbNullString), "empty string"
End Sub


Public Sub TestFilterEngineManagedProps()
    Dim dIn As Dictionary
    Dim dOut As Dictionary

    Set dIn = New Dictionary
    dIn.CompareMode = TextCompare
    dIn.Add "Description", "desc"
    dIn.Add "FCMinDesignVer", "16.0.12600.10000"
    dIn.Add "FCMinReadVer", "16.0.12600.10000"
    dIn.Add "FCMinWriteVer", "16.0.12600.10000"
    dIn.Add "ColumnWidth", 1440

    Set dOut = FilterEngineManagedProps(dIn)
    TestAssert dOut.Count = 2, "only non-engine props remain"
    TestAssert dOut.Exists("Description"), "Description kept"
    TestAssert dOut.Exists("ColumnWidth"), "ColumnWidth kept"
    TestAssert Not dOut.Exists("FCMinDesignVer"), "FCMinDesignVer removed"
    TestAssert Not dOut.Exists("FCMinWriteVer"), "FCMinWriteVer removed"
    TestAssert dOut.CompareMode = TextCompare, "compare mode preserved"
    TestAssert dIn.Count = 5, "input dictionary not mutated"
End Sub


Public Sub TestFilterEngineManagedProps_NoEngineProps()
    Dim dIn As Dictionary
    Dim dOut As Dictionary

    Set dIn = New Dictionary
    dIn.Add "Description", "desc"
    dIn.Add "Caption", "cap"

    Set dOut = FilterEngineManagedProps(dIn)
    TestAssert dOut.Count = 2, "all properties retained when none are engine-managed"
End Sub


Public Sub TestCloseOpenObjectsForTypeNoOpTypes()
    TestAssert CloseOpenObjectsForType(edbCommandBar, acSaveYes), "command bars no-op succeeds"
    TestAssert CloseOpenObjectsForType(edbModule, acSaveYes), "modules no-op succeeds"
    TestAssert CloseOpenObjectsForType(edbVbeReference, acSaveYes), "references no-op succeeds"
End Sub


Public Sub TestCloseOpenObjectsForTypeNoOpTypesReportNoCancel()
    Dim blnCanceled As Boolean
    blnCanceled = True
    TestAssert CloseOpenObjectsForType(edbModule, acSaveYes, blnCanceled), "modules no-op succeeds"
    TestAssert Not blnCanceled, "a close that did nothing is not a cancellation"
    blnCanceled = True
    TestAssert CloseOpenObjectsForContainers(New Collection, acSaveYes, blnCanceled), "no containers succeed"
    TestAssert Not blnCanceled, "no containers is not a cancellation"
End Sub


Public Sub TestIsNativeCloseCancel()
    TestAssert IsNativeCloseCancel(2501, True), "canceled prompt with the object open is a native cancel"
    TestAssert Not IsNativeCloseCancel(2501, False), "2501 with the object closed is not a cancel"
    TestAssert Not IsNativeCloseCancel(2467, True), "another close error with the object open is ordinary"
    TestAssert Not IsNativeCloseCancel(0, True), "no error is not a cancel"
End Sub


Public Sub TestOperationNativeCloseCanceledFlag()
    Dim blnWas As Boolean
    blnWas = Operation.NativeCloseCanceled
    Operation.NativeCloseCanceled = True
    TestAssert Operation.NativeCloseCanceled, "the flag records a canceled native close"
    Operation.NativeCloseCanceled = False
    TestAssert Not Operation.NativeCloseCanceled, "the flag can be cleared"
    Operation.NativeCloseCanceled = blnWas
End Sub

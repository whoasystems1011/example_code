Attribute VB_Name = "HostConstructors"
' WARNING TREAT THIS CODE MODULE AS READ-ONLY (CODE ON THIS MODULE IS PERIODICALLY DELETED AND REBUILT)

Option Explicit
Option Private Module

Public Const MODEL_NAME_MULTIFAMILY_AMENITY As String = "MultifamilyAmenity"
Public Const MODEL_NAME_MULTIFAMILY_RENT_COMP_UNIT As String = "MultifamilyRentCompUnit"
Public Const MODEL_NAME_MULTIFAMILY_RENT_COMP As String = "MultifamilyRentComp"


Public Function GetDb_MultifamilyAmenity() As clsDb
    Set GetDb_MultifamilyAmenity = GetDbOrRaise(ThisWorkbook, "MultifamilyAmenity")
End Function

Public Function GetDb_MultifamilyRentCompUnit() As clsDb
    Set GetDb_MultifamilyRentCompUnit = GetDbOrRaise(ThisWorkbook, "MultifamilyRentCompUnit")
End Function

Public Function GetDb_MultifamilyRentComp() As clsDb
    Set GetDb_MultifamilyRentComp = GetDbOrRaise(ThisWorkbook, "MultifamilyRentComp")
End Function



Public Function GetSheetInfoById(record_id As String, Optional always_false As Boolean = False) As clsSheetInfo
    Dim record As New clsSheetInfo
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetSheetInfoById = record
    Else
        Set GetSheetInfoById = Nothing
    End If
End Function

Public Function GetShelterById(record_id As String, Optional always_false As Boolean = False) As clsShelter
    Dim record As New clsShelter
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetShelterById = record
    Else
        Set GetShelterById = Nothing
    End If
End Function

Public Function GetTestCatById(record_id As String, Optional always_false As Boolean = False) As clsTestCat
    Dim record As New clsTestCat
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetTestCatById = record
    Else
        Set GetTestCatById = Nothing
    End If
End Function

Public Function GetMultifamilyAmenityById(record_id As String, Optional always_false As Boolean = False) As clsMultifamilyAmenity
    Dim record As New clsMultifamilyAmenity
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetMultifamilyAmenityById = record
    Else
        Set GetMultifamilyAmenityById = Nothing
    End If
End Function

Public Function GetMultifamilyRentCompUnitById(record_id As String, Optional always_false As Boolean = False) As clsMultifamilyRentCompUnit
    Dim record As New clsMultifamilyRentCompUnit
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetMultifamilyRentCompUnitById = record
    Else
        Set GetMultifamilyRentCompUnitById = Nothing
    End If
End Function

Public Function GetMultifamilyRentCompById(record_id As String, Optional always_false As Boolean = False) As clsMultifamilyRentComp
    Dim record As New clsMultifamilyRentComp
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetMultifamilyRentCompById = record
    Else
        Set GetMultifamilyRentCompById = Nothing
    End If
End Function

Private Function TestMacrosAreEnabled() As Boolean
    ' NOTE this function is called on Setting Registry sheet
    TestMacrosAreEnabled = True
End Function

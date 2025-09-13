Attribute VB_Name = "ResetMultifamily"
Option Explicit
Option Private Module

Const module_name As String = "ResetMultifamily"
Private msg As String


Public Sub ResetMultifamilyComponents()
    ' NOTE
    ' - this Public procedure is optional
    ' - calling this will run only the resets on this module
    ' - sometimes you want to reset just one component (Multifamily) without affecting others
    ResetSingleModule ThisWorkbook, ThisWorkbook, module_name
End Sub


'------------------|
'--Private Resets--|
'------------------|
Private Sub Reset_MultifamilyRentComps_AndUnits()
    ' NOTE
    ' - first delete the unit records (they depend on the property records, deleted second)
    Call DeleteAllDbTableRowsByModel(ThisWorkbook, MODEL_NAME_MULTIFAMILY_RENT_COMP_UNIT)
    Call DeleteAllDbTableRowsByModel(ThisWorkbook, MODEL_NAME_MULTIFAMILY_RENT_COMP)
End Sub



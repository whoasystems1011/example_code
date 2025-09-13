Attribute VB_Name = "ValidateBusinessExpense"
Option Explicit
Option Private Module

Private Const module_name As String = "ValidateBusinessExpense"
Private msg_fail As String


Private Sub Validate_BusinessExpense_Module()
    Call ValidateSingleModule(ThisWorkbook, ThisWorkbook, module_name)
End Sub



'------------------|
'--Private Checks--|
'------------------|
Private Function GetValidateErrors_BusinessExpenseRecords(WB_HOST As Workbook) As ErrorBank
    Const OWNER As String = "GetValidateErrors_BusinessExpenseRecords()"
    Dim EB_MAIN As ErrorBank
    Set EB_MAIN = CreateErrorBank(OWNER)
    
    Dim EB_RECORD As ErrorBank
    Dim record As clsBusinessExpense
    
    For Each record In GetDb_BusinessExpense.GetAllRecords
        Set EB_RECORD = record.GetValidateErrors
        Set EB_MAIN = EB_MAIN.MergeBanks(EB_RECORD)
    Next record
    
    Set GetValidateErrors_BusinessExpenseRecords = EB_MAIN
End Function





Attribute VB_Name = "ValidateAI_Modules"
Option Explicit
Option Private Module

Private Const module_name As String = "ValidateAI_Modules"
Private msg_fail As String


Private Sub Sub_ValidateAI_Modules()
    Call ValidateSingleModule(ThisWorkbook, ThisWorkbook, module_name)
End Sub



'------------------|
'--Private Checks--|
'------------------|
Private Sub ScratchTEST_XXXX()
    Dim EB As ErrorBank
    Set EB = GetValidateErrors_PromptRecords(ThisWorkbook)
    Call EB.Render_MessageBoxErrorList(True)
End Sub


Private Function GetValidateErrors_PromptRecords(WB_HOST As Workbook) As ErrorBank
    Const OWNER As String = "GetValidateErrors_PromptRecords()"
    Dim EB_MAIN As ErrorBank
    Set EB_MAIN = CreateErrorBank(OWNER)
    
    Dim EB_RECORD As ErrorBank
    Dim record As clsPrompt
    
    For Each record In GetDb_Prompt.GetAllRecords
        Set EB_RECORD = record.GetValidateErrors
        Set EB_MAIN = EB_MAIN.MergeBanks(EB_RECORD)
    Next record
    
    'EB_MAIN.AssertFalse True, "Intentional assert failure for python testing"
    
    Set GetValidateErrors_PromptRecords = EB_MAIN
End Function






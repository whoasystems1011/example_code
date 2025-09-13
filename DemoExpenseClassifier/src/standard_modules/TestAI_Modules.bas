Attribute VB_Name = "TestAI_Modules"
Option Explicit
Option Private Module
Private Const module_name As String = "TestAI_Modules"

Private msg_fail As String


Private Sub Test_AI_ModulesSub()
    Call TestSingleModule(ThisWorkbook, ThisWorkbook, module_name)
End Sub





Private Function GetTestErrors_AI_ConstantsAndRanges(WB_HOST As Workbook) As ErrorBank
    Const OWNER As String = "GetTestErrors_AI_ConstantsAndRanges()"
    Dim EB As ErrorBank
    Set EB = CreateErrorBank(OWNER)

    ' ARRANGE
    Dim DB As clsDb
    Set DB = GetDb_Prompt
    
    ' CHECK all the constants exist as records in the DataBase
    EB.AssertType DB.GetRecordByIndex(PROMPT_NAME_EXPENSE_CLASSIFY_SELECT), "clsPrompt", "#T1"
    EB.AssertType DB.GetRecordByIndex(PROMPT_NAME_EXPENSE_CLASSIFY_MODEL), "clsPrompt", "#T2"
    
    ' Add additional PROMPT_NAME constants here as they are created...
    
    ' ROLLUP
    Set GetTestErrors_AI_ConstantsAndRanges = EB
End Function


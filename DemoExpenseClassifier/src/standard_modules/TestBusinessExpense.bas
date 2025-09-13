Attribute VB_Name = "TestBusinessExpense"
Option Explicit
Option Private Module

Private Const module_name As String = "TestBusinessExpense"
Private msg_fail As String


Private Sub Test_BusinessExpense()
    Call TestSingleModule(ThisWorkbook, ThisWorkbook, module_name)
End Sub



'------------------|
'--Private Checks--|
'------------------|
Private Function GetTestErrorsConstantsAndRanges(WB_HOST As Workbook) As ErrorBank
    Const OWNER As String = "GetTestErrorsConstantsAndRanges()"
    Dim EB As ErrorBank
    Set EB = CreateErrorBank(OWNER)
    
    Dim ChoiceRange As Range
    Set ChoiceRange = GetRangeOrRaise(ThisWorkbook, RANGENAME_CHOICES_COST_CATEGORY)
    
    Const MINIMUM_ALLOWED_CHOICES As Long = 15
    
    ' CHECK basic aspects of cost category choices range
    EB.AssertTrue ChoiceRange.Cells.Count >= MINIMUM_ALLOWED_CHOICES, "#T1"
    EB.AssertRangeContains ChoiceRange, "", "#T2 blank string should be an allowed choice"
    EB.AssertRangeContains ChoiceRange, COST_CATEGORY_SUPPLIES, "#T3"
    EB.AssertRangeContains ChoiceRange, COST_CATEGORY_NEEDS_REVIEW, "#T4"

    ' ROLLUP
    Set GetTestErrorsConstantsAndRanges = EB
End Function



Private Function GetTestErrors_BusinessExpense_Validation(WB_HOST As Workbook) As ErrorBank
    Const OWNER As String = "GetTestErrors_BusinessExpense_Validation"
    Dim EB_MAIN As ErrorBank
    Set EB_MAIN = CreateErrorBank(OWNER)

    ' ARRANGE
    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense_FIXTURE_LOADED(EB_MAIN)
    
    Dim EB_VALIDATE As ErrorBank
    Dim record As clsBusinessExpense
    Set record = DB.GetFirstRecordOrNothing
    
    ' CHECK valid state 1
    record.cost_category = ""
    record.followup_question = ""
    Set EB_VALIDATE = record.GetValidateErrors
    
    EB_MAIN.AssertEqual EB_VALIDATE.CountErrors, 0, "#T1"
    
    ' CHECK valid state 2
    record.cost_category = COST_CATEGORY_SUPPLIES
    record.followup_question = ""
    Set EB_VALIDATE = record.GetValidateErrors
    
    EB_MAIN.AssertEqual EB_VALIDATE.CountErrors, 0, "#T2"
    
    ' CHECK invalid state 1 (flagged for context, but followup question is blank)
    record.cost_category = COST_CATEGORY_NEEDS_REVIEW
    record.followup_question = ""
    Set EB_VALIDATE = record.GetValidateErrors
    
    Dim error_text As String
    error_text = EB_VALIDATE.GetErrorsPlaintext_String

    EB_MAIN.AssertEqual EB_VALIDATE.CountErrors, 1, "#T3"
    EB_MAIN.AssertStringContains error_text, "cannot be blank", "#T4"

    ' CHECK invalid state 2 (followup question is non-blank, but record is tagged for review)
    record.cost_category = ""
    record.followup_question = "This should not be populated"
    Set EB_VALIDATE = record.GetValidateErrors
    
    EB_MAIN.AssertEqual EB_VALIDATE.CountErrors, 1, "#T5"
    
    error_text = EB_VALIDATE.GetErrorsPlaintext_String
    EB_MAIN.AssertStringContains error_text, "should be blank", "#T6"
    
    ' CLEANUP
    ' - clear both local fields to avoid leaving this workbook in an invalid state
    ' - this approach is not robust or ideal; but it solves the problem today
    record.cost_category = ""
    record.followup_question = ""

    ' ROLLUP
    Set GetTestErrors_BusinessExpense_Validation = EB_MAIN
End Function



Private Function GetTestErrors_BusinessExpense_Status(WB_HOST As Workbook) As ErrorBank
    Const OWNER As String = "GetTestErrors_BusinessExpense_Status()"
    Dim EB As ErrorBank
    Set EB = CreateErrorBank(OWNER)
    
    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense_FIXTURE_LOADED(EB)

    Dim record As clsBusinessExpense
    Set record = DB.GetFirstRecordOrNothing
    
    ' CHECK unclassified record
    record.cost_category = ""
    EB.AssertEqual record.status, Status_Unclassified, "#T1"

    ' CHECK classified record
    record.cost_category = COST_CATEGORY_SUPPLIES
    EB.AssertEqual record.status, Status_Classified, "#T2"

    ' CHECK needs review record
    record.cost_category = COST_CATEGORY_NEEDS_REVIEW
    EB.AssertEqual record.status, Status_NeedsReview, "#T3"
    
    ' CLEANUP
    ' - clear both local fields to avoid leaving this workbook in an invalid state
    ' - this approach is not robust or ideal; but it solves the problem today
    record.cost_category = ""
    record.followup_question = ""

    Set GetTestErrors_BusinessExpense_Status = EB
End Function








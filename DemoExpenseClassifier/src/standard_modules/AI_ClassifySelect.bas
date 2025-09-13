Attribute VB_Name = "AI_ClassifySelect"
Option Explicit
Option Private Module
Private msg As String



Public Sub Classify_Expense_Selection()
    
    Dim selectionColumn As Range
    Set selectionColumn = Selection
    
    ' BYPASS if selection is not valid
    If ReturnTrueAndWarnUser_IfSelectionRangeInvalid(selectionColumn) Then
        Exit Sub
    End If

    Dim prompt As clsPrompt
    Set prompt = GetPrompt_ByName(PROMPT_NAME_EXPENSE_CLASSIFY_SELECT)
    
    Dim final_prompt As String
    final_prompt = GeneratePrompt_ExpenseClassify_Select(prompt.base_text, selectionColumn)
    
    ' Create the callback, and load the first argument
    Dim partial_vba_callback As clsPartialFunction
    Set partial_vba_callback = Create_PartialFunction(ThisWorkbook, prompt.vba_callback_name, selectionColumn)
    
    ' FORK based on whether to use the prompt interceptor
    Dim chatComp As ChatCompletionUDT
    If ShouldUsePromptInterceptorForm Then
        chatComp = OpenFormPromptInterceptor(ThisWorkbook, final_prompt, partial_vba_callback)
    Else
        chatComp = ExecutePrompt_WithCallback_VBA(final_prompt, partial_vba_callback)
    End If
    
    ' Update the prompt record with the latest
    If chatComp.totalTokens > 0 Then
        prompt.full_example_prompt = final_prompt
        prompt.full_example_response = chatComp.latestMessage
    End If

End Sub



Private Function GeneratePrompt_ExpenseClassify_Select(BaseText As String, selectionColumn As Range) As String

    ' Generate Choice Text
    Dim ChoiceRange As Range
    Set ChoiceRange = Range(RANGENAME_CHOICES_COST_CATEGORY)

    Dim ChoiceText As String
    ChoiceText = RangeToHorizontalString(ChoiceRange)
    
    ' Generate Expense Description Text
    Dim ExpenseText As String
    ExpenseText = RangeToVerticalString(selectionColumn)
    
    Dim FinalText As String
    FinalText = Replace(BaseText, REPLACE_TAG_CATEGORIES, ChoiceText)
    FinalText = Replace(FinalText, REPLACE_TAG_EXPENSES, ExpenseText)
    
    GeneratePrompt_ExpenseClassify_Select = FinalText
End Function




Private Function ReturnTrueAndWarnUser_IfSelectionRangeInvalid(selectionColumn As Range) As Boolean
    Const OWNER As String = "ReturnTrueAndWarnUser_IfSelectionRangeInvalid()"

    Dim reasons As New Collection
    Debug.Assert reasons.Count = 0
    
    ' CHECK at least two cells are selected
    If selectionColumn.Cells.Count < 2 Then
        reasons.Add "At least two cells must be selected"
    End If
    
    ' CHECK only ONE column is selected
    If selectionColumn.Columns.Count <> 1 Then
        reasons.Add "Exactly one column must be selected when running this procedure"
    End If
    
    ' CHECK all selected cells are populated
    If RangeCountPopulatedCells(selectionColumn) <> selectionColumn.Cells.Count Then
        reasons.Add "Every selected cell must be populated with an expense description"
    End If
    
    ' CHECK adjacent column is completely blank (to avoid accidental overwriting)
    Dim adjacentColumn As Range
    Set adjacentColumn = selectionColumn.Offset(0, 1)
    
    If Not RangeIsBlank(adjacentColumn) Then
        reasons.Add "The column adjacent-right of the current selection must be empty"
    End If
    
    ' RECONCILE
    Dim ProblemsExist As Boolean
    
    If reasons.Count = 0 Then
        ProblemsExist = False
    Else
        ProblemsExist = True
        
        Dim failure_reasons As String
        failure_reasons = CollectionToNumberedList(reasons)
        MsgBox failure_reasons, vbExclamation, OWNER
    End If

    ReturnTrueAndWarnUser_IfSelectionRangeInvalid = ProblemsExist
End Function




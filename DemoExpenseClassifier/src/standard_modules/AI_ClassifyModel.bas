Attribute VB_Name = "AI_ClassifyModel"
Option Explicit
Option Private Module
Private msg As String




Public Sub Classify_Expenses_NextTenRecords()
    Const OWNER As String = "Classify_Expenses_NextTenRecords()"
    
    Dim UnclassifiedExpenses As Collection
    Set UnclassifiedExpenses = AppCommon.GetBusinessExpenses_Unclassified
    
    Dim NextTenUnclassifiedExpenses As Collection
    Set NextTenUnclassifiedExpenses = SliceCollection(UnclassifiedExpenses, 1, 11)  ' half open interval (half-closed interval)
    Debug.Assert NextTenUnclassifiedExpenses.Count <= 10
    
    Dim Dict_IndexToId As New Dictionary
    Dim Dict_IndexToExpenseDescript As New Dictionary
    
    Dim record As clsBusinessExpense
    Dim index As Long
    
    For index = 1 To NextTenUnclassifiedExpenses.Count
        Set record = NextTenUnclassifiedExpenses(index)
        Dict_IndexToId.Add index, record.id
        Dict_IndexToExpenseDescript.Add index, record.description
    Next index

    Dim prompt As clsPrompt
    Set prompt = GetPrompt_ByName(PROMPT_NAME_EXPENSE_CLASSIFY_MODEL)
    
    Dim final_prompt As String
    final_prompt = GeneratePrompt_ExpenseClassify_Model(prompt.base_text, Dict_IndexToExpenseDescript)
    
    ' Create the callback, and load the first argument
    Dim partial_vba_callback As clsPartialFunction
    Set partial_vba_callback = Create_PartialFunction(ThisWorkbook, prompt.vba_callback_name, Dict_IndexToId)

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
    
    ' Reload DB complex / recalculate in order to be sure the formulas
    ' on the Home sheet are updated!
    Call WhoaReloadHostDatabaseComplex(ThisWorkbook, WhoaReloadFull, OWNER)
End Sub



Private Sub ClassifyBusinessExpense_FromJson(Dict_IndexToId As Dictionary, responseDict As Dictionary)

    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense
    
    Dim record As clsBusinessExpense
    Dim record_id As String
    
    Dim data_dict As Dictionary
    Dim index As Variant
    
    For Each index In responseDict
    
        record_id = Dict_IndexToId(CLng(index))
        Set data_dict = responseDict(CStr(index))
        
        ' Pragmatic sanity checks
        Debug.Assert data_dict.Exists("c")
        Debug.Assert data_dict.Exists("fq")
        
        Set record = DB.GetRecordById(record_id)
        Debug.Print "Updating next BusinessExpense record, id = " & record.id
        
        record.cost_category = data_dict("c")
        record.followup_question = data_dict("fq")
    Next index
    
End Sub




Private Function GeneratePrompt_ExpenseClassify_Model( _
    BaseText As String, _
    Dict_IndexToExpenseDescript As Dictionary _
) As String

    ' Generate Choice Text
    Dim ChoiceText As String
    ChoiceText = RangeToHorizontalString(Range(RANGENAME_CHOICES_COST_CATEGORY))

    ' Generate Expense Description Text
    Dim ExpenseText As String
    ExpenseText = DictionaryToString(Dict_IndexToExpenseDescript)

    Dim FinalText As String
    FinalText = Replace(BaseText, REPLACE_TAG_CATEGORIES, ChoiceText)
    FinalText = Replace(FinalText, REPLACE_TAG_EXPENSES, ExpenseText)

    GeneratePrompt_ExpenseClassify_Model = FinalText
End Function





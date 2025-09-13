Attribute VB_Name = "AppCommon"
Option Explicit
Option Private Module
Private msg As String


' ---------------- BUSINESS EXPENSE QUERIES -----------

Public Function GetBusinessExpenses_Classified() As Collection
    Const OWNER As String = "GetBusinessExpenses_Classified()"
    Dim Records As New Collection
    
    Dim record As clsBusinessExpense
    For Each record In GetDb_BusinessExpense.GetAllRecords
        If record.status = Status_Classified Then
            Records.Add record
        End If
        
    Next record
    
    Set GetBusinessExpenses_Classified = Records
End Function



Public Function GetBusinessExpenses_Unclassified() As Collection
    Const OWNER As String = "GetBusinessExpenses_Unclassified()"
    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense
    
    Dim Records As Collection
    Set Records = DB.GetRecordsByAttr("status", Status_Unclassified)
    
    Set GetBusinessExpenses_Unclassified = Records
End Function



Public Function GetBusinessExpenses_NeedReview() As Collection
    Set GetBusinessExpenses_NeedReview = GetDb_BusinessExpense.GetRecordsByAttr("status", Status_NeedsReview)
End Function





' ---------------- PROMPT QUERIES -----------
Public Function GetPrompt_ByName(prompt_name As String) As clsPrompt
    Set GetPrompt_ByName = GetDb_Prompt.GetRecordByIndex(prompt_name)
End Function



Public Function ShouldUsePromptInterceptorForm() As Boolean
    ShouldUsePromptInterceptorForm = WhoaGetSettingValue(ThisWorkbook, AI_SETTING_USE_INTERCEPTOR)
End Function



' ------- STRING UTILITIES ---------------

Public Function ExtractPromptReplaceTags(base_text As String) As Collection
    Set ExtractPromptReplaceTags = StringExtractEnclosedSubstrings( _
        search_text:=base_text, _
        openTag:="<", _
        closeTag:=">" _
    )
End Function


Public Function RangeToVerticalString(cellRange As Range) As String
    Dim Dict As New Dictionary
    
    Dim index As Long    ' Use Long instead of Integer always.
    Dim cell As Range
    
    For Each cell In cellRange
        index = index + 1
        Dict.Add index, cell.Value
    Next cell
    
    RangeToVerticalString = JsonConverter.ConvertToJson(Dict, 2)
End Function



Public Function RangeToHorizontalString(ColumnRange As Range) As String
    Const OWNER As String = "RangeToListString_Horizontal()"

    Dim Collect As New Collection
    Dim cell As Range
    Dim val As Variant

    For Each cell In ColumnRange
        val = cell.Value
        If IsEmpty(val) Then
            Collect.Add ""
        Else
            Collect.Add val
        End If
    Next cell

    RangeToHorizontalString = JsonConverter.ConvertToJson(Collect)
End Function












Attribute VB_Name = "HostConstructors"
' WARNING TREAT THIS CODE MODULE AS READ-ONLY (CODE ON THIS MODULE IS PERIODICALLY DELETED AND REBUILT)

Option Explicit
Option Private Module

Public Const MODEL_NAME_PROMPT As String = "Prompt"
Public Const MODEL_NAME_BUSINESS_EXPENSE As String = "BusinessExpense"


Public Function GetDb_Prompt() As clsDb
    Set GetDb_Prompt = GetDbOrRaise(ThisWorkbook, "Prompt")
End Function

Public Function GetDb_BusinessExpense() As clsDb
    Set GetDb_BusinessExpense = GetDbOrRaise(ThisWorkbook, "BusinessExpense")
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

Public Function GetPromptById(record_id As String, Optional always_false As Boolean = False) As clsPrompt
    Dim record As New clsPrompt
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetPromptById = record
    Else
        Set GetPromptById = Nothing
    End If
End Function

Public Function GetBusinessExpenseById(record_id As String, Optional always_false As Boolean = False) As clsBusinessExpense
    Dim record As New clsBusinessExpense
    record.Real_Initialize record_id
    If record.IsViableInstance Or always_false Then
        Set GetBusinessExpenseById = record
    Else
        Set GetBusinessExpenseById = Nothing
    End If
End Function

Private Function TestMacrosAreEnabled() As Boolean
    ' NOTE this function is called on Setting Registry sheet
    TestMacrosAreEnabled = True
End Function

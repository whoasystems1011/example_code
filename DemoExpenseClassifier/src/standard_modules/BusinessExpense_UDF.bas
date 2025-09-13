Attribute VB_Name = "BusinessExpense_UDF"
Option Explicit
'Option Private Module  ' Not here, because we want to see the functions at the worksheet level
Private msg As String



Public Function GetNumberExpenses_Total() As Long
    GetNumberExpenses_Total = GetDb_BusinessExpense.RecordCount
End Function


Public Function GetNumberExpenses_Classifed() As Long
    GetNumberExpenses_Classifed = GetBusinessExpenses_Classified.Count
End Function



Public Function GetNumberExpenses_Unclassified() As Long
    GetNumberExpenses_Unclassified = GetBusinessExpenses_Unclassified.Count
End Function



Public Function GetNumberExpenses_NeedReview() As Long
    GetNumberExpenses_NeedReview = GetBusinessExpenses_NeedReview.Count
End Function





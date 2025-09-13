Attribute VB_Name = "ResetWorkbook"
Option Explicit
Option Private Module
Private msg As String



Public Sub Reset_BusinessExpense_System()
    Call GetDb_BusinessExpense.DeleteAllRecordRows
End Sub

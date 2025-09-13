Attribute VB_Name = "AI_AgentTools"
Option Explicit
Option Private Module
Private msg As String



Public Sub PasteAnswersAdjacentRight(selectionColumn As Range, responseDict As Dictionary)

    Dim ai_category As String   ' Office Expense
    Dim expense_cell As Range
    Dim paste_cell As Range
    
    Dim index_key As Variant
    
    For Each index_key In responseDict
        ai_category = responseDict(index_key)
        
        Set expense_cell = selectionColumn(index_key)
        Set paste_cell = expense_cell.Offset(0, 1)
        
        paste_cell.Value = ai_category
        
    Next index_key
End Sub




Attribute VB_Name = "HostShortcuts"
Option Explicit


Public Function GetCustomWorkbookShortcuts() As Dictionary
    ' Registers custom keyboard shortcut under CTRL SHIFT Y (try that shortcut)
    ' Learn more here: https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey
    Dim D As New Dictionary
    
    D("^+Y") = "ExampleShortcutCallback"
    D("^+G") = "Classify_Expense_Selection"
    D("^+I") = "Import_BusinessExpense_Records"
    
    Set GetCustomWorkbookShortcuts = D
End Function


Private Sub ExampleShortcutCallback()
    MsgBox "Yellow World!", vbInformation, "This text will be title of messagebox"
End Sub

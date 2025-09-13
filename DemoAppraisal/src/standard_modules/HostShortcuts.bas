Attribute VB_Name = "HostShortcuts"
' NOTE when multiple WB_HOST workbooks are open, whichever was loaded more recently
' will prevail (in the event both specify the same shortcut pattern)
Option Explicit
Option Private Module


Public Function GetCustomWorkbookShortcuts() As Dictionary
    ' Registers custom keyboard shortcut under CTRL SHIFT Y (try that shortcut)
    ' Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey
    Dim D As New Dictionary
    
    D("^+Y") = "ExampleShortcutCallback"
    D("^+C") = "OpenForm_RecordList_MultifamilyRentComp"

    Set GetCustomWorkbookShortcuts = D
End Function




Private Sub ExampleShortcutCallback()
    MsgBox "Yellow World!", vbInformation, "This text will be title of messagebox"
End Sub



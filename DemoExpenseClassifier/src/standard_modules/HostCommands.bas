Attribute VB_Name = "HostCommands"
Option Explicit


Private Sub WhoaEvent_RegisterCustomCommands()
    Dim kword As String
    Dim callback As String
    Dim dtext As String
    Dim etext As String

    ' Register Example Command
    kword = "testhost"
    callback = "Command_ExampleHostCommand"
    dtext = "Testing out host command execution"
    etext = "`testhost`"
    RegisterCommandHOST ThisWorkbook, kword, callback, dtext, etext
    
    
    ' Register Classify Expenses Command
    kword = "classifyselect"
    callback = "Command_Classify_Selection"
    dtext = "Classifies the current selected column of expenses, outputs categories to the right."
    etext = "`classifyselect`"
    RegisterCommandHOST ThisWorkbook, kword, callback, dtext, etext
End Sub



Private Sub Command_Classify_Selection(UserCommand As String)
    Call Classify_Expense_Selection
End Sub



Private Sub Command_ExampleHostCommand(UserCommand As String)
    MsgBox "Command_ExampleHostCommand", vbInformation, "This text will be title of messagebox"
End Sub

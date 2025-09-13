Attribute VB_Name = "HostCommands"
Option Explicit
Option Private Module


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
End Sub



Private Sub Command_ExampleHostCommand(UserCommand As String)
    MsgBox "Command_ExampleHostCommand", vbInformation, "This text will be title of messagebox"
End Sub

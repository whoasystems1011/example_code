Attribute VB_Name = "HostEvents"
Option Explicit


Private Sub WhoaEvent_PostReloadWorkbook(load_type As WhoaDbLoadType)
    ' This procedure runs after the Reload Workbook button is clicked in the Whoa Excel Ribbon tab
    ' Specifically, this function is called AFTER the Addin has performed all of its own operations.
    Debug.Print "Welcome to a new Host Workbook"
    
    Call AppReload_ContextMenu
End Sub


Private Sub WhoaEvent_LoadCustomCache()
    Const OWNER As String = "WhoaEvent_LoadCustomCache()"
    LogInfo "Hitting Host Script: " & OWNER
End Sub





Attribute VB_Name = "AppContextMenu"
Option Explicit
Option Private Module
Private msg As String       ' NOTE only needed for P_RaiseIf_ContextPopup_AlreadyAdded()

Private Const POPUP_CAPTION As String = "AI Hacking"
Private Const BUTTON_CAPTION_1 As String = "Classify Expenses"
Private Const BUTTON_CAPTION_2 As String = "MessageBoxCurrentTime"


Public Sub AppReload_ContextMenu()
    ' @LESSON make it hard to break by ONLY having ONE public exposure
    Call P_CustomContextMenu_Remove
    Call P_CustomContextMenu_Add
End Sub



Private Sub P_CustomContextMenu_Add()
    Const OWNER As String = "P_CustomContextMenu_Add()"
    
    Dim cBar As CommandBar
    Set cBar = Application.CommandBars("Cell")

    Dim popup As CommandBarPopup
    Dim sub_button_1 As CommandBarButton
    Dim sub_button_2 As CommandBarButton

    Set popup = cBar.Controls.Add(msoControlPopup, temporary:=True)
    popup.Caption = POPUP_CAPTION
    
    Set sub_button_1 = popup.Controls.Add(msoControlButton)
    sub_button_1.Caption = BUTTON_CAPTION_1
    sub_button_1.OnAction = "Classify_Expense_Selection"
    
    Set sub_button_2 = popup.Controls.Add(msoControlButton)
    sub_button_2.Caption = BUTTON_CAPTION_2
    sub_button_2.OnAction = "MessageBoxCurrentTime"
    
    ' Optional Styling
    sub_button_1.FaceId = 59
    sub_button_1.Style = msoButtonIconAndCaption
End Sub



Private Sub P_CustomContextMenu_Remove()
    ' NOTE this suppresses runtime error 5, which occurs when
    ' we try to remove it when it doesn't exist
    Const OWNER As String = "P_CustomContextMenu_Remove()"
    
    On Error Resume Next
    Application.CommandBars("Cell").Controls(POPUP_CAPTION).Delete
    
    ' Any error besides runtime 5, which occurs when the commandbar has not been added at all,
    ' is suppressed.
    If Err.Number > 5 Then
        Err.Raise Err.Number, OWNER, Err.description
    End If
    
    On Error GoTo 0
End Sub



' ------------- EXAMPLE CALLBACK ---------------
Private Sub MessageBoxCurrentTime()
    MsgBox "The current time is: " & Format(Time, "hh:mm:ss AM/PM"), vbInformation, "Current Time"
End Sub

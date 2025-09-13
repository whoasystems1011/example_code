VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub CommandButton1_Click()
    Call OpenForm_PromptExplorer
End Sub


Private Sub CommandButton2_Click()
    Call OpenFormWhoaSettings_Custom(ThisWorkbook, SETTING_GROUP_AI)
End Sub


Private Sub CommandButton3_Click()
    Call Classify_Expenses_NextTenRecords
End Sub



Private Sub CommandButton4_Click()
    Dim UnclassifiedRecords As Collection
    Set UnclassifiedRecords = AppCommon.GetBusinessExpenses_Unclassified
    Call OpenFormGenericRecord_List(ThisWorkbook, MODEL_NAME_BUSINESS_EXPENSE)
End Sub


Attribute VB_Name = "BusinessExpense_Types"
Option Explicit


Public Enum BusinessExpenseStatus
    Status_Unknown = 0
    Status_Unclassified = 1
    Status_Classified = 2
    Status_NeedsReview = 3
End Enum



Public Function BusinessExpenseStatus_ToString(e As BusinessExpenseStatus) As String
    Dim output As String
    
    Select Case e
        Case Status_Unknown: output = "Status_Unknown"
        Case Status_Unclassified: output = "Status_Unclassified"
        Case Status_Classified: output = "Status_Classified"
        Case Status_NeedsReview: output = "Status_NeedsReview"
        Case Else: output = "INVALID ENUM VALUE"
    End Select

    BusinessExpenseStatus_ToString = output
End Function

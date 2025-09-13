Attribute VB_Name = "Reports_Static"
Option Explicit




Public Sub GenerateStaticReport()
    ' NOTE
    ' - we never implemented this sub
    ' - the point was to demonstrate how to create "static reports" where the table
    '   is generated with static values (as opposed to dynamic UDF formulas)
    
    ' INTENT
    ' - create a new worksheet and inject expense information for all
    '   expenses above some minimum threshold...

    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense
    
    ' NOTE
    ' - MIN_THRESHOLD_AMOUNT could become a Range input (or Whoa Setting)
    Const MIN_THRESHOLD_AMOUNT As Long = 500
    
    Dim record As clsBusinessExpense
    
    For Each record In DB.GetAllRecords
        If record.amount > 500 Then
            ' we would add code here
        End If
        
    Next record
End Sub

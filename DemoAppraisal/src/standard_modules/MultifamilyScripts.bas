Attribute VB_Name = "MultifamilyScripts"
Option Explicit
Option Private Module


Public Sub DesiginateAllPropertiesAsComps()
    ' NOTE
    ' - this procedure just changes every imported rent comp status to "comparable"
    Const new_status As String = "Excluded"

    Dim prop As clsMultifamilyRentComp
    
    For Each prop In GetDb_MultifamilyRentComp.GetAllRecords
        prop.status = new_status
        Debug.Print "Reclassified " & prop.name & " to status=" & new_status
    Next prop
End Sub


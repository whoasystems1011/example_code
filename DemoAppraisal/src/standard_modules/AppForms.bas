Attribute VB_Name = "AppForms"
Option Explicit
Option Private Module
Private msg As String


Public Sub OpenForm_RecordList_MultifamilyRentComp()
    ' NOTE
    ' - the goal here is just have ONE function that calls the addin
    ' - a good goal is to minimize direct calls to the addin, by wrapping them in your own function
    ' - if the called addin function is renamed (in a new Addin version release),
    '   you would only have to update one spot (this function)
    Call OpenFormGenericRecord_List(ThisWorkbook, MODEL_NAME_MULTIFAMILY_RENT_COMP)
End Sub

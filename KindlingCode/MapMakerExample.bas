' Worksheet vba module: MapMaker
Option Explicit


' ----------------- EXAMPLE A: MANDATORY -----------------
' - This code MUST get placed into your workbook (for MapMaker worksheet to work)
Private Sub CommandButton1_Click()
    Call GenerateKML_FromMapMaker
End Sub


Private Sub GenerateKML_FromMapMaker()
    ' Get the filename from the user (or pass in whatever you want)
    Dim FileName As String
    FileName = InputBox("Provide a name for new KML file")  ' This opens an InputBox
    
    ' This exports a Google Earth .KML file to the same directory as the HOST
    Call ExportMapKML_FromMapMaker(ThisWorkbook, FileName)
End Sub



' ----------------- EXAMPLE B: OPTIONAL/ADVANCED -----------------
' - Remaining code below shows how to manually generate a Google Earth KML from any context

Private Sub CommandButton2_Click()
    Call GenerateKML_Manually
End Sub


Private Sub GenerateKML_Manually()
    
    Const LONG_POINT_1 As Double = "-97.3229675292969"
    Const LAT_POINT_1 As Double = "30.1220397949219"
    
    Const LONG_POINT_2 As Double = "-97.3312911987305"
    Const LAT_POINT_2 As Double = "30.1097812652588"
    
    Dim map_pin_one As MapPin
    Dim map_pin_two As MapPin
    
    Set map_pin_one = Create_MapPin( _
        pin_title:="Pin ONE", _
        pin_color:="Red", _
        pin_hover_text:="Hover text for pin One", _
        longitude:=LONG_POINT_1, _
        latitude:=LAT_POINT_1 _
    )
    
    Set map_pin_two = Create_MapPin( _
        pin_title:="Pin TWO", _
        pin_color:="Blue", _
        pin_hover_text:="Hover text for pin TWO", _
        longitude:=LONG_POINT_2, _
        latitude:=LAT_POINT_2 _
    )
    
    ' Roll the into a Collection container
    Dim PinsCollect As New Collection
    PinsCollect.Add map_pin_one
    PinsCollect.Add map_pin_two
    
    ' Export a KML file
    Const TEST_FILE_NAME As String = "TwoRandomPoints"
    
    ' Export the KML file (it will have the two pins mapped)
    Call ExportMapKML_FromMapPins( _
        WB_HOST:=ThisWorkbook, _
        PinsCollect:=PinsCollect, _
        FileName:=TEST_FILE_NAME, _
        OpenExportFolder:=True _
    )
End Sub

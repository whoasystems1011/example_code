Attribute VB_Name = "MultifamilyConstants"
Option Explicit
Option Private Module
Private msg As String

Public Const MF_SETTING_LIHTCDB_USER_EMAIL As String = "LIHTCDB_USER_EMAIL"
Public Const MF_SETTING_LIHTCDB_USER_API_KEY As String = "LIHTCDB_USER_API_KEY"

Public Const MF_COLOR_SUBJECT As Long = vbBlue
Public Const MF_COLOR_COMP As Long = WHOA_COLOR_GREEN_STRONG
Public Const MF_COLOR_EXCLUDED As Long = WHOA_COLOR_RED_STRONG
Public Const MF_COLOR_NULL As Long = WHOA_COLOR_GRAY_DISABLED

Public Const LIHTC_ERROR_DEV As Long = WhoaError_LatestAssigned + 500


Public Function get_rent_comp_listview_color(prop_status As String) As Long
    ' NOTE
    ' - abstraction called from multiple db model class modules (properties, and child units)
    ' - this is a good example of DRY
    Const OWNER As String = "get_rent_comp_listview_color()"
    Dim RAISE_MISSING_STATUS As Boolean
    
    Dim color As Long
    color = PUBLIC_MISSING_INDEX_LONG

    Select Case prop_status
        Case "Subject": color = MF_COLOR_SUBJECT
        Case "Comparable": color = MF_COLOR_COMP
        Case "Excluded": color = MF_COLOR_EXCLUDED
        Case "Null": color = MF_COLOR_NULL
        Case "-": color = MF_COLOR_NULL     ' Dash can be same as null for now...
        Case Else: RAISE_MISSING_STATUS = True
    End Select
    
    If RAISE_MISSING_STATUS Then
        msg = OWNER & " unable to map property Status: " & PrettyValueType(prop_status)
        MsgBox msg, vbCritical, OWNER
        Err.Raise LIHTC_ERROR_DEV, OWNER, msg
    End If
    
    get_rent_comp_listview_color = color
End Function




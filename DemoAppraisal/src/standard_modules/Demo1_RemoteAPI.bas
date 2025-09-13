Attribute VB_Name = "Demo1_RemoteAPI"
Option Explicit
Option Private Module
Private msg As String

Const MARK_LIHTCDB_API_KEY As String = "hRsdfsdfsdfIAeo9AGT-zbFcmMI0="
Const MARK_EMAIL_ADDRESS As String = "redacted@email.com"


Private Sub PlayRemoteAPI_AMENITY()
    Const URL_AMENITY = "http://127.0.0.1:8000/api/multifamily/amenities/json/v1/"

    ' open POST connection
    Dim request As New WinHttpRequest
    request.Open "POST", URL_AMENITY, False
    
    ' attach credentials
    request.SetRequestHeader "Authorization", "Bearer " & MARK_LIHTCDB_API_KEY
    request.SetRequestHeader "X-User-Email", MARK_EMAIL_ADDRESS
    
    ' fire request
    Call request.Send
    
    ' display results
    Call PrettyPrintJson(request.ResponseText)
End Sub



Private Sub PlayRemoteAPI_RentPROPERTY()
    Dim property_rent_ids As New Collection
    property_rent_ids.Add "45304"
    property_rent_ids.Add "43230"

    Const URL_PROPERTY_LIST = "http://127.0.0.1:8000/api/multifamily/rent-comps/json/v1/"
    
    Dim PostDict As New Dictionary
    PostDict.Add "target_record_ids", property_rent_ids
    
    ' open POST connection
    Dim request As New WinHttpRequest
    request.Open "POST", URL_PROPERTY_LIST, False
    
    ' attach credentials
    request.SetRequestHeader "Authorization", "Bearer " & MARK_LIHTCDB_API_KEY
    request.SetRequestHeader "X-User-Email", MARK_EMAIL_ADDRESS
    
    Dim PostJson As String
    PostJson = JsonConverter.ConvertToJson(PostDict)
    
    ' fire request
    Call request.Send(PostJson)
    
    ' display result
    Call PrettyPrintJson(request.ResponseText)
End Sub



Private Sub PlayRemoteAPI_RentUNIT()
    ' Related unit mix: http://127.0.0.1:8000/multifamily/property/43230/
    
    Dim parent_property_ids As New Collection
    parent_property_ids.Add "45304"
    parent_property_ids.Add "43230"

    Const URL_PROPERTY_LIST = "http://127.0.0.1:8000/api/multifamily/rent-comp-units/json/v1/"
    
    Dim PostDict As New Dictionary
    PostDict.Add "parent_property_ids", parent_property_ids
    
    ' open POST connection
    Dim request As New WinHttpRequest
    request.Open "POST", URL_PROPERTY_LIST, False
    request.SetRequestHeader "Authorization", "Bearer " & MARK_LIHTCDB_API_KEY
    request.SetRequestHeader "X-User-Email", MARK_EMAIL_ADDRESS
    
    Dim PostJson As String
    PostJson = JsonConverter.ConvertToJson(PostDict)
    
    Call request.Send(PostJson)
    Call PrettyPrintJson(request.ResponseText)
End Sub



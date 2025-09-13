Attribute VB_Name = "MultifamilyApi"
Option Explicit
Option Private Module
Private msg As String


Private Function CreateConnectionLIHTC() As ConnectionDjango
    Dim UserApiKey As String
    Dim UserEmail As String
    
    Const BaseDomain As String = "http://127.0.0.1:8000"
    
    ' NOTE
    ' - the setting values are being fetched from the Whoa Setting registry sheet
    UserApiKey = WhoaGetSettingValue(ThisWorkbook, MF_SETTING_LIHTCDB_USER_API_KEY)
    UserEmail = WhoaGetSettingValue(ThisWorkbook, MF_SETTING_LIHTCDB_USER_EMAIL)
    
    Set CreateConnectionLIHTC = CreateConnectionDjango( _
        WB_HOST:=ThisWorkbook, _
        base_domain:=BaseDomain, _
        user_api_key:=UserApiKey, _
        user_email:=UserEmail _
    )
End Function


'----------------|
'--Data Queries--|
'----------------|
Public Function ApiQueryMultifamilyAmenityCollection() As Collection
    Dim connect As ConnectionDjango
    Set connect = CreateConnectionLIHTC
    Dim PostDict As New Dictionary
    Const relativeUrl As String = "/api/multifamily/amenities/json/v1/"
    Set ApiQueryMultifamilyAmenityCollection = connect.QueryDbForChunkedRecordCollect(relativeUrl, PostDict)
End Function


Public Function ApiQueryMultifamilyRentCompListsCollection() As Collection
    Dim connect As ConnectionDjango
    Set connect = CreateConnectionLIHTC()
    Dim PostDict As New Dictionary
    Const relativeUrl As String = "/api/multifamily/rent-comp-lists/json/v1/"
    Set ApiQueryMultifamilyRentCompListsCollection = connect.QueryDbForChunkedRecordCollect(relativeUrl, PostDict)
End Function


Public Function ApiQueryMultifamilyRentCompIdsByCompListIdCollection(comp_list_id As Variant) As Collection
    ' NOTE
    ' - this one is a intermediate black sheep HELPER which converts
    '   comp_list_id into property_pk_array

    ' step 1: traditional query
    Dim connect As ConnectionDjango
    Set connect = CreateConnectionLIHTC
    Dim PostDict As New Dictionary
    
    ' WARNING: must include trailing '/' (or else POST data is lost)
    Const relativeUrl = "/api/multifamily/rent-comp-ids-by-comp-list/json/v1/"
    PostDict.Add "target_complist_id", comp_list_id
    
    Dim item_dict As Dictionary
    Dim RawObjectsCollect As New Collection
    Dim FlatIdsCollect As New Collection
    
    Set RawObjectsCollect = connect.QueryDbForChunkedRecordCollect(relativeUrl, PostDict)
    
    ' step 2: flatten traditional query into single id list collection
    For Each item_dict In RawObjectsCollect
        FlatIdsCollect.Add item_dict("id")
    Next item_dict

    Set ApiQueryMultifamilyRentCompIdsByCompListIdCollection = FlatIdsCollect
End Function



Public Function ApiQueryMultifamilyRentCompsCollection(property_rent_ids As Collection) As Collection
    Dim connect As ConnectionDjango
    Set connect = CreateConnectionLIHTC
    Dim PostDict As New Dictionary
    
    ' WARNING: must include trailing '/' (or else POST data is lost due to REDIRECT)
    Const relativeUrl = "/api/multifamily/rent-comps/json/v1/"
    PostDict.Add "target_record_ids", property_rent_ids
    Set ApiQueryMultifamilyRentCompsCollection = connect.QueryDbForChunkedRecordCollect(relativeUrl, PostDict)
End Function


Public Function ApiQueryMultifamilyRentCompUnitsCollection(property_rent_ids As Collection) As Collection
    Dim connect As ConnectionDjango
    Set connect = CreateConnectionLIHTC
    Dim PostDict As New Dictionary
    Const relativeUrl As String = "/api/multifamily/rent-comp-units/json/v1/"
    PostDict.Add "parent_property_ids", property_rent_ids
    Set ApiQueryMultifamilyRentCompUnitsCollection = connect.QueryDbForChunkedRecordCollect(relativeUrl, PostDict)
End Function





'-----------------------|
'--LIHTCDB Pulse Check--|
'-----------------------|
Public Function RemoteLihtcDbServerRunning() As Boolean
    ' NOTE
    ' - this is a dodgy check-- depending on whether trailing slash is included
    '   onto relativeURL-- the server may return a 304.
    ' - all that really matters is we do not get back NULL_SENTINEL
    On Error GoTo FAIL_FAST
    Const OWNER As String = "RemoteLihtcDbServerRunning()"
    
    Const relativeUrl As String = "/pulse/"

    Dim connect As ConnectionDjango
    Set connect = CreateConnectionLIHTC()
    
    Const LOCAL_NULL_SENTINEL As Long = -999
    Const TIMEOUT_IN_MILLISECONDS As Double = 200
    
    Dim status_code As Long
    status_code = connect.QueryForHttpStatusCode( _
        relativeUrl:=relativeUrl, _
        timeout_ms:=TIMEOUT_IN_MILLISECONDS, _
        NULL_RETURN:=LOCAL_NULL_SENTINEL _
    )

    RemoteLihtcDbServerRunning = status_code <> LOCAL_NULL_SENTINEL
    Exit Function
    
FAIL_FAST:
    msg = PrettyErrorText(Err, OWNER, ExtraInfo:="Returning False to caller")
    MsgBox msg, vbExclamation, OWNER
    LogWarning msg
    RemoteLihtcDbServerRunning = False
End Function





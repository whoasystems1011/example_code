Attribute VB_Name = "MultifamilyImports"
Option Explicit
Option Private Module
Private msg As String

Private Const COMP_LIST_ID_BASTROP_TX As String = "72"
Private Const COMP_IDS_CSV_BASTROP_TX As String = "43168,43886,43230,45304,44418,43885,50815,50814,50816,50813"


'------------------------------|
'--Multifamily AMENITY Import--|
'------------------------------|
Public Sub RefreshRecords_MultifamilyAmenity()
    Dim DB As clsDb
    Set DB = GetDb_MultifamilyAmenity
    
    ' Query raw data from the webserver
    Dim AllRemoteRecords As Collection
    Set AllRemoteRecords = ApiQueryMultifamilyAmenityCollection
    
    ' Execute import operation
    Dim result As ResultIngestRecord
    result = DB.IngestExternalRecords(AllRemoteRecords)

    ' Display import summary
    Call ResultIngestRecord_DisplayInfo(result)
End Sub



'-----------------------------------------|
'--Multifamily PROPERTY Import (Round 1)--|
'-----------------------------------------|
'Public Sub RefreshRecords_MultifamilyRentComp_BastropTexas()
'
'    ' Get a contrived list of things
'    Dim PropertyIdsCollect As Collection
'    Set PropertyIdsCollect = CsvToCollection(COMP_IDS_CSV_BASTROP_TX)
'
'    ' Query raw data from the webserver
'    Dim WebRecords As Collection
'    Set WebRecords = ApiQueryMultifamilyRentCompsCollection(PropertyIdsCollect)
'
'    ' Execute import operation
'    Dim DB As clsDb
'    Set DB = GetDb_MultifamilyRentComp
'
'    Dim result As ResultIngestRecord
'    result = DB.IngestExternalRecords(WebRecords)
'
'    ' Display import summary
'    Call ResultIngestRecord_DisplayInfo(result)
'End Sub




'------------------------------------------------|
'--Multifamily PROPERTY & UNIT Import (Round 2)--|
'------------------------------------------------|


Public Sub RefreshRecords_MultifamilyRentComp_BastropTexas_WITH_UNITS()
    Dim PropertyIdsCollect As Collection
    Set PropertyIdsCollect = CsvToCollection(COMP_IDS_CSV_BASTROP_TX)

    Call P_ImportMultifamilyRentComps_And_Units(PropertyIdsCollect)
End Sub








Private Function P_ImportMultifamilyRentComps_And_Units(PropertyIds As Collection) As WhoaSimpleResult
    Const OWNER As String = "P_ImportMultifamilyRentComps_And_Units()"

    ' BYPASS / STOP the process if the server is offline
    If P_ReturnTrueAndWarnUserIf_LihtcDbServerOffline Then
        P_ImportMultifamilyRentComps_And_Units = wNO_ACTION
        Exit Function
    End If

    ' assume false before importing
    P_ImportMultifamilyRentComps_And_Units = wFAILURE

    ' WARNING
    ' - properties must be imported first, then units.
    Call P_ImportMultifamilyRentComps(PropertyIds)
    Call P_ImportMultifamilyRentCompUnits(PropertyIds)

    P_ImportMultifamilyRentComps_And_Units = wSUCCESS
End Function


Private Sub P_ImportMultifamilyRentComps(PropertyIds As Collection)
    ' PROPERTY-LEVEL import
    Dim RemoteRecords As Collection
    Set RemoteRecords = ApiQueryMultifamilyRentCompsCollection(PropertyIds)

    ' TRANSFORM fields
    ' - convert the "id" representation of state/county "str" representation ("43" -> "TX")
    ' - we don't actually DEFINE and "state" or "county" fields in the demo
    ' - but this shows how you would intercept a mapping

    Dim subdict As Dictionary
    For Each subdict In RemoteRecords
        subdict("state") = subdict("state_code")
        subdict("county") = subdict("county_name")
    Next subdict

    Dim DB_PROP As clsDb
    Set DB_PROP = GetDb_MultifamilyRentComp

    Dim result As ResultIngestRecord
    result = DB_PROP.IngestExternalRecords(WebRecords:=RemoteRecords)

    ' Temp debug print
    Debug.Print String(50, "-")
    Debug.Print ResultIngestRecord_ToString(result, "Import Summary: " & DB_PROP.model_name)
End Sub


Private Sub P_ImportMultifamilyRentCompUnits(PropertyIds As Collection)
    ' UNIT-LEVEL import

    ' Step 1: Cache all pre-existing units that are related to the set of parent properties being changed.
    Dim UnitIdsExisting As Collection
    Set UnitIdsExisting = P_GetPreExistingMultifamilyUnitIds(PropertyIds)

    ' Step 2: Query the latest set of unit records
    Dim WebRecords As Collection
    Set WebRecords = ApiQueryMultifamilyRentCompUnitsCollection(PropertyIds)

    ' Step 3: Fully insert (or rehydrate) the incoming records from the web
    Dim DB_UNITS As clsDb
    Set DB_UNITS = GetDb_MultifamilyRentCompUnit

    Dim result As ResultIngestRecord
    result = DB_UNITS.IngestExternalRecords(WebRecords, UnitIdsExisting)

    ' Temp debug print
    Debug.Print String(50, "-")
    Debug.Print ResultIngestRecord_ToString(result, "Import Summary: " & DB_UNITS.model_name)
End Sub


Private Function P_GetPreExistingMultifamilyUnitIds(PropertyIds As Collection) As Collection
    ' NOTE
    ' - helper function for the rent comp import
    ' - convert collection to dictionary for HASH LOOKUP speed (halo grenade physics!)

    Dim PropertyIdsDict As Dictionary
    Set PropertyIdsDict = CollectionToDictionary(PropertyIds)

    Dim UnitIdsExisting As New Collection
    Dim unit As clsMultifamilyRentCompUnit
    
    '
    For Each unit In GetDb_MultifamilyRentCompUnit.GetAllRecords
        If PropertyIdsDict.Exists(CStr(unit.parent_property.id)) Then
            UnitIdsExisting.Add unit.id
        End If
    Next unit

    Set P_GetPreExistingMultifamilyUnitIds = UnitIdsExisting
End Function



'---------------------------------|
'--Private (Server Health Check)--|
'---------------------------------|
Private Function P_ReturnTrueAndWarnUserIf_LihtcDbServerOffline() As Boolean
    Const OWNER As String = "P_ReturnTrueAndWarnUserIf_LihtcDbServerOffline()"
    
    ' BYPASS return if the server is online (implicit return False)
    If RemoteLihtcDbServerRunning Then
        Exit Function
    End If
    
    ' Craft warning message
    Dim SB As StringBuilder
    Set SB = Create_StringBuilder
    SB.Append "The LIHTC DB WebServer is not available. Try again later."
    MsgBox SB.Str, vbExclamation, OWNER
    
    ' return boolean flag
    P_ReturnTrueAndWarnUserIf_LihtcDbServerOffline = True
End Function





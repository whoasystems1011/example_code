VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultifamilyRentComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iDbRecord
Private GV_RECORD_ID As String
Private GV_MODEL_DB As clsDb
Private msg As String
Private msg_fail As String
Private Const GV_MODEL_NAME As String = "MultifamilyRentComp"


'-------------------------------------------|
'--Standard Model Interface (factory-code)--|
'-------------------------------------------|

Private Function iDbRecord_ConvertToConcrete() As Object
    Dim concrete As clsMultifamilyRentComp
    Set concrete = Me
    Set iDbRecord_ConvertToConcrete = concrete
End Function


Private Function iDbRecord_id() As Variant
    iDbRecord_id = GV_RECORD_ID
End Function


Private Function iDbRecord_GetWorkbook() As Workbook
    Set iDbRecord_GetWorkbook = ThisWorkbook
End Function


Private Function iDbRecord_ModelName() As String
    iDbRecord_ModelName = GV_MODEL_NAME
End Function


Private Function iDbRecord_SsnInstance(Optional attr_name As Variant = "") As String
    iDbRecord_SsnInstance = GenerateDbModelRecordSSN(GV_MODEL_NAME, GV_RECORD_ID, attr_name)
End Function


Private Function iDbRecord_GetAttr(attr_name As Variant, Optional NULL_RETURN As Variant) As Variant
    iDbRecord_GetAttr = GetDbRecordAttr(GV_MODEL_DB, Me, attr_name, NULL_RETURN)
End Function


Private Function iDbRecord_GetAttrObj(attr_name As Variant) As Object
    Set iDbRecord_GetAttrObj = CallByNameSafeObject(Me, attr_name, VbGet)
End Function


Private Function iDbRecord_HasAttr(attr_name As Variant) As Boolean
    iDbRecord_HasAttr = HasAttribute(obj:=Me, attr_name:=attr_name)
End Function


Private Sub iDbRecord_SetAttr(attr_name As Variant, newValue As Variant)
    GV_MODEL_DB.SaveLocalValueToDb GV_RECORD_ID, CStr(attr_name), newValue
End Sub


Private Function iDbRecord_ToDict(Optional EXCLUDE_FORMULA_COLUMNS As Boolean = False) As Dictionary
    Set iDbRecord_ToDict = GV_MODEL_DB.ConvertRecordToDict(Me, EXCLUDE_FORMULA_COLUMNS)
End Function


Public Function model_name() As String
    model_name = iDbRecord_ModelName
End Function


Public Function get_attr(attr_name As Variant, Optional NULL_RETURN As Variant) As Variant
    get_attr = iDbRecord_GetAttr(attr_name, NULL_RETURN)
End Function


Public Sub set_attr(attr_name As Variant, newValue As Variant)
    iDbRecord_SetAttr attr_name, newValue
End Sub


Public Function ssn_instance(Optional attr_name As Variant = "") As String
    ssn_instance = iDbRecord_SsnInstance(attr_name)
End Function


Public Function ToDict(Optional EXCLUDE_FORMULA_COLUMNS As Boolean = False) As Dictionary
    Set ToDict = iDbRecord_ToDict(EXCLUDE_FORMULA_COLUMNS)
End Function


Public Function ToJson() As String
    ToJson = JsonConverter.ConvertToJson(Me.ToDict)
End Function


Public Sub Real_Initialize(record_id As String)
    GV_RECORD_ID = record_id
    Set GV_MODEL_DB = GetDbOrRaise(iDbRecord_GetWorkbook, GV_MODEL_NAME)
End Sub


Public Function IsViableInstance() As Boolean
    IsViableInstance = GV_MODEL_DB.IdExists(GV_RECORD_ID)
End Function


Public Function get_field_choice_arr(attr_name As Variant) As Variant
    get_field_choice_arr = GV_MODEL_DB.GetFieldValidationChoiceArr(attr_name)
End Function



'--------------------------|
'--Getters (factory-code)--|
'--------------------------|

Public Property Get id() As Variant
    id = GV_RECORD_ID
End Property


Public Property Get name() As Variant
    name = iDbRecord_GetAttr("name")
End Property


Public Property Get program() As Variant
    program = iDbRecord_GetAttr("program")
End Property


Public Property Get tenancy() As Variant
    tenancy = iDbRecord_GetAttr("tenancy")
End Property


Public Property Get total_units() As Variant
    total_units = iDbRecord_GetAttr("total_units")
End Property


Public Property Get effective_rent_date() As Date
    effective_rent_date = iDbRecord_GetAttr("effective_rent_date")
End Property


Public Property Get street_address() As Variant
    street_address = iDbRecord_GetAttr("street_address")
End Property


Public Property Get city() As Variant
    city = iDbRecord_GetAttr("city")
End Property


Public Property Get latitude() As Variant
    latitude = iDbRecord_GetAttr("latitude")
End Property


Public Property Get longitude() As Variant
    longitude = iDbRecord_GetAttr("longitude")
End Property


Public Property Get edit_url_direct() As String
    edit_url_direct = iDbRecord_GetAttr("edit_url_direct")
End Property


Public Property Get amenities() As clsRelatedManager
    Set amenities = GV_MODEL_DB.GetRelatedManager(Me.id, "amenities")
End Property


Public Property Get name_length() As Variant
    name_length = iDbRecord_GetAttr("name_length")
End Property


Public Property Get status() As Variant
    status = iDbRecord_GetAttr("status")
End Property


'--------------------------|
'--Setters (factory-code)--|
'--------------------------|

Public Property Let id(newValue As Variant)
    MsgBoxExceptTesting "ID is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let name(newValue As Variant)
    MsgBoxExceptTesting "NAME is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let program(newValue As Variant)
    MsgBoxExceptTesting "PROGRAM is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let tenancy(newValue As Variant)
    MsgBoxExceptTesting "TENANCY is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let total_units(newValue As Variant)
    MsgBoxExceptTesting "TOTAL_UNITS is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let effective_rent_date(newValue As Date)
    MsgBoxExceptTesting "EFFECTIVE_RENT_DATE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let street_address(newValue As Variant)
    MsgBoxExceptTesting "STREET_ADDRESS is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let city(newValue As Variant)
    MsgBoxExceptTesting "CITY is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let latitude(newValue As Variant)
    MsgBoxExceptTesting "LATITUDE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let longitude(newValue As Variant)
    MsgBoxExceptTesting "LONGITUDE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let edit_url_direct(newValue As String)
    MsgBoxExceptTesting "EDIT_URL_DIRECT is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let amenities(newValue As clsRelatedManager)
    MsgBoxExceptTesting "AMENITIES is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let name_length(newValue As Variant)
    iDbRecord_SetAttr "name_length", newValue
End Property


Public Property Let status(newValue As Variant)
    iDbRecord_SetAttr "status", newValue
End Property




'------------------------------------------|
'--Universal Custom Extensions (Optional)--|
'------------------------------------------|

Public Function custom_display_repr() As String
    custom_display_repr = iDbRecord_SsnInstance
End Function


Public Function custom_get_listview_color() As Long
    custom_get_listview_color = get_rent_comp_listview_color(Me.status)
End Function


Public Function custom_get_web_edit_url() As String
    custom_get_web_edit_url = Me.edit_url_direct
End Function


'---------------------------|
'--Unique Business Methods--|
'---------------------------|
Public Function days_ago_surveyed() As Long
    ' Returns the number of days between today and the effective_rent_date
    ' Negative values mean the date is in the future
    ' REF https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datediff-function
    days_ago_surveyed = VBA.DateDiff("d", Me.effective_rent_date, VBA.Date)
End Function


Public Function associated_units() As Collection

    Dim UnitCollect As New Collection
    Dim unit As clsMultifamilyRentCompUnit
    
    For Each unit In GetDb_MultifamilyRentCompUnit.GetAllRecords
        If unit.parent_property.id = Me.id Then
            UnitCollect.Add unit
        End If
    Next unit

    Set associated_units = UnitCollect
End Function




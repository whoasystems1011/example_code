VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultifamilyRentCompUnit"
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
Private Const GV_MODEL_NAME As String = "MultifamilyRentCompUnit"


'-------------------------------------------|
'--Standard Model Interface (factory-code)--|
'-------------------------------------------|

Private Function iDbRecord_ConvertToConcrete() As Object
    Dim concrete As clsMultifamilyRentCompUnit
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


Public Property Get dummy_remote() As Variant
    dummy_remote = iDbRecord_GetAttr("dummy_remote")
End Property


Public Property Get beds() As Variant
    beds = iDbRecord_GetAttr("beds")
End Property


Public Property Get baths() As Variant
    baths = iDbRecord_GetAttr("baths")
End Property


Public Property Get square_feet() As Long
    square_feet = iDbRecord_GetAttr("square_feet")
End Property


Public Property Get rent_restriction_ami() As String
    rent_restriction_ami = iDbRecord_GetAttr("rent_restriction_ami")
End Property


Public Property Get rent_restriction_pbra() As Variant
    rent_restriction_pbra = iDbRecord_GetAttr("rent_restriction_pbra")
End Property


Public Property Get tenancy() As Variant
    tenancy = iDbRecord_GetAttr("tenancy")
End Property


Public Property Get quantity() As Variant
    quantity = iDbRecord_GetAttr("quantity")
End Property


Public Property Get quantity_vacant() As Variant
    quantity_vacant = iDbRecord_GetAttr("quantity_vacant")
End Property


Public Property Get dynamic_design_type() As Variant
    dynamic_design_type = iDbRecord_GetAttr("dynamic_design_type")
End Property


Public Property Get dynamic_number_stories() As Variant
    dynamic_number_stories = iDbRecord_GetAttr("dynamic_number_stories")
End Property


Public Property Get rent_asking_net() As Variant
    rent_asking_net = iDbRecord_GetAttr("rent_asking_net")
End Property


Public Property Get concession() As Variant
    concession = iDbRecord_GetAttr("concession")
End Property


Public Property Get getting_max_legal_rent() As Variant
    getting_max_legal_rent = iDbRecord_GetAttr("getting_max_legal_rent")
End Property


Public Property Get is_subsidized() As Variant
    is_subsidized = iDbRecord_GetAttr("is_subsidized")
End Property


Public Property Get has_waiting_list() As Variant
    has_waiting_list = iDbRecord_GetAttr("has_waiting_list")
End Property


Public Property Get remarks() As Variant
    remarks = iDbRecord_GetAttr("remarks")
End Property


Public Property Get vacancy_rate() As Variant
    vacancy_rate = iDbRecord_GetAttr("vacancy_rate")
End Property


Public Property Get unit_type() As String
    unit_type = iDbRecord_GetAttr("unit_type")
End Property


Public Property Get unit_program() As String
    unit_program = iDbRecord_GetAttr("unit_program")
End Property


Public Property Get rpsf_asking_net() As Variant
    rpsf_asking_net = iDbRecord_GetAttr("rpsf_asking_net")
End Property


Public Property Get parent_property() As clsMultifamilyRentComp
    Set parent_property = GetDbRecordById(iDbRecord_GetWorkbook, "MultifamilyRentComp", iDbRecord_GetAttr("parent_property"))
End Property


Public Property Get dummy_local() As Variant
    dummy_local = iDbRecord_GetAttr("dummy_local")
End Property


'--------------------------|
'--Setters (factory-code)--|
'--------------------------|

Public Property Let id(newValue As Variant)
    MsgBoxExceptTesting "ID is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let dummy_remote(newValue As Variant)
    MsgBoxExceptTesting "DUMMY_REMOTE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let beds(newValue As Variant)
    MsgBoxExceptTesting "BEDS is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let baths(newValue As Variant)
    MsgBoxExceptTesting "BATHS is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let square_feet(newValue As Long)
    MsgBoxExceptTesting "SQUARE_FEET is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let rent_restriction_ami(newValue As String)
    MsgBoxExceptTesting "RENT_RESTRICTION_AMI is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let rent_restriction_pbra(newValue As Variant)
    MsgBoxExceptTesting "RENT_RESTRICTION_PBRA is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let tenancy(newValue As Variant)
    MsgBoxExceptTesting "TENANCY is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let quantity(newValue As Variant)
    MsgBoxExceptTesting "QUANTITY is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let quantity_vacant(newValue As Variant)
    MsgBoxExceptTesting "QUANTITY_VACANT is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let dynamic_design_type(newValue As Variant)
    MsgBoxExceptTesting "DYNAMIC_DESIGN_TYPE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let dynamic_number_stories(newValue As Variant)
    MsgBoxExceptTesting "DYNAMIC_NUMBER_STORIES is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let rent_asking_net(newValue As Variant)
    MsgBoxExceptTesting "RENT_ASKING_NET is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let concession(newValue As Variant)
    MsgBoxExceptTesting "CONCESSION is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let getting_max_legal_rent(newValue As Variant)
    MsgBoxExceptTesting "GETTING_MAX_LEGAL_RENT is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let is_subsidized(newValue As Variant)
    MsgBoxExceptTesting "IS_SUBSIDIZED is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let has_waiting_list(newValue As Variant)
    MsgBoxExceptTesting "HAS_WAITING_LIST is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let remarks(newValue As Variant)
    MsgBoxExceptTesting "REMARKS is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let vacancy_rate(newValue As Variant)
    MsgBoxExceptTesting "VACANCY_RATE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let unit_type(newValue As String)
    MsgBoxExceptTesting "UNIT_TYPE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let unit_program(newValue As String)
    MsgBoxExceptTesting "UNIT_PROGRAM is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let rpsf_asking_net(newValue As Variant)
    MsgBoxExceptTesting "RPSF_ASKING_NET is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let parent_property(newValue As clsMultifamilyRentComp)
    MsgBoxExceptTesting "PARENT_PROPERTY is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let dummy_local(newValue As Variant)
    iDbRecord_SetAttr "dummy_local", newValue
End Property




'------------------------------------------|
'--Universal Custom Extensions (Optional)--|
'------------------------------------------|

Public Function custom_display_repr() As String
    custom_display_repr = iDbRecord_SsnInstance
End Function


Public Function custom_get_listview_color() As Long
    Const OWNER As String = "custom_get_listview_color"
    Debug.Assert StringContains(OWNER, WHOA_DBRECORD_CB_LISTVIEW_COLOR)
    custom_get_listview_color = PUBLIC_MISSING_INDEX_LONG
End Function


Public Function custom_get_web_edit_url() As String
    Const OWNER As String = "custom_get_web_edit_url()"
    Debug.Assert StringContains(OWNER, WHOA_DBRECORD_CB_WEB_EDIT_URL)
    MsgBox "Model " & WrapAsterisk(GV_MODEL_NAME) & " has not defined any behavior for: " & OWNER
    
    custom_get_web_edit_url = ""
End Function


'---------------------------|
'--Unique Business Methods--|
'---------------------------|



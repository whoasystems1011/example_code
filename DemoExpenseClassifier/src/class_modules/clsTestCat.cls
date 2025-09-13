VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestCat"
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
Private Const GV_MODEL_NAME As String = "TestCat"


'-------------------------------------------|
'--Standard Model Interface (factory-code)--|
'-------------------------------------------|

Private Function iDbRecord_ConvertToConcrete() As Object
    Dim concrete As clsTestCat
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


Public Property Get score() As Variant
    score = iDbRecord_GetAttr("score")
End Property


Public Property Get dummy_text_field() As String
    dummy_text_field = iDbRecord_GetAttr("dummy_text_field")
End Property


Public Property Get remote_height() As Variant
    remote_height = iDbRecord_GetAttr("remote_height")
End Property


Public Property Get is_alive() As Boolean
    is_alive = iDbRecord_GetAttr("is_alive")
End Property


Public Property Get birthday() As Date
    birthday = iDbRecord_GetAttr("birthday")
End Property


Public Property Get years_old() As Long
    years_old = iDbRecord_GetAttr("years_old")
End Property


Public Property Get name() As String
    name = iDbRecord_GetAttr("name")
End Property


Public Property Get gender() As Variant
    gender = iDbRecord_GetAttr("gender")
End Property


Public Property Get color() As Variant
    color = iDbRecord_GetAttr("color")
End Property


Public Property Get number_days_alive() As Long
    number_days_alive = iDbRecord_GetAttr("number_days_alive")
End Property


Public Property Get number_nipples() As Long
    number_nipples = iDbRecord_GetAttr("number_nipples")
End Property


Public Property Get remote_height_2x() As Long
    remote_height_2x = iDbRecord_GetAttr("remote_height_2x")
End Property


Public Property Get record_number() As Long
    record_number = iDbRecord_GetAttr("record_number")
End Property


Public Property Get id_plus_two() As Long
    id_plus_two = iDbRecord_GetAttr("id_plus_two")
End Property


Public Property Get shelter() As clsShelter
    Set shelter = GetDbRecordById(iDbRecord_GetWorkbook, "Shelter", iDbRecord_GetAttr("shelter"))
End Property


Public Property Get fav_sheets() As clsRelatedManager
    Set fav_sheets = GV_MODEL_DB.GetRelatedManager(Me.id, "fav_sheets")
End Property


'--------------------------|
'--Setters (factory-code)--|
'--------------------------|

Public Property Let id(newValue As Variant)
    MsgBoxExceptTesting "ID is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let score(newValue As Variant)
    MsgBoxExceptTesting "SCORE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let dummy_text_field(newValue As String)
    MsgBoxExceptTesting "DUMMY_TEXT_FIELD is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let remote_height(newValue As Variant)
    MsgBoxExceptTesting "REMOTE_HEIGHT is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let is_alive(newValue As Boolean)
    iDbRecord_SetAttr "is_alive", newValue
End Property


Public Property Let birthday(newValue As Date)
    iDbRecord_SetAttr "birthday", newValue
End Property


Public Property Let years_old(newValue As Long)
    iDbRecord_SetAttr "years_old", newValue
End Property


Public Property Let name(newValue As String)
    iDbRecord_SetAttr "name", newValue
End Property


Public Property Let gender(newValue As Variant)
    iDbRecord_SetAttr "gender", newValue
End Property


Public Property Let color(newValue As Variant)
    iDbRecord_SetAttr "color", newValue
End Property


Public Property Let number_days_alive(newValue As Long)
    iDbRecord_SetAttr "number_days_alive", newValue
End Property


Public Property Let number_nipples(newValue As Long)
    iDbRecord_SetAttr "number_nipples", newValue
End Property


Public Property Let remote_height_2x(newValue As Long)
    iDbRecord_SetAttr "remote_height_2x", newValue
End Property


Public Property Let record_number(newValue As Long)
    iDbRecord_SetAttr "record_number", newValue
End Property


Public Property Let id_plus_two(newValue As Long)
    iDbRecord_SetAttr "id_plus_two", newValue
End Property


Public Property Set shelter(related_record As clsShelter)
    Me.set_attr "shelter", CallByNameSafe(related_record, "id", VbGet, "")
End Property







'------------------------------------------|
'--Universal Custom Extensions (Optional)--|
'------------------------------------------|

Public Function custom_display_repr() As String
    custom_display_repr = iDbRecord_SsnInstance
End Function


Public Function custom_get_listview_color() As Long
    custom_get_listview_color = PUBLIC_MISSING_INDEX_LONG
End Function


'---------------------------|
'--Unique Business Methods--|
'---------------------------|
Public Function ReturnDictionary() As Dictionary
    ' This function is used for internal testing purposes
    ' It only needs to be defined on the TestCat model's class module
    Set ReturnDictionary = CreateDictionary
End Function

Public Function RaiseErrorForTesting() As Long
    ' This function is used for internal testing purposes
    ' It only needs to be defined on the TestCat model's class module
    Err.Raise WhoaErrIntentional, "clsTestCat.RaiseErrorForTesting()", "Raising an error for unit test purposes"
    RaiseErrorForTesting = 123
End Function



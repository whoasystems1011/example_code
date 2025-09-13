VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetInfo"
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
Private Const GV_MODEL_NAME As String = "SheetInfo"


'-------------------------------------------|
'--Standard Model Interface (factory-code)--|
'-------------------------------------------|

Private Function iDbRecord_ConvertToConcrete() As Object
    Dim concrete As clsSheetInfo
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


Public Property Get codename_test() As Variant
    codename_test = iDbRecord_GetAttr("codename_test")
End Property


Public Property Get codename() As Variant
    codename = iDbRecord_GetAttr("codename")
End Property


Public Property Get name() As String
    name = iDbRecord_GetAttr("name")
End Property


Public Property Get groups() As Variant
    groups = iDbRecord_GetAttr("groups")
End Property


Public Property Get notes() As Variant
    notes = iDbRecord_GetAttr("notes")
End Property


Public Property Get todos() As Variant
    todos = iDbRecord_GetAttr("todos")
End Property


'--------------------------|
'--Setters (factory-code)--|
'--------------------------|

Public Property Let id(newValue As Variant)
    MsgBoxExceptTesting "ID is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let codename_test(newValue As Variant)
    MsgBoxExceptTesting "CODENAME_TEST is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let codename(newValue As Variant)
    iDbRecord_SetAttr "codename", newValue
End Property


Public Property Let name(newValue As String)
    iDbRecord_SetAttr "name", newValue
End Property


Public Property Let groups(newValue As Variant)
    iDbRecord_SetAttr "groups", newValue
End Property


Public Property Let notes(newValue As Variant)
    iDbRecord_SetAttr "notes", newValue
End Property


Public Property Let todos(newValue As Variant)
    iDbRecord_SetAttr "todos", newValue
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


'---------------------------|
'--Unique Business Methods--|
'---------------------------|



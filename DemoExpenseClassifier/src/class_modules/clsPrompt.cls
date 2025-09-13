VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrompt"
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
Private Const GV_MODEL_NAME As String = "Prompt"


'-------------------------------------------|
'--Standard Model Interface (factory-code)--|
'-------------------------------------------|

Private Function iDbRecord_ConvertToConcrete() As Object
    Dim concrete As clsPrompt
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


Public Property Get dummy_local() As Variant
    dummy_local = iDbRecord_GetAttr("dummy_local")
End Property


Public Property Get name() As String
    name = iDbRecord_GetAttr("name")
End Property


Public Property Get vba_callback_name() As String
    vba_callback_name = iDbRecord_GetAttr("vba_callback_name")
End Property


Public Property Get base_text() As String
    base_text = iDbRecord_GetAttr("base_text")
End Property


Public Property Get full_example_prompt() As Variant
    full_example_prompt = iDbRecord_GetAttr("full_example_prompt")
End Property


Public Property Get full_example_response() As Variant
    full_example_response = iDbRecord_GetAttr("full_example_response")
End Property


Public Property Get notes() As String
    notes = iDbRecord_GetAttr("notes")
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


Public Property Let dummy_local(newValue As Variant)
    iDbRecord_SetAttr "dummy_local", newValue
End Property


Public Property Let name(newValue As String)
    iDbRecord_SetAttr "name", newValue
End Property


Public Property Let vba_callback_name(newValue As String)
    iDbRecord_SetAttr "vba_callback_name", newValue
End Property


Public Property Let base_text(newValue As String)
    iDbRecord_SetAttr "base_text", newValue
End Property


Public Property Let full_example_prompt(newValue As Variant)
    iDbRecord_SetAttr "full_example_prompt", newValue
End Property


Public Property Let full_example_response(newValue As Variant)
    iDbRecord_SetAttr "full_example_response", newValue
End Property


Public Property Let notes(newValue As String)
    iDbRecord_SetAttr "notes", newValue
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


Public Sub custom_callback_edit_form(listview_form As UserForm, Optional unload_callback As clsPartialFunction)
    Unload listview_form
    Call AppForms.OpenForm_PromptExplorer(Me.name, unload_callback)
End Sub


'---------------------------|
'--Unique Business Methods--|
'---------------------------|
Public Function GetTemplateReplaceTags() As Collection
    Set GetTemplateReplaceTags = ExtractPromptReplaceTags(Me.base_text)
End Function


Public Function GetValidateErrors() As ErrorBank
    Const OWNER As String = "clsPrompt.GetValidateErrors()"
    Dim EB As ErrorBank
    Set EB = CreateErrorBank(OWNER)
    
    msg_fail = "#T1 no <REPLACE_TAGS> found in prompt base text: " & Me.name
    EB.AssertTrue Me.GetTemplateReplaceTags.Count > 0, msg_fail
    
    If Me.vba_callback_name <> "" Then
        msg_fail = "#T2 prompt.vba_callback_name must correspond to a defined procedure: "
        EB.AssertVbaProcedureExists ThisWorkbook, Me.vba_callback_name, msg_fail
    End If
    
    Set GetValidateErrors = EB
End Function




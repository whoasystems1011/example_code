VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBusinessExpense"
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
Private Const GV_MODEL_NAME As String = "BusinessExpense"


'-------------------------------------------|
'--Standard Model Interface (factory-code)--|
'-------------------------------------------|

Private Function iDbRecord_ConvertToConcrete() As Object
    Dim concrete As clsBusinessExpense
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


Public Property Get transact_date() As Variant
    transact_date = iDbRecord_GetAttr("transact_date")
End Property


Public Property Get vendor_name() As Variant
    vendor_name = iDbRecord_GetAttr("vendor_name")
End Property


Public Property Get description() As String
    description = iDbRecord_GetAttr("description")
End Property


Public Property Get amount() As Variant
    amount = iDbRecord_GetAttr("amount")
End Property


Public Property Get cost_category() As Variant
    cost_category = iDbRecord_GetAttr("cost_category")
End Property


Public Property Get followup_question() As Variant
    followup_question = iDbRecord_GetAttr("followup_question")
End Property


'--------------------------|
'--Setters (factory-code)--|
'--------------------------|

Public Property Let id(newValue As Variant)
    MsgBoxExceptTesting "ID is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let transact_date(newValue As Variant)
    MsgBoxExceptTesting "TRANSACT_DATE is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let vendor_name(newValue As Variant)
    MsgBoxExceptTesting "VENDOR_NAME is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let description(newValue As String)
    MsgBoxExceptTesting "DESCRIPTION is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let amount(newValue As Variant)
    MsgBoxExceptTesting "AMOUNT is a REMOTE field and cannot be changed with VBA. Change value in external source (database), then resync (reimport) to the excel workbook."
End Property


Public Property Let cost_category(newValue As Variant)
    iDbRecord_SetAttr "cost_category", newValue
End Property


Public Property Let followup_question(newValue As Variant)
    iDbRecord_SetAttr "followup_question", newValue
End Property




'------------------------------------------|
'--Universal Custom Extensions (Optional)--|
'------------------------------------------|

Public Function custom_display_repr() As String
    custom_display_repr = iDbRecord_SsnInstance
End Function


Public Function custom_get_listview_color() As Long
    Const HORRID_COLOR As Long = 39880  ' RGB(200, 155, 0)
    Dim output As Long
    Select Case Me.status
        Case Status_Unclassified: output = vbBlue
        Case Status_NeedsReview: output = WHOA_COLOR_RED_STRONG
        Case Status_Classified: output = vbBlack
        Case Else: output = vbCyan   ' NOTE Cyan means indeterminate state (which is "theoretically" impossible)
    End Select
    
    custom_get_listview_color = output
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
Public Function status() As BusinessExpenseStatus
    Dim output As BusinessExpenseStatus
    
    Select Case Me.cost_category
        Case "": output = Status_Unclassified
        Case COST_CATEGORY_NEEDS_REVIEW: output = Status_NeedsReview
        Case Else: output = Status_Classified
    End Select
    status = output
End Function


Public Function status_str() As String
    status_str = BusinessExpenseStatus_ToString(Me.status)
End Function


Public Function GetValidateErrors() As ErrorBank
    Const OWNER As String = "clsBusinessExpense.GetValidateErrors()"
    Dim EB As ErrorBank
    Set EB = CreateErrorBank(OWNER)
    
    If Me.cost_category = COST_CATEGORY_NEEDS_REVIEW Then
    
        msg_fail = "#T1 followup question cannot be blank for expenses classified as: " & COST_CATEGORY_NEEDS_REVIEW
        EB.AssertTruthy Me.followup_question, msg_fail
        
    ElseIf Me.followup_question <> "" Then
        msg_fail = "#T2 followup question should be blank unless expense is classified as: " & COST_CATEGORY_NEEDS_REVIEW
        EB.AssertEqual Me.cost_category, COST_CATEGORY_NEEDS_REVIEW, msg_fail
    End If
    
    Set GetValidateErrors = EB
End Function




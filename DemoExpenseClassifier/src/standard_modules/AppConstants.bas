Attribute VB_Name = "AppConstants"
Option Explicit
Option Private Module
Private msg As String


' Old World Part 1
Public Const RANGENAME_CHOICES_COST_CATEGORY As String = "CHOICES_COST_CATEGORY"

Public Const REPLACE_TAG_CATEGORIES As String = "<ALLOWED_CLASSIFICATION_CATEGORIES>"
Public Const REPLACE_TAG_EXPENSES As String = "<EXPENSE_LINE_ITEMS>"


' New World Part 2
'Public Const MODEL_NAME_PROMPT As String = "Prompt"
'Public Const MODEL_NAME_BUSINESS_EXPENSE As String = "BusinessExpense"

' New World Part 3
' Dont have to do any model name creation
' Dont have to do any Get_DBModel

' PROMPT names
Public Const PROMPT_NAME_EXPENSE_CLASSIFY_SELECT As String = "ExpenseClassifierSelection"
Public Const PROMPT_NAME_EXPENSE_CLASSIFY_MODEL As String = "ExpenseClassifierModel"


' SETTINGS
Public Const SETTING_GROUP_AI As String = "AI"
Public Const AI_SETTING_USE_INTERCEPTOR As String = "AI_INTERCEPT_OUTGOING_PROMPTS"


Public Const COST_CATEGORY_NEEDS_REVIEW As String = "Needs Context"
Public Const COST_CATEGORY_SUPPLIES As String = "Office Supplies"


Public Const COOL_COLOR_BLUE As Long = 14766623


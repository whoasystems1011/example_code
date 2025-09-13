Attribute VB_Name = "AA_Demo"
Option Explicit
'Option Private Module
Private msg As String
Private msg_fail As String


Private Const P_NUMBER_OF_SHEETS_TO_ADD As Long = 2


Private Sub LoopExpenses_Video_WhyVba()
    Dim expense As clsBusinessExpense
    For Each expense In GetDb_BusinessExpense.GetAllRecords
        Debug.Print expense.description & expense.cost_category
    Next expense
    Debug.Print String(50, "-")
End Sub



Private Sub Demo_CustomWhoaSettings()

    ' Read the current setting value
    Dim result As Variant
    result = WhoaGetSettingValue(ThisWorkbook, "AI_EXAMPLE_SETTING")
    MsgBox result
    
    ' Change the setting value
    Call WhoaChangeSettingValue(ThisWorkbook, "AI_EXAMPLE_SETTING", "Parrots")
    
End Sub


' UDF (USER DEFINED FUNCTION)
Public Function MultiplyByThree(some_number As Variant) As Long
    MultiplyByThree = some_number * 3
End Function



'' UDF (USER DEFINED FUNCTION)
'Public Function MultiplyByThree(some_number As Variant)
'    Const OWNER As String = "MultiplyByThree()"
'
'    On Error GoTo EH
'    MultiplyByThree = some_number * 3
'
'    Exit Function
'EH:
'    msg = "RUNTIME ERR: " & Err.Number
'    MultiplyByThree = msg
'End Function










' VARIABLES
Private Sub DEMO_VARIABLES()
    
    Const MY_CONSTANT As String = "I am a constant"
    
    ' SCALER VARIABLE TYPES
    ' - they do NOT need to use the 'Set' operator
    Dim MY_STRING As String
    MY_STRING = "I am a String"
    
    Dim MyTrueFalse As Boolean
    MyTrueFalse = True

    ' REMEMBER - do not use Integer type, use Long instead
    'Dim MyNumberBad As Integer
    'MyNumberBad = 900000
    Dim MyNumberGood As Long
    MyNumberGood = 900000

    ' OBJECT TYPES
    Dim Dict As Dictionary
    Set Dict = New Dictionary

    Dim Collect As Collection
    Set Collect = New Collection

    MsgBox "Done! No errors"
End Sub




' PROCEDURES
Public Sub MySubroutine()
    Dim some_number As Long
    some_number = MyFunction
    
    MsgBox some_number
End Sub



Public Function MyFunction() As Long
    MyFunction = 23
End Function



















Private Sub Demo_JSON_Serialize()
    ' https://www.forexfactory.com/calendar/graph/142633?limit=5&site_id=1
    
    ' Load a dictionary with arbitrary counts of animals
    Dim Dict_AnimalCounts As New Dictionary
    
    Dict_AnimalCounts.Add "Cats", 25
    Dict_AnimalCounts.Add "Alligators", 16
    Dict_AnimalCounts.Add "Crocodiles", 11
    Dict_AnimalCounts.Add "Frogs", 15
    
    ' Convert the dictionary to JSON
    Dim JsonStr As String
    JsonStr = JsonConverter.ConvertToJson(Dict_AnimalCounts, 2)
    Debug.Print JsonStr
    
End Sub



Private Sub Demo_JSON_Deserialize()
    Call GetScratchSheet_SCRUBBED(ThisWorkbook)

    Dim JsonStr As String
    JsonStr = Range("ZZZ_JSON_STRING_CELL").Value
    
    Debug.Print "Raw JSON Text (Read from Worksheet Cell Range)"
    Debug.Print JsonStr
    
    Dim Dict_AnimalCounts As Dictionary
    Set Dict_AnimalCounts = JsonConverter.ParseJson(JsonStr)
    
    Call PrettyPrintDictionary(Dict_AnimalCounts, "ANIMAL COUNTS DICTIONARY")
    
End Sub




' DATABASE RECORDS


Private Sub DemoDbRecords()
    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense
    
    Dim record As clsBusinessExpense
    ' Loop each record and print description
    For Each record In DB.GetAllRecords
        Debug.Print record.description
        'record.description = "DONKEYS"   ' <<< this would break
    Next record
End Sub





Private Sub Demo_DB_Ingest_RemoteRecords()

    Const OWNER As String = "Demo_DB_Ingest_RemoteRecords()"
    Dim DB_EXPENSE As clsDb
    Set DB_EXPENSE = GetDb_BusinessExpense
    
    Dim D1 As New Dictionary
    
    D1.Add "id", 1
    D1.Add "amount", 9999
    D1.Add "description", "Crispy LLM Fajitas"
    D1.Add "vendor_name", "Mo's Taco Truck"
    D1.Add "transact_date", DateSerial(2015, 1, 21)
    
    Dim D2 As New Dictionary
    
    D2.Add "id", 4      ' REMEMBER try changing this to 3
    D2.Add "amount", 101
    D2.Add "description", "Crispy AMERICA!!!!!!!!!!!!"
    D2.Add "vendor_name", "Mo's Taco Truck!!!!!!!!!!!!!!"
    D2.Add "transact_date", DateSerial(2015, 1, 21)
    
    Call PrettyPrintDictionary(D2, "FAKE EXPENSE RECORD: #2")
    
    Dim WebRecords As New Collection
    WebRecords.Add D1
    WebRecords.Add D2
    
    Dim result As ResultIngestRecord
    result = DB_EXPENSE.IngestExternalRecords(WebRecords)
    
    Call ResultIngestRecord_DisplayInfo(result)
End Sub








' PARTIAL FUNCTIONS

Private Sub Demo_PartialFunctions()

    'Call MessageBoxPilotNames("Bruce Willis", "Nick cage")
    Dim partial As clsPartialFunction
    Set partial = Create_PartialFunction(ThisWorkbook, "MessageBoxPilotNames", "Bruce")
    
    partial.AddArg "Nick"
    
    Call partial.ExecuteCall
End Sub


Private Sub MessageBoxPilotNames(pilot_name As String, copilot_name As String)
    MsgBox "The pilot and copilot are " & pilot_name & " and " & copilot_name & ", respectively."
End Sub








Private Sub Demo_PromptModel()
    Dim DB As clsDb
    Set DB = GetDb_Prompt

    Dim record As clsPrompt
    
    Set record = GetPrompt_ByName(PROMPT_NAME_EXPENSE_CLASSIFY_SELECT)
    
    MsgBox TypeName(record)
    
    Dim DeleteIds As New Collection
    DeleteIds.Add 3
    DeleteIds.Add 4
    DeleteIds.Add 5
    
    Call DB.DeleteRecordRows(DeleteIds)
    
    'MsgBox whiskers.name
    'whiskers.name = "Paws"
    'whiskers.remote_height = 23

End Sub





Private Sub Demo_TestCat_Model()

    Dim DB As clsDb
    Set DB = GetDbOrRaise(ThisWorkbook, "TestCat")
    
    Dim CatCollect As Collection
    Set CatCollect = DB.GetAllRecords
    
    Dim cat As clsTestCat
    
    For Each cat In CatCollect
        Debug.Print cat.name
    Next cat
    Debug.Print String(50, "-")
    
    Dim whiskers As clsTestCat
    Set whiskers = DB.GetRecordById(4)
    
    'MsgBox whiskers.name
    'whiskers.name = "Paws"
    'whiskers.remote_height = 23

End Sub



' UNIT TESTS


'Private Sub Validate_Single_BusinessExpenseRecord()
'    Const OWNER As String = "Validate_Single_BusinessExpenseRecord()"
'    Dim EB As ErrorBank
'    Set EB = CreateErrorBank(OWNER)
'
'    Dim record As clsBusinessExpense
'    Set record = GetDb_BusinessExpense.GetFirstRecordOrNothing
'
'    'Dim TrueBlank As String
'    'TrueBlank = ""
'
'    'MsgBox "A: " & PrettyValueType(record.followup_question)
'    'MsgBox "True Blank: " & PrettyValueType(TrueBlank)
'
'    ' CHECK expenses tagged for review also have a non-blank followup question
'    If record.cost_category = COST_CATEGORY_NEEDS_REVIEW Then
'        msg_fail = "#T1 followup question cannot be blank for expenses classified as: " & COST_CATEGORY_NEEDS_REVIEW
'        'EB.AssertNotEqual record.followup_question, "", msg_fail   ' No, watch out, it will be Empty not ""
'        EB.AssertTruthy record.followup_question, msg_fail
'
'    ' CHECK expenses with a non-blank followup question are also tagged for review
'    ElseIf record.followup_question <> "" Then
'        msg_fail = "#T2 followup question should be blank unless expense is classified as: " & COST_CATEGORY_NEEDS_REVIEW
'        EB.AssertEqual record.cost_category, COST_CATEGORY_NEEDS_REVIEW, msg_fail
'    End If
'
'    Call EB.Render_DynamicResultsSummary(ThisWorkbook, OWNER)
'End Sub


Private Function GetTestErrors_StringExtractNumbers(WB_HOST As Workbook) As ErrorBank
    Const OWNER As String = "GetTestErrors_StringExtractNumbers()"
    Dim EB As ErrorBank
    Set EB = CreateErrorBank(OWNER)
    
    ' CHECK main cases
    EB.AssertType StringExtractNumbers("AAA3"), "String", "#T0 string return to allow leading 0s"
    EB.AssertEqual StringExtractNumbers("AAA3"), "3", "#T1"
    EB.AssertEqual StringExtractNumbers("B2BB"), "2", "#T2"
    EB.AssertEqual StringExtractNumbers(""), "", "#T3"
    EB.AssertEqual StringExtractNumbers("34543"), "34543", "#T4"
    EB.AssertEqual StringExtractNumbers("(*)2&&*^"), "2", "#T5"
    
    ' CHECK decimal points are not counted (decimals are ignored)
    EB.AssertEqual StringExtractNumbers("2.3"), "23", "#T6"
    EB.AssertEqual StringExtractNumbers("fdf2.3f3"), "233", "#T7"
    
    ' CHECK additional edge cases
    EB.AssertEqual StringExtractNumbers("no numbers"), "", "#T8"
    EB.AssertEqual StringExtractNumbers("123abc456"), "123456", "#T9"
    EB.AssertEqual StringExtractNumbers("007James"), "007", "#T10"
    EB.AssertEqual StringExtractNumbers(" 12 "), "12", "#T11"
    
    'EB.AssertTrue
    'EB.assertFalse
    'EB.AssertTruthy
    'EB.AssertNotEqual
    'EB.AssertRangeNameExists
    'EB.AssertRangeContains    ' Blank String for example
    'EB.AssertDbModelEmpty
    'EB.AssertSheetExists
    'EB.AssertStringContains
    'EB.AssertStringNotContains
    'EB.AssertIsNumeric
    'EB.AssertPathExists
    'EB.AssertType                  ' Dictionary vs Collect (Json response) for Example
    'EB.AssertVbaProcedureExists    ' Vba Callback for Example
    
    ' ROLLUP
    Set GetTestErrors_StringExtractNumbers = EB
End Function


Private Sub Demo_ErrorBank()
    Const OWNER As String = "Demo_ErrorBank"
    Dim EB As ErrorBank
    Set EB = CreateErrorBank(OWNER)
    
    ' ARRANGE
    Dim result As Long

    ' CHECK the standard case
    EB.AssertEqual AddTwoNumbers(2, 5), 7, "#T1"
    
    ' CHECK error case (passing a non-number)
    On Error Resume Next
    result = AddTwoNumbers("Cats", 11)
    EB.AssertEqual Err.Number, 13, "#T2 should be a type mismatch error 13"
    On Error GoTo 0
    
    Debug.Assert Err.Number = 0
    Call EB.Render_DynamicResultsSummary(ThisWorkbook, OWNER)
End Sub


Private Function AddTwoNumbers(NumberA As Long, NumberB As Long) As Long
    AddTwoNumbers = NumberA * NumberB
End Function





' NAMED ARRAY
Private Sub Demo_NamedArray()
    Dim corner_cell As Range
    Set corner_cell = GetScratchRangeTableCORNER(ThisWorkbook, 5, 3)
    Call RangeGotoAndSelect(corner_cell)
    
    Dim NR As NamedArray
    Set NR = Create_NamedArrayByCornerRange(corner_cell)
    
    MsgBox NR.TotalDataRows
    
End Sub








' ENUMS





' GETTERS and SETTERS (if we have time)

''Private P_FAVORITE_ANIMAL As String
'
'Public Property Get FavoriteAnimal() As String
'
'
'End Property
'
'
'Public Property Set FavoriteAnimal(new_favorite_animal As String)
'
'
'End Property
'


Attribute VB_Name = "Fixtures"
Option Explicit
Option Private Module
Private msg As String

Public Const FIXTURE_NAME_EXPENSE_RECORDS As String = "TestRecords_BusinessExpense"



Public Function GetDb_BusinessExpense_FIXTURE_LOADED(EB As ErrorBank) As clsDb
    Const OWNER As String = "GetDb_BusinessExpense_FIXTURE_LOADED()"
    
    ' Main job
    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense
    
    ' Bypass / raise smoke if fixture string missing (before deleting existing records)
    If Not RootFixtures.FixtureRowExists(ThisWorkbook, FIXTURE_NAME_EXPENSE_RECORDS) Then
        msg = "Could not find fixture: " & FIXTURE_NAME_EXPENSE_RECORDS & " ... "
        msg = msg & " Raising smoke to stop further unit tests."
        MsgBox msg, vbExclamation, OWNER
        
        Call EB.StopRunningChecksASAP(reason:=msg)
        
        Set GetDb_BusinessExpense_FIXTURE_LOADED = DB
        Exit Function
    End If
    
    ' Clear all existing records
    Call DB.DeleteAllRecordRows
    
    Dim JsonStr As String
    JsonStr = GetFixtureString(ThisWorkbook, FIXTURE_NAME_EXPENSE_RECORDS)
    
    Dim FakeWebRecords As Collection
    Set FakeWebRecords = JsonConverter.ParseJson(JsonStr)
    
    Dim result As ResultIngestRecord
    result = DB.IngestExternalRecords(FakeWebRecords)
    
    ' EXAMPLE - suppressing messagebox during tests
    '    MsgBoxExceptTesting "DKLJFDLKFJDF"
    '
    '    If Not WhoaUnitTestsAreRunning Then
    '        MsgBox "DKLJFDLKFJDF"
    '    End If
    '
    '    If Not WhoaUnitTestsAreRunning Then
    '        Call ResultIngestRecord_DisplayInfo(result)
    '    End If

    Set GetDb_BusinessExpense_FIXTURE_LOADED = DB
End Function



Private Sub RegenerateDbFixtures_BusinessExpense()
    Const OWNER As String = "RegenerateDbFixtures_BusinessExpense()"
    Const NUMBER_FAKE_RECORDS As Long = 5
    
    Dim model_name As String
    model_name = MODEL_NAME_BUSINESS_EXPENSE
    
    Dim DB As clsDb
    Set DB = GetDbOrRaise(ThisWorkbook, model_name, OWNER)
    
    ' BYPASS if no records are imported yet
    If DB.RecordCount = 0 Then
        msg = "No records are imported (cannot generate a TRUTHY fixture string)"
        MsgBox msg, vbInformation, OWNER
        Exit Sub
    End If

    ' BYPASS / Warn user if this operation would overwrite an Existing Fixture
    If FixtureRowExists(ThisWorkbook, FIXTURE_NAME_EXPENSE_RECORDS) Then
    
        msg = "Model " & model_name & " has an existing fixture string... press YES to overwrite it"
        
        If MsgBox(msg, vbYesNo + vbInformation, OWNER) <> vbYes Then
            Exit Sub
        End If
    End If

    ' Create JSON string from the first N records
    Dim JsonStr As String
    JsonStr = Generate_FixtureString_CurrentRecords(MODEL_NAME_BUSINESS_EXPENSE, NUMBER_FAKE_RECORDS)

    Call RootFixtures.CreateFixtureIfMissing( _
        ThisWorkbook, _
        FIXTURE_NAME_EXPENSE_RECORDS, _
        default_value:=JsonStr, _
        fixture_type:="DbRecords" _
    )
    
End Sub





Private Sub SCRATCH_TEST_Generate_FixtureString_CurrentRecords()
    Debug.Print Generate_FixtureString_CurrentRecords(MODEL_NAME_BUSINESS_EXPENSE, 5)
End Sub


Private Function Generate_FixtureString_CurrentRecords(model_name As String, number_records As Long) As String
    Const OWNER As String = "Generate_FixtureString_CurrentRecords()"

    Dim DB As clsDb
    Set DB = GetDbOrRaise(ThisWorkbook, model_name, OWNER)
    
    Dim AllRecords As Collection
    Set AllRecords = GetDbRecordsByModel(ThisWorkbook, model_name)
    
    Dim SubsetRecords As Collection
    Set SubsetRecords = SliceCollection(AllRecords, 1, number_records + 1)
    
    Dim FakeWebRecords As New Collection    ' ok to use one liner
    Dim record As clsBusinessExpense

    For Each record In SubsetRecords
        FakeWebRecords.Add record.ToDict
    Next record
    
    Dim JsonStr As String
    JsonStr = JsonConverter.ConvertToJson(FakeWebRecords, 1)
    
    Generate_FixtureString_CurrentRecords = JsonStr
End Function


Attribute VB_Name = "Imports"
' Imports.bas (Standard Module inside an XLSM Host Workbook, name it something descriptive like "Imports")
Option Explicit

' Always Private module unless you have a reason not to (like you want to see a functions in UDF / cell formulas)
Option Private Module

' Shared variable msg can be used in any function for short text
' Saves many of "Dim msg as String" lines
' This would definitely be Private because every module has its own "msg"
Private msg As String


' Below is the procedure we call to import data
' You can call this sub from a sheet-level command button, keyboard shortcut... wherever you want!
' You can also run this sub by putting your cursor anywhere inside of it, then pressing F5
Public Sub Import_BusinessExpense_Records()
    Const OWNER As String = "Import_BusinessExpense_Records()"
    
    ' Prompt the user to select a file path
    ' - if your raw data is not a .csv, there are several other helper functions are available
    '   eg PromptForFileOpenWorkbook() will cover .xlsx
    Dim FilePath As String
    FilePath = PromptForFileOpenCSV
    
    ' BYPASS if user cancels the file dialog
    If FilePath = "" Then
        Exit Sub ' MANHOLE return
    End If

    Dim WB_HOST As Workbook
    Set WB_HOST = ThisWorkbook
    
    ' Enable performance mode to both:
    ' 1. speed up the import
    ' 2. prevent screen updating (so the data csv never appears)
    Call WhoaPerformanceSettings_ON(WB_HOST, OWNER)
    
    ' Open the CSV containing the data as a workbook
    Dim WB_DATA As Workbook
    Set WB_DATA = WorkbookOpenSafe(FilePath)
    
    Dim data_sheet As Worksheet
    Set data_sheet = WB_DATA.Worksheets(1)
    
    Dim NR As NamedArray
    Set NR = Create_NamedArrayByCornerRange(data_sheet.Cells(1, 1))
    
    Dim WebRecords As New Collection
    Dim record_dict As Dictionary
    Dim row_number As Long
    
    ' Loop each row in the data file, creating a dictionary per row
    For row_number = 1 To NR.TotalDataRows
    
        ' create a new dictionary for each row
        Set record_dict = New Dictionary
        
        ' load the dictionary
        ' - left side of "=" operator is internal DB field names, eg "vendor_name"
        ' - right side  "=" operator is external csv header titles, eg "vendor"
        
        record_dict.Add "id", NR.GetValueByRowNumber(row_number, "id")
        record_dict("transact_date") = NR.GetValueByRowNumber(row_number, "date")
        record_dict("vendor_name") = NR.GetValueByRowNumber(row_number, "vendor")
        record_dict("description") = NR.GetValueByRowNumber(row_number, "description")
        record_dict("amount") = NR.GetValueByRowNumber(row_number, "amount")
        
        ' load each dictionary into the WebRecords collection
        WebRecords.Add record_dict
    Next row_number
    
    ' Acquire the DB you want to import the data into
    Dim DB As clsDb
    Set DB = GetDb_BusinessExpense
    
    ' Pass a single collection into the DB to ingest the data
    Dim result As ResultIngestRecord
    result = DB.IngestExternalRecords(WebRecords)
    
    WB_DATA.Close SaveChanges:=False
    
    ' Turn off performance at the end
    ' - This restores whatever XL Calculation mode (Automatic or Manual) in effect at the beginning
    Call WhoaPerformanceSettings_OFF(WB_HOST, OWNER)
    
    ' Show the user counts of "created records" vs "modified records"
    Call ResultIngestRecord_DisplayInfo(result)
End Sub


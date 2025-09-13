Attribute VB_Name = "JsonExample"
Option Explicit
Option Private Module
Private msg As String


Private Sub ForexJsonExample()

    Const URL As String = "https://www.forexfactory.com/calendar/graph/142633?limit=5&site_id=1"
    
    Dim request As WinHttpRequest
    Set request = New WinHttpRequest
    
    request.Open "GET", URL, False
    request.Send
    
    Dim responseText As String
    responseText = request.responseText
    
    Dim responseDict As Dictionary
    Set responseDict = JsonConverter.ParseJson(responseText)
    
    ' Temporary pretty string for demonstration purposes
    Dim prettyResponseText As String
    prettyResponseText = JsonConverter.ConvertToJson(responseDict, 2)
    Debug.Print prettyResponseText
    
    Dim EventsCollection As Collection
    Set EventsCollection = responseDict("data")("events")
    
    Dim D As Dictionary
    For Each D In EventsCollection
        Debug.Print "EVENT: " & D("id"); " - " & D("date")
    Next D
    
End Sub

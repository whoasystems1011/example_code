Attribute VB_Name = "AI_Core"
Option Explicit
Option Private Module
Private msg As String



Public Function ExecutePrompt_WithCallback_VBA( _
    final_prompt As String, _
    partial_vba_callback As clsPartialFunction _
) As ChatCompletionUDT

    Dim chatComp As ChatCompletionUDT
    chatComp = AI_MessageChatbot_OpenAI(ThisWorkbook, final_prompt)
    Call ChatCompletionUDT_DisplayInfo(chatComp)

    Dim responseText As String
    responseText = chatComp.latestMessage

    ' BYPASS if the chatbot response is invalid (not a json)
    If ReturnTrueAndWarnUser_IfChatbotResponseIsInvalid(responseText) Then
        Exit Function
    End If

    Dim responseObject As Object      ' NOTE its a Collection or a Dictionary
    Set responseObject = JsonConverter.ParseJson(responseText)
    
    ' Attach the chatbot response, parsed into an object
    partial_vba_callback.AddArg responseObject
    
    ' Execute the callback
    Call partial_vba_callback.ExecuteCall
    

    ' Return the chatbots raw response text (in case the caller needs it)
    ExecutePrompt_WithCallback_VBA = chatComp
End Function



Private Function ReturnTrueAndWarnUser_IfChatbotResponseIsInvalid(response_text As String) As Boolean
    Const OWNER As String = "ReturnTrueAndWarnUser_IfChatbotResponseIsInvalid()"

    Dim reasons As New Collection
    Debug.Assert reasons.Count = 0
    
    ' CHECK chatbot provided a valid json (vs "Hi! How can i assist you today")
    If Not IsJson(response_text) Then
        reasons.Add "The chatbot respponse must be a valid JSON"
    End If
    
    ' RECONCILE
    Dim ProblemsExist As Boolean
    
    If reasons.Count = 0 Then
        ProblemsExist = False
    Else
        ProblemsExist = True
        Dim failure_reasons As String
        failure_reasons = CollectionToNumberedList(reasons)
        MsgBox failure_reasons, vbExclamation, OWNER
    End If

    ReturnTrueAndWarnUser_IfChatbotResponseIsInvalid = ProblemsExist
End Function




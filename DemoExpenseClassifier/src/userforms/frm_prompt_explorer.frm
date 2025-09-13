VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_prompt_explorer 
   Caption         =   "Prompt Explorer"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20190
   OleObjectBlob   =   "frm_prompt_explorer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_prompt_explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msg As String

Private GV_ACTIVE_PROMPT As clsPrompt
Private GV_UNLOAD_CALLBACK As clsPartialFunction


'--------------|
'--Initialize--|
'--------------|
Public Sub Real_Initialize(prompt_name As String, Optional unload_callback As clsPartialFunction)
    Set GV_ACTIVE_PROMPT = GetPrompt_ByName(prompt_name)
    Set GV_UNLOAD_CALLBACK = unload_callback
    Call RefreshDisplayElements
End Sub



Private Sub btn_validate_prompt_Click()
    Call ValidateActivePrompt
End Sub


Private Sub UserForm_Initialize()
    Call CenterUserFormOnScreen(Me)
End Sub


Private Sub UserForm_Activate()
    ' BYPASS if a caller somewhere called frm_prompt_explorer.show directly,
    ' instead of calling OpenForm_PromptExplorer()
    If GV_ACTIVE_PROMPT Is Nothing Then
        msg = "Error - please use the AppForms function to open PromptExplorer. Unloading form."
        MsgBox msg, vbExclamation, Me.Caption
        Unload Me
    End If
End Sub




'----------------|
'--Main Scripts--|
'----------------|
Private Sub RefreshDisplayElements()
    Me.txt_base_text.Value = GV_ACTIVE_PROMPT.base_text
    Me.txt_name.Value = GV_ACTIVE_PROMPT.name
    Me.txt_vba_callback_name.Value = GV_ACTIVE_PROMPT.vba_callback_name
    Me.txt_notes.Value = GV_ACTIVE_PROMPT.notes
    
    ' Initialize all controls so they have White background
    ' This is because they momentarily change to yellow during
    ' initialization (this makes sure they end up white)
    Dim ctrl As Object
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.BackColor = vbWhite
        End If
    Next ctrl
End Sub


Private Sub SaveInputValuesToDb()
    GV_ACTIVE_PROMPT.base_text = Me.txt_base_text.Value
    GV_ACTIVE_PROMPT.name = Me.txt_name.Value
    
    GV_ACTIVE_PROMPT.vba_callback_name = Me.txt_vba_callback_name.Value
    GV_ACTIVE_PROMPT.notes = Me.txt_notes.Value
    
    MsgBox "Record values updated successfully", vbInformation, Me.Caption
    
    Call RefreshDisplayElements
End Sub


Private Sub GotoNextPrompt()
    Dim DB As clsDb
    Set DB = GetDb_Prompt
    
    Dim next_record_id As String
    next_record_id = DB.GetNextId_OrBlank(GV_ACTIVE_PROMPT.id)
    
    If next_record_id = "" Then
        Set GV_ACTIVE_PROMPT = DB.GetFirstRecordOrNothing
        Debug.Assert TypeName(GV_ACTIVE_PROMPT) <> "Nothing"
    Else
        Set GV_ACTIVE_PROMPT = DB.GetRecordById(next_record_id)
    End If
    
    Call RefreshDisplayElements
End Sub



'-----------------|
'--Button Events--|
'-----------------|
Private Sub btn_save_Click()
    Call SaveInputValuesToDb
End Sub


Private Sub btn_open_notepad_Click()
    Call ExportTextFile(Me.txt_base_text, GV_ACTIVE_PROMPT.name, True)
End Sub


Private Sub ValidateActivePrompt()
    Dim EB As ErrorBank
    Set EB = GV_ACTIVE_PROMPT.GetValidateErrors
    Call EB.Render_MessageBoxErrorList(RenderEvenIfNoErrors:=True)
End Sub



'------------------|
'--Special Events--|
'------------------|
Private Sub txt_name_Change()
    txt_name.BackColor = vbYellow
End Sub


Private Sub txt_base_text_Change()
    txt_base_text.BackColor = vbYellow
End Sub


Private Sub txt_notes_Change()
    txt_notes.BackColor = vbYellow
End Sub


Private Sub txt_vba_callback_name_Change()
    txt_vba_callback_name.BackColor = vbYellow
End Sub


Private Sub label_goto_next_prompt_Click()
    Call GotoNextPrompt
End Sub



'--------|
'--Exit--|
'--------|
Private Sub btn_exit_Click()
    Call FormUnloadWithDynamicRedirect(Me, GV_UNLOAD_CALLBACK)
End Sub


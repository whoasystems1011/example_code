Attribute VB_Name = "AppForms"
Option Explicit



Public Sub OpenForm_PromptExplorer(Optional prompt_name As String = "", Optional unload_callback As clsPartialFunction)
    
    If prompt_name = "" Then
        prompt_name = PROMPT_NAME_EXPENSE_CLASSIFY_SELECT
    End If

    Call frm_prompt_explorer.Real_Initialize(prompt_name, unload_callback)
    frm_prompt_explorer.Show vbModeless
End Sub

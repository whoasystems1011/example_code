Attribute VB_Name = "Demo3_ProgressBar"
Option Explicit
Option Private Module
Private msg As String



Private Sub DEMO_PROGRESS_BAR_SIMPLE()
    
    ' Get the progress bar instance
    Dim pbar As clsProgressBar
    Set pbar = GetWhoaProgressBar_AndDisplay("Whoa Progress Bar - Test Drive")
    
    ' Provide the total number of steps to the progress bar
    ' This is needed for the bar to calculate "how much to advance" each time
    ' pBar.IncrementStep() is called
    
    Const TOTAL_STEPS As Long = 5
    pbar.StartActivity TOTAL_STEPS, "Starting up something"
    
    Dim count As Long
    For count = 1 To TOTAL_STEPS
        
        ' calling pbar.IncrementStep() advances the bar forward
        ' You can also update the text (that appears under the bar) as a way
        ' to inform user where they are in the overall process
        pbar.IncrementStep "Running step: " & count
        Call WhoaSleepHard(NumberSeconds:=1.05)
        
        ' On the second iteration, change the bar color to red
        If count = 2 Then
            
            ' Change the bar COLOR
            pbar.ChangeBarColor vbRed
            
            ' Change the bar TEXT
            ' Sometimes this is useful when you specifically DON'T want to
            ' call pbar.IncrementStep (which would advance the bar), but still want
            ' to emit some info to the user.
            
            pbar.ChangeActivityText "Changing bar color to RED!"
            
            ' Pause a moment
            ' Without a pause, the "changed bar" text would be overwritten instantly.
            Call WhoaSleepHard(NumberSeconds:=0.55)
        End If
        
    Next count
    
    ' Shutdown the bar
    ' This may not always be necessary, but it's never harmful
    Call pbar.Shutdown
End Sub





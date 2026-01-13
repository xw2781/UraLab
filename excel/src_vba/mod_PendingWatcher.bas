Attribute VB_Name = "mod_PendingWatcher"
Public NextCheckTime As Date

Private Sub ScheduleNextCheck()
    NextCheckTime = Now + TimeSerial(0, 0, 4)
    Application.OnTime NextCheckTime, "PendingWatcher_Tick"
End Sub

Public Sub PendingWatcher_Tick()
    If disableWatcher Then Exit Sub

    ' Only act when Excel is idle
    If pendingUpdate _
       And Application.CalculationState = xlDone _
       And Application.Interactive Then

        pendingUpdate = False
        Call CalculateWorkbook
    End If

    ScheduleNextCheck
End Sub

Public Sub StartPendingWatcher()
    disableWatcher = False
    ScheduleNextCheck
End Sub

Public Sub StopPendingWatcher()
    On Error Resume Next
    disableWatcher = True
    Application.OnTime NextCheckTime, "PendingWatcher_Tick", , False
End Sub

Sub testVar()
    pendingUpdate = 0
End Sub

Sub check_status()
    qqq "---------------"
    If disableWatcher Then
        qqq "Watcher Disabled"
    Else
        qqq "Watcher Enabled"
    End If
    If pendingUpdate Then qqq "!pendingUpdate"
End Sub


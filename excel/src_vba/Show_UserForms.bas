Attribute VB_Name = "Show_UserForms"
' ufLoading

Sub Show_ufLoading()
    If disable_ufLoading Then
        Exit Sub
    End If

    If RateLimited("Show_ufLoading") Then
        pendingUpdate = True
        skipDataProcess = True
        ' qqq "Rate Limited1"
        ' Application.OnTime EarliestTime:=Now + TimeSerial(0, 0, 3), Procedure:="CalculateWorkbook", Schedule:=True
        Exit Sub
    End If

    ufLoading.Show vbModeless
End Sub

' ufProgressBar

Sub Show_ufProgressBar()
    ufProgressBar.Show vbModeless
    DoEvents
End Sub



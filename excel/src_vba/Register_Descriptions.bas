Attribute VB_Name = "Register_Descriptions"
Sub Register_ADASTri_Help_Safe()
    Dim wb As Workbook
    Dim wasAddin As Boolean

    Set wb = ThisWorkbook                 ' the XLAM
    wasAddin = wb.IsAddin

    ' Temporarily show the add-in so MacroOptions can edit its metadata
    If wasAddin Then
        wb.IsAddin = False
        On Error Resume Next
        Windows(wb.Name).Visible = True
        On Error GoTo 0
    End If

    ' ---- register descriptions ----
    Application.MacroOptions _
        Macro:=wb.Name & "!ADASTri", _
        Description:="Returns a triangle dataset from the ADAS data system with optional cumulative, transpose, calendar, and dimension controls.", _
        Category:="ADAS Tools", _
        ArgumentDescriptions:=Array( _
            "Full path key (string). Example: ""PIC2\PA\NJ\Core Direct\PD"".", _
            "Dataset/Triangle Name (string). Example: ""Net Loss--Paid"", ""Claim Counts--Reported"", etc.", _
            "TRUE = Return cumulative triangle. FALSE = Return incremental triangle. (Default = TRUE)", _
            "TRUE = Return triangle transposed (Development x Origin). FALSE = Standard orientation. (Default = FALSE)", _
            "TRUE = Convert to calendar view; FALSE = Standard accident/origin view. (Default = FALSE)", _
            "Virtual Project Name. ""Default"" uses currently active project in environment.", _
            "Number of Origin periods to return (Default = 12).", _
            "Number of Development periods to return (Default = 12).", _
            "Type selector or metric variant (optional). If omitted, uses dataset’s default type.", _
            "TRUE = Suppress warning messages. FALSE = Show warnings. (Optional)" _
        )

    ' Re-hide the add-in
    If wasAddin Then
        wb.IsAddin = True
    End If

    ' MsgBox "ADASTri help registered.", vbInformation
End Sub

Sub SetAddinDescription()
    With ThisWorkbook.BuiltinDocumentProperties
        .item("Title").Value = "ADAS"
        .item("Comments").Value = "Adaptive Database and Analytics System"
        .item("Subject").Value = "Actuarial Utilities"
    End With
    ThisWorkbook.Save
End Sub




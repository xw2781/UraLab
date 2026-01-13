Attribute VB_Name = "Ribbon_Control"

'+----------+
'|  Group 1 |
'+----------+

' Connection and Login
Sub uiSetupConnection2(control As IRibbonControl)
    SetupConnection2
End Sub

' Calculate Sheet
Sub uiRefreshSheet(control As IRibbonControl)
    CalculateSheet
End Sub

' Calculate Workbook
Sub uiRefreshWorkbook(control As IRibbonControl)
    CalculateWorkbook
End Sub

' Refresh Database
Sub uiRefreshDatabase(control As IRibbonControl)
    RefreshDatabase
End Sub

'+----------+
'|  Group 2 |
'+----------+

' Insert Function
Sub uiInsertFunction(control As IRibbonControl)
    Application.Dialogs(xlDialogFunctionWizard).Show
End Sub

' Clear Formulas
Sub uiClearResQFormulae2(control As IRibbonControl)
    MsgBox "NaN"
End Sub

' Load Reserving Classes
Sub uiLoadReservingClasses2(control As IRibbonControl)
    ufLoadReservingClasses.Show vbModeless
End Sub

' Select Dataset
Sub uiSelectDatasets(control As IRibbonControl)
    ufSelectDataset.Show vbModeless
End Sub

'+----------+
'|  Group 3 |
'+----------+

' Reset References
Sub uiResetAddinReferences(control As IRibbonControl)
    ResetAddinReferences
End Sub

' Load Add-in
Sub uiLoadAddIn(control As IRibbonControl)
    LoadAddIn
End Sub

' Unload Add-in
Sub uiUnloadAddIn(control As IRibbonControl)
    UnloadAddIn
End Sub

'+----------+
'|  Group 4 |
'+----------+

' Check Updates
Sub uiCheckUpdates(control As IRibbonControl)
    ' CheckUpdates
    ufBuildTri.Show
End Sub

' User Settings
Sub uiSettings(control As IRibbonControl)
    ufSettings.Show vbModeless
End Sub

' About
Sub uiAbout(control As IRibbonControl)
    ufAbout.Show vbModeless
End Sub


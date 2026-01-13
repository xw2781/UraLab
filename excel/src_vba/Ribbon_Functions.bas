Attribute VB_Name = "Ribbon_Functions"
Sub CalculateWorkbook()
    If disableProgressBar Then
        CalculateWorkbookNoUI
    Else
        CalculateWorkbookWithUI
    End If
End Sub

Sub CalculateWorkbookWithUI()
    Dim item As Variant
    Dim totalDatasets As Long
    Dim countDataset As Long
    Dim removeDataStat As Long
    Dim currentSheet As Worksheet
    Dim oldCalcMode As XlCalculation
    
    On Error GoTo ErrorHandler
    
    Set currentSheet = ActiveWorkbook.ActiveSheet
    
    oldCalcMode = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    'Application.Calculation = xlCalculationManual
    
    errCount = 0
    removeDataStat = removeData
    disable_ufLoading = True
    skipDataProcess = False
    doubleRefresh = Fasle
    
    Show_ufProgressBar
    ufProgressBar.LabelTitle.Caption = "Searching for ADAS formulas ..."
    DoEvents
   
    ' Step (1) Search & Send Requests
    Call SearchADASFormulas
    
    If processedArrays Is Nothing Then Exit Sub
    If processedArrays.Count = 0 Then Exit Sub
    
    totalDatasets = processedArrays.Count
    countDataset = 0
    
    If doubleRefresh Then
        ' Step (2) Pull Datasets
        ufProgressBar.LabelTitle.Caption = "Refreshing datasets ..."
        ufProgressBar.LabelDetails.Caption = totalDatasets & " dataset(s) need to be refreshed"
        removeData = False
        disableRequest = True
        For Each item In processedArrays
            If cancelUpdate Then GoTo CleanExit
            Call RefreshADASBlock(CStr(item))
            
            countDataset = countDataset + 1
            ' Refresh UI
            ' ufProgressBar.LabelBody.Caption = "<" & Replace(item, "!", ">! ")
            ufProgressBar.LabelDetails.Caption = countDataset & "/" & totalDatasets & " dataset(s) updated"
            ufProgressBar.UpdateProgressBar countDataset / totalDatasets * 100
            If countDataset Mod 20 = 0 Then DoEvents
        Next item
        
        ufProgressBar.LabelTitle.Caption = "Calculating all formulas in this workbook, please wait ..."
        Application.Wait Now + TimeValue("0:00:01")
        
    End If
    
    Application.StatusBar = "[" & ActiveWorkbook.Name & "] - Refreshed at " & Format(Now, "hh:mm:ss")
    
CleanExit:
    removeData = removeDataStat
    Application.Calculation = oldCalcMode
    Unload ufProgressBar
    ufProgressBar.ClearText
    disable_ufLoading = False
    cancelUpdate = False
    disableRequest = False
    currentSheet.Activate
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "[" & ActiveWorkbook.Name & "] - Workbook not refreshed! Updated @ " & Format(Now, "hh:mm:ss")
    Resume CleanExit
    
End Sub

Sub CalculateSheet()
    Dim item As Variant
    Dim totalDatasets As Long
    Dim countDataset As Long
    Dim removeDataStat As Long
    Dim currentSheet As Worksheet
    Dim oldCalcMode As XlCalculation
    
    On Error GoTo ErrorHandler
    
    Set currentSheet = ActiveWorkbook.ActiveSheet
    
    oldCalcMode = Application.Calculation
    'Application.Calculation = xlCalculationManual
    
    errCount = 0
    removeDataStat = removeData
    disable_ufLoading = True
    skipDataProcess = False
    doubleRefresh = Fasle
    
    Show_ufProgressBar
    ufProgressBar.LabelTitle.Caption = "Searching for ADAS formulas ..."
    DoEvents
   
    ' Step (1) Search & Send Requests
    SearchADASFormulas True
    
    If processedArrays Is Nothing Then Exit Sub
    If processedArrays.Count = 0 Then Exit Sub
    
    totalDatasets = processedArrays.Count
    countDataset = 0
    
    If doubleRefresh Then
        ' Step (2) Pull Datasets
        ufProgressBar.LabelTitle.Caption = "Refreshing datasets ..."
        ufProgressBar.LabelDetails.Caption = totalDatasets & " dataset(s) need to be refreshed"
        removeData = False
        disableRequest = True
        For Each item In processedArrays
            If cancelUpdate Then GoTo CleanExit
            Call RefreshADASBlock(CStr(item))
            
            countDataset = countDataset + 1
            ' Refresh UI
            ufProgressBar.LabelBody.Caption = "<" & Replace(item, "!", ">! ")
            ufProgressBar.LabelDetails.Caption = countDataset & "/" & totalDatasets & " dataset(s) updated"
            ufProgressBar.UpdateProgressBar countDataset / totalDatasets * 100
            If countDataset Mod 20 = 0 Then DoEvents
        Next item
        
        ufProgressBar.LabelTitle.Caption = "Calculating all formulas in this workbook, please wait ..."
        Application.Wait Now + TimeValue("0:00:01")
    End If
    
    Application.StatusBar = "[" & currentSheet.Name & "] - Refreshed at " & Format(Now, "hh:mm:ss")
    Application.Wait Now + TimeValue("0:00:01")
    
CleanExit:
    removeData = removeDataStat
    Application.Calculation = oldCalcMode
    Unload ufProgressBar
    ufProgressBar.ClearText
    disable_ufLoading = False
    cancelUpdate = False
    disableRequest = False
    Exit Sub
    
ErrorHandler:
    Resume CleanExit
    
End Sub

Sub CalculateWorkbookNoUI()
    On Error GoTo ErrorHandler
    skipDataProcess = False
    disable_ufLoading = True

    ' (1) Send Request Only, No Wait
    disableRequest = False
    disableWaitTime = True
    Application.Calculate
    
    ' (2) Pull Cached Datasets
    disableWaitTime = False
    Application.CalculateFull

CleanExit:
    disableWaitTime = False
    disable_ufLoading = False
    Exit Sub
    
ErrorHandler:
    Resume CleanExit
    
End Sub

Public Sub SearchADASFormulas(Optional ByVal ActiveSheetOnly As Boolean = False)

    Dim ws As Worksheet
    Dim cell As Range
    Dim arrCell As Range
    Dim cellKey As String
    Dim arrKey As String
    Dim totalSheets As Long
    Dim countSheets As Long
    Dim countCells As Long
    
    On Error GoTo ErrorHandler
    disableWaitTime = True
    
    Set processedCells = New Collection
    Set processedArrays = New Collection

    ' Decide scope
    If ActiveSheetOnly Then
        totalSheets = 1
    Else
        totalSheets = ActiveWorkbook.Worksheets.Count
    End If

    countSheets = 1
    countCells = 1
    
    ' Loop sheets
    For Each ws In ActiveWorkbook.Worksheets
        
        ' Skip non-active sheets if needed
        If ActiveSheetOnly Then
            If ws.Name <> ActiveSheet.Name Then GoTo ContinueLoop
        End If
        
        If cancelUpdate Then GoTo CleanExit
        If ws.Name = "ResQ Settings" Then GoTo ContinueLoop
        
        ufProgressBar.LabelBody.Caption = "Reading worksheet <" & ws.Name & ">"
        DoEvents
        
        On Error Resume Next
        For Each cell In ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo ErrorHandler
            If ActiveSheetOnly Then
                ufProgressBar.LabelDetails.Caption = "Looking at cell " & cell.Address(0, 0)
                If countCells Mod 20 = 0 Then DoEvents
            End If
            cellKey = ws.Name & "!" & cell.Address
            If Not KeyExists(processedCells, cellKey) Then
                If InStr(cell.formula, "ADAS") > 0 Then
                    If cell.HasArray Then
                        cell.CurrentArray.FormulaArray = cell.CurrentArray.FormulaArray
                        arrKey = ws.Name & "!" & cell.CurrentArray.Address
                        processedArrays.Add arrKey, arrKey
                        
                        For Each arrCell In cell.CurrentArray
                            cellKey = ws.Name & "!" & arrCell.Address
                            processedCells.Add cellKey, cellKey
                        Next arrCell
                    Else
                        cell.Formula2 = cell.Formula2
                        processedCells.Add cellKey, cellKey
                        processedArrays.Add cellKey, cellKey
                    End If
                End If
            End If
            countCells = countCells + 1
        Next cell
        
        ufProgressBar.LabelDetails.Caption = countSheets & "/" & totalSheets & " sheet(s) reviewed"
        ufProgressBar.UpdateProgressBar countSheets / totalSheets * 100
        countSheets = countSheets + 1
        DoEvents
        
ContinueLoop:
    Next ws
    
CleanExit:
    ufProgressBar.ClearText
    cancelUpdate = False
    disableWaitTime = False
    Exit Sub
    
ErrorHandler:
    Resume CleanExit

End Sub

Function KeyExists(coll As Collection, key As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    item = coll(key)
    KeyExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        SheetExists = False
    Else
        SheetExists = True
    End If
End Function

Public Sub SetupConnection2()
    Dim sheet1 As Worksheet
    Dim Sheet2 As Worksheet
        
    If Not SheetExists("ResQ Settings") Then
        Set sheet1 = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Sheets(1))
        sheet1.Name = "ResQ Settings"
    End If
    
    If Not SheetExists("Project Details") Then
        Set Sheet2 = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Sheets(2))
        Sheet2.Name = "Project Details"
    End If
    
    Set sheet1 = ActiveWorkbook.Sheets("ResQ Settings")
        sheet1.Columns("A").ColumnWidth = 72.71
        sheet1.Columns("B").ColumnWidth = 44.71
        
        sheet1.Range("A1").Value = "Connection Name"
        sheet1.Range("A2").Value = "Windows Authentication"
        sheet1.Range("A3").Value = "User Name"
        sheet1.Range("A7").Value = "Default Project Name"
        sheet1.Range("A9").Value = "Project Names"
        If sheet1.Range("B7").Value = "" Then sheet1.Range("B7").Value = "NJ_Annual_Prod_2025 Q4-Nov"
    
    Set Sheet2 = ActiveWorkbook.Sheets("Project Details")
        Sheet2.Columns("B").ColumnWidth = 22.14
        Sheet2.Columns("C").ColumnWidth = 39.43
        
        Sheet2.Range("B4:C11").FormulaArray = "=ADASProjectSettings()"
        Sheet2.Range("C4:C11").Interior.Color = 10092543
        Sheet2.Range("C4:C11").HorizontalAlignment = xlCenter
        Sheet2.Range("C4:C11").Font.Color = 255 ' Red
        Sheet2.Range("C4:C11").Font.Bold = True
        
        Dim borders() As Variant
        borders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical)
        For i = LBound(borders) To UBound(borders)
            With Sheet2.Range("B4:C11").borders(borders(i))
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        Next i
        
        Sheet2.Range("C6:C8").NumberFormat = "m/d/yy"
        
    On Error Resume Next
    
    On Error GoTo 0
End Sub

Sub RefreshDatabase()

    Dim filePath As String
    Dim fileNumber As Integer
    Dim maxWaitTime As Long
    Dim startTime As Long
    Dim currentTime As String
    Dim resultPath As String

    currentTime = Format(Now, "yyyy-mm-dd_hh-mm-ss") & Format(Timer - Int(Timer), ".000")
    filePath = "E:\ResQ\Excel Add-ins\requests\" & "request-" & currentTime & ".txt"
    resultPath = "E:\ResQ\Excel Add-ins\data\" & "data-" & currentTime & ".csv"
    
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
        Print #fileNumber, "Function = RefreshDatabase"
        Print #fileNumber, "ResultPath = " & resultPath
        Print #fileNumber, "UserName = " & Environ("USERNAME")
    Close #fileNumber
    
    maxWaitTime = 3 ' Maximum wait time in seconds
    startTime = Timer
    Do While Timer < startTime + maxWaitTime
        If Dir(resultPath) <> "" Then Exit Do
        DoEvents
    Loop
    
    ' Check if the output file exists
    If Dir(resultPath) = "" Then
        MsgBox "Unable to connect, please try again later."
    Else
        MsgBox "Connection Updated!"
        ' Kill resultPath
    End If
    
End Sub

Sub LoadAddIn()
    errCount = 0
    skipDataProcess = False
    pendingUpdate = False
    Application.StatusBar = "Calculation Resumed"
End Sub

Sub UnloadAddIn()
    On Error GoTo ErrorHandler
    Dim addIn As addIn
    
    For Each addIn In AddIns
        'If InStr(addIn.name, "ADAS") > 0 Then
        If addIn.Name = "ADAS.xlam" Or addIn.Name = "ADAS_BETA.xlam" Then
            addIn.Installed = False
            Exit For
        End If
    Next addIn
    
ErrorHandler:
    ' MsgBox "An error occurred: " & Err.Description
    ' MsgBox "An error occurred when unloading the add-in"
    Err.Clear
    On Error GoTo 0
End Sub

Sub CheckUpdates()
    Dim exePath As String
    Dim retVal As Long
    
    exePath = "\\Ne7saswpn02\e\ResQ\Excel Add-ins\Update\dist\AutoUpdate.exe"
    
    ' Run the executable
    retVal = Shell(exePath, vbNormalFocus)
    
    ' Check the return Value
End Sub

Sub ReplaceInWorkbook(findText As String, replaceText As String)
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Replace What:=findText, Replacement:=replaceText, _
            LookAt:=xlPart, SearchOrder:=xlByRows, _
            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Next ws
    
End Sub

Sub ResetAddinReferences()

    Dim link As Variant
    Dim book As Workbook
    Dim oldCalcMode As XlCalculation
    Dim hasOldLink As Boolean
    Dim hasNewLink As Boolean
    Dim hasBetaLink As Boolean
    Dim TextADAS As String
    Dim TextADAS_BETA As String
    Dim TextResQ As String
    
    On Error GoTo CleanExit
    
    oldCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Set book = ActiveWorkbook
    
    TextADAS = "='E:\ADAS\Excel Add-ins\ADAS.xlam'!ADAS"
    TextADAS_BETA = "='E:\ADAS\Excel Add-ins\beta\ADAS_BETA.xlam'!ADAS"
    TextResQ = "='C:\Program Files\Willis Towers Watson\ResQ\Addins\ResQ.xlam'!ResQ"
    
    ' Check all links in the workbook
    If Not IsEmpty(book.linkSources()) Then
        For Each link In book.linkSources()
            If InStr(link, "ResQ.xlam") > 0 Then hasOldLink = True
            If InStr(link, "ADAS") > 0 Then hasNewLink = True
            If InStr(link, "ADAS_BETA.xlam") > 0 Then hasBetaLink = True
        Next link
    End If
    
    If hasOldLink Then ' Change to ADAS
        skipDataProcess = True
        ReplaceInWorkbook "='C:\Program Files (x86)\Willis Towers Watson\ResQ\Addins\ResQ.xlam'!ResQ", "=ADAS"
        ReplaceInWorkbook TextResQ, "=ADAS"
        ReplaceInWorkbook "=ResQ", "=ADAS"
        Application.StatusBar = "ADAS Excel Add-in activated."
        
    ElseIf Not hasOldLink And hasNewLink Then ' Change to ResQ
        If Dir("C:\Program Files\Willis Towers Watson\ResQ\Addins\ResQ.xlam") <> "" Then
            If hasBetaLink Then
                ReplaceInWorkbook TextADAS_BETA, TextResQ
            Else
                ReplaceInWorkbook TextADAS, TextResQ
            End If
            ReplaceInWorkbook "=ADAS", TextResQ
            Application.StatusBar = "ResQ Excel Add-in activated."
        Else
            Application.StatusBar = "Error: ResQ Excel Add-in can only be activated on Remote Desktop!"
        End If
    End If
    
    If ThisWorkbook.Name = "ADAS.xlam" And hasBetaLink Then
        Application.StatusBar = ""
        ReplaceInWorkbook "E:\ADAS\Excel Add-ins\beta\ADAS_BETA.xlam", "E:\ADAS\Excel Add-ins\ADAS.xlam"
        Application.StatusBar = "ADAS - Update Completed!"
    End If
    
CleanExit:
    Application.Calculation = oldCalcMode
    skipDataProcess = False
    
End Sub



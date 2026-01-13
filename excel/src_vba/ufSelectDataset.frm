VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSelectDataset 
   Caption         =   "Load Datasets"
   ClientHeight    =   5355
   ClientLeft      =   195
   ClientTop       =   795
   ClientWidth     =   8805.001
   OleObjectBlob   =   "ufSelectDataset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSelectDataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' in-memory data from the sheet
' mData is a 2D array including headers (row 1)
Private mData As Variant
Private mColCat As Long, mColName As Long, mColFmt As Long

Private Sub UserForm_Initialize()
    Dim wb As Workbook, ws As Worksheet
    Dim oldScr As Boolean, oldEvt As Boolean, oldCalc As XlCalculation
    
    'Me.lbl1.Font.Size = 12
    
    oldScr = Application.ScreenUpdating
    oldEvt = Application.EnableEvents
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo clean_fail
    
    Set wb = Workbooks.Open(fileName:=VP_SETTINGS_PATH, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    Set ws = wb.Worksheets(VP_SETTINGS_SHEET)
    
    ' Load the used range to an array
    mData = ws.UsedRange.Value2
    
    ' Find column indices by header names
    mColCat = FindHeaderCol("Category")
    mColName = FindHeaderCol("Name")
    mColFmt = FindHeaderCol("Data Format") ' you asked to add this filter
    
    ' Build filter dropdown lists (with "All" at top)
    PopulateComboFromUnique cboCategory, mColCat
    PopulateComboFromUnique cboFormat, mColFmt
    
    ' Empty keyword to start
    txtSearch.text = ""
    
    ' Show full list initially
    ApplyFilters
    
clean_exit:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvt
    Application.ScreenUpdating = oldScr
    Exit Sub

clean_fail:
    MsgBox "Unable to load dataset list:" & vbCrLf & VP_SETTINGS_PATH & " / " & VP_SETTINGS_SHEET & vbCrLf & Err.Description, vbExclamation
    Resume clean_exit
End Sub

' ==== Filtering ====

Private Sub ApplyFilters()
    ' Applies Category, Data Format, and keyword filters; repopulates lstNames
    Dim r As Long, lastRow As Long
    Dim cat As String, fmt As String, kw As String
    Dim nm As String
    Dim bag As Object ' Scripting.Dictionary to keep names unique (optional)
    
    If IsEmpty(mData) Then Exit Sub
    lastRow = UBound(mData, 1)
    
    cat = Trim$(cboCategory.text)
    fmt = Trim$(cboFormat.text)
    kw = LCase$(Trim$(txtSearch.text))
    
    Set bag = CreateObject("Scripting.Dictionary")
    
    lstNames.Clear
    
    For r = 2 To lastRow ' skip header row
        nm = CStr(mData(r, mColName))
        If Len(nm) > 0 Then
            If MatchOrAll(mData(r, mColCat), cat) _
               And MatchOrAll(mData(r, mColFmt), fmt) _
               And ContainsCI(nm, kw) Then
                   
                   If Not bag.Exists(nm) Then
                       bag.Add nm, True
                       lstNames.AddItem nm
                   End If
            End If
        End If
    Next r
End Sub

Private Function MatchOrAll(ByVal Value As Variant, ByVal sel As String) As Boolean
    ' True if filter is blank/"All" or Value equals selection (case-insensitive)
    Dim v As String: v = CStr(Value)
    If Len(sel) = 0 Or LCase$(sel) = "all" Then
        MatchOrAll = True
    Else
        MatchOrAll = (StrComp(v, sel, vbTextCompare) = 0)
    End If
End Function

Private Function ContainsCI(ByVal hay As String, ByVal needle As String) As Boolean
    ' Case-insensitive "contains". Empty needle means True.
    If Len(needle) = 0 Then
        ContainsCI = True
    Else
        ContainsCI = (InStr(1, LCase$(hay), needle, vbTextCompare) > 0)
    End If
End Function

' ==== Populate filter combos with unique values (+ "All") ====

Private Sub PopulateComboFromUnique(cb As MSForms.ComboBox, ByVal colIdx As Long)
    Dim dict As Object, r As Long, lastRow As Long, v As String
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = UBound(mData, 1)
    For r = 2 To lastRow
        v = CStr(mData(r, colIdx))
        If Len(v) > 0 Then
            If Not dict.Exists(v) Then dict.Add v, True
        End If
    Next r
    
    cb.Clear
    cb.AddItem "All"
    
    Dim k As Variant
    For Each k In dict.Keys
        cb.AddItem CStr(k)
    Next k
    
    cb.ListIndex = 0 ' default to "All"
End Sub

Private Function FindHeaderCol(ByVal headerName As String) As Long
    ' Finds zero-based header by exact match (case-insensitive) in row 1 of mData
    Dim c As Long, lastCol As Long
    lastCol = UBound(mData, 2)
    For c = 1 To lastCol
        If StrComp(CStr(mData(1, c)), headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    Err.Raise 5, , "Header not found: " & headerName
End Function

' ==== Events: whenever a filter changes, re-apply filters ====

Private Sub cboCategory_Change()
    ApplyFilters
End Sub

Private Sub cboFormat_Change()
    ApplyFilters
End Sub

Private Sub txtSearch_Change()
    ' Live keyword filtering as the user types (initial Value is empty)
    ApplyFilters
End Sub

' Optional: double-click on a name to select immediately
Private Sub lstNames_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdSelect_Click
End Sub

' ==== Buttons ====

Private Sub cmdSelect_Click()
    Dim picked As String
    If lstNames.ListIndex >= 0 Then
        picked = CStr(lstNames.List(lstNames.ListIndex))
    Else
        picked = Trim$(txtSearch.text) ' or however you want to fall back
    End If
    If Len(picked) = 0 Then Exit Sub

    Dim tgt As Range, owner As Range
    Set tgt = ActiveCell

    ' 1) If active cell itself is ADAS formula -> update its second arg
    If tgt.HasFormula And IsADASFormula(tgt.Formula2) Then
        If UpdateADASArg(2, tgt, picked) Then Exit Sub
    End If

    ' 2) If active cell is inside a spill from an ADAS formula -> update that owner
    Set owner = FindADASOwnerForCell(tgt)
    If Not owner Is Nothing Then
        If UpdateADASArg(2, owner, picked) Then Exit Sub
    End If

    ' 3) Otherwise just write the Value to the active cell
    tgt.Value = picked
    
    Me.lstNames.SetFocus

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



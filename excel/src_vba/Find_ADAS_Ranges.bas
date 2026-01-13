Attribute VB_Name = "Find_ADAS_Ranges"
Option Explicit

Sub List_ADAS_Ranges()
    Dim ws As Worksheet
    Dim rngFormulas As Range
    Dim c As Range
    Dim dict As Object
    Dim key As String, addr As String, kind As String
    Dim arrRng As Range, spillRng As Range
    Dim k As Variant
    Dim outWS As Worksheet
    Dim rowOut As Long
    
    Application.ScreenUpdating = False
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each ws In ActiveWorkbook.Worksheets
        ' Get all formula cells on this sheet (skip if none)
        On Error Resume Next
        Set rngFormulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        
        If Not rngFormulas Is Nothing Then
            For Each c In rngFormulas
                If c.HasFormula Then
                    If InStr(1, c.formula, "ADAS", vbTextCompare) > 0 Then
                        
                        addr = ""
                        kind = ""
                        
                        ' 1) Legacy CSE array formulas
                        If c.HasArray Then
                            Set arrRng = c.CurrentArray
                            addr = arrRng.Address(True, True)
                            kind = "CSE array"
                        
                        Else
                            ' 2) Dynamic spill formulas (Excel 365+)
                            On Error Resume Next
                            Set spillRng = c.SpillingToRange
                            On Error GoTo 0
                            
                            If Not spillRng Is Nothing Then
                                addr = spillRng.Address(True, True)
                                kind = "Dynamic spill"
                            Else
                                ' 3) Single-cell (non-array) output
                                addr = c.Address(True, True)
                                kind = "Single cell"
                            End If
                        End If
                        
                        key = ws.Name & "!" & addr
                        If Not dict.Exists(key) Then
                            dict.Add key, kind
                        End If
                        
                        Set arrRng = Nothing
                        Set spillRng = Nothing
                    End If
                End If
            Next c
        End If
        
        Set rngFormulas = Nothing
    Next ws
    
    '--- output results ---
    Debug.Print "Found " & dict.Count & " unique ADAS output ranges:"
    For Each k In dict.Keys
        Debug.Print k & "  (" & dict(k) & ")"
    Next k
    
    ' Optional: write to a new worksheet
    On Error Resume Next
    Set outWS = ThisWorkbook.Worksheets("ADAS_Ranges")
    On Error GoTo 0
    If outWS Is Nothing Then
        Set outWS = ThisWorkbook.Worksheets.Add
        outWS.Name = "ADAS_Ranges"
    Else
        outWS.Cells.Clear
    End If
    
    outWS.[A1].Value = "Sheet!Range"
    outWS.[B1].Value = "Type"
    rowOut = 2
    
    For Each k In dict.Keys
        outWS.Cells(rowOut, 1).Value = k
        outWS.Cells(rowOut, 2).Value = dict(k)
        rowOut = rowOut + 1
    Next k
    
    Application.ScreenUpdating = True
End Sub


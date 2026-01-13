Attribute VB_Name = "mod_Test"
Public Sub GetReferencingCellsEx(ByVal target As Range)
    Dim ws As Worksheet
    Dim fcell As Range          ' cell containing formula
    Dim arrCell As Range        ' cell inside array/spill range
    
    Dim addr As String
    Dim arrKey As String
    Dim cellKey As String
    
    Dim formulaCells As Range
    Dim arrRange As Range       ' array output or spill range
    
    ' (Re)initialize global collections
    Set processedCells = New Collection
    Set processedArrays = New Collection
    
    Set ws = target.Parent
    addr = target.Address(False, False)   ' e.g. "A1"
    
    ' Get all formula cells on the sheet (can error if none)
    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If formulaCells Is Nothing Then Exit Sub   ' no formulas
    
    For Each fcell In formulaCells
        
        ' Skip formulas that don't mention the cell at all
        If InStr(1, fcell.formula, addr, vbTextCompare) = 0 Then GoTo NextCell
        
        ' Skip the target cell itself
        If fcell.Address = target.Address Then GoTo NextCell
        
        ' ---------- Determine the output range ----------
        Set arrRange = Nothing
        
        ' 1) Try dynamic spill range (Excel 365+)
        On Error Resume Next
        Set arrRange = fcell.spillRange
        On Error GoTo 0
        
        ' 2) If no spill, but legacy CSE array
        If arrRange Is Nothing Then
            If fcell.HasArray Then
                Set arrRange = fcell.CurrentArray
            End If
        End If
        
        ' 3) If still nothing, treat the single formula cell as its range
        If arrRange Is Nothing Then
            Set arrRange = fcell
        End If
        
        ' ---------- Add the array/output range key ----------
        arrKey = ws.Name & "!" & arrRange.Address
        If Not KeyExists(processedArrays, arrKey) Then
            processedArrays.Add arrKey, arrKey
        End If
        
        ' ---------- Add every cell inside that range ----------
        For Each arrCell In arrRange.Cells
            cellKey = ws.Name & "!" & arrCell.Address
            If Not KeyExists(processedCells, cellKey) Then
                processedCells.Add cellKey, cellKey
            End If
        Next arrCell
        
NextCell:
    Next fcell
End Sub

Public Sub Test_GetReferencingCellsEx()
    Dim v As Variant
    
    ' Run the main routine for the active cell
    GetReferencingCellsEx ActiveCell
    
    Debug.Print "=== processedArrays ==="
    For Each v In processedArrays
        Debug.Print v
    Next v
    
    Debug.Print "=== processedCells ==="
    For Each v In processedCells
        Debug.Print v
    Next v
End Sub

Sub GetArrRng()
    qqq ActiveCell.CurrentRegion.Address
End Sub

Public Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col(i)
    Next i
    CollectionToArray = arr
End Function

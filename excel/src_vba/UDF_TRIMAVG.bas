Attribute VB_Name = "UDF_TRIMAVG"
' TRIMAVG(rng, [n_year], [exclude], [weights])
' - Filters rng like lastNRow (skip: errors, empty, "", whitespace-only strings, "0" as text,
'   numeric 0, and cells formatted with Strikethrough), then keeps the last N rows.
' - exclude: if contains "hi" (any case) drop the single highest Value; if contains "lo" drop the single lowest.
' - weights: optional single-column range the same height as rng; computes a weighted average.
'   The same row-exclusions applied to rng are applied to weights (rows removed in data are also removed in weights).
'   Non-numeric or <=0 weights are ignored (that pair is dropped). If weights is "None" (default), uses simple average.

Public Function TRIMAVG(rng As Range, Optional n_year As Variant = "All", _
                        Optional exclude As String = "None", Optional weights As Variant = "None") As Variant
    Dim arr As Variant, wArr As Variant
    Dim keepIdx() As Long, idxCount As Long
    Dim r As Long, rows As Long
    Dim v As Variant
    Dim take As Long
    Dim dataVals() As Double, wtVals() As Double
    Dim i As Long, n As Long
    Dim dropHi As Boolean, dropLo As Boolean
    Dim idxMax As Long, idxMin As Long
    Dim sumW As Double, sumWX As Double
    Dim excl As String
    
    On Error GoTo Fail
    
    ' Validate data range: single column
    If rng Is Nothing Or rng.Columns.Count <> 1 Then
        TRIMAVG = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Read values
    arr = rng.Value
    rows = UBound(arr, 1)
    
    ' Optional weights: accept "None" or a same-height, single-column range
    Dim hasWeights As Boolean
    hasWeights = False
    If Not IsMissing(weights) Then
        If VarType(weights) = vbString Then
            If UCase$(Trim$(CStr(weights))) <> "NONE" Then
                TRIMAVG = CVErr(xlErrValue)
                Exit Function
            End If
        ElseIf TypeOf weights Is Range Then
            Dim wRng As Range
            Set wRng = weights
            If wRng.Columns.Count <> 1 Or wRng.rows.Count <> rng.rows.Count Then
                TRIMAVG = CVErr(xlErrValue)
                Exit Function
            End If
            wArr = wRng.Value
            hasWeights = True
        Else
            ' Unsupported weights type
            TRIMAVG = CVErr(xlErrValue)
            Exit Function
        End If
    End If
    
    ' Collect indices of rows to keep using the same exclusion logic as lastNRow
    ReDim keepIdx(1 To rows)
    idxCount = 0
    For r = 1 To rows
        v = arr(r, 1)
        
        ' Skip errors
        If IsError(v) Then GoTo NextR
        
        ' Skip empties
        If IsEmpty(v) Then GoTo NextR
        
        ' Skip "" and whitespace-only strings; treat "0" (text) as zero
        If VarType(v) = vbString Then
            If Len(Trim$(CStr(v))) = 0 Then GoTo NextR
            If Trim$(CStr(v)) = "0" Then GoTo NextR
        End If
        
        ' Skip numeric zeros
        If IsNumeric(v) Then
            If CDbl(v) = 0 Then GoTo NextR
        End If
        
        ' Skip strikethrough-formatted cells
        If rng.Cells(r).Font.Strikethrough = True Then GoTo NextR
        
        ' Keep row index
        idxCount = idxCount + 1
        keepIdx(idxCount) = r
NextR:
    Next r
    
    If idxCount = 0 Then
        TRIMAVG = CVErr(xlErrDiv0)
        Exit Function
    End If
    
    ' Determine how many to take from the bottom (last N)
    If IsMissing(n_year) Then
        take = idxCount
    ElseIf IsNumeric(n_year) Then
        take = CLng(n_year)
    ElseIf VarType(n_year) = vbString Then
        If UCase$(Trim$(CStr(n_year))) = "ALL" Or Len(Trim$(CStr(n_year))) = 0 Then
            take = idxCount
        ElseIf IsNumeric(n_year) Then
            take = CLng(n_year)
        Else
            TRIMAVG = CVErr(xlErrValue)
            Exit Function
        End If
    Else
        TRIMAVG = CVErr(xlErrValue)
        Exit Function
    End If
    
    If take <= 0 Then
        TRIMAVG = CVErr(xlErrDiv0)
        Exit Function
    ElseIf take > idxCount Then
        take = idxCount
    End If
    
    ' Build data (and weights) arrays from the last N kept indices; keep only numeric data
    ReDim dataVals(1 To take)
    If hasWeights Then ReDim wtVals(1 To take)
    
    n = 0
    For i = idxCount - take + 1 To idxCount
        r = keepIdx(i)
        If IsNumeric(arr(r, 1)) Then
            If Not IsEmpty(arr(r, 1)) Then
                n = n + 1
                dataVals(n) = CDbl(arr(r, 1))
                If hasWeights Then
                    If Not IsError(wArr(r, 1)) And IsNumeric(wArr(r, 1)) Then
                        wtVals(n) = CDbl(wArr(r, 1))
                    Else
                        wtVals(n) = 0#
                    End If
                End If
            End If
        End If
    Next i
    
    If n = 0 Then
        TRIMAVG = CVErr(xlErrDiv0)
        Exit Function
    End If
    
    ' Shrink arrays to actual count n
    If n < take Then
        ReDim Preserve dataVals(1 To n)
        If hasWeights Then ReDim Preserve wtVals(1 To n)
    End If
    
    ' Parse hi/lo exclusions
    excl = LCase$(Trim$(exclude))
    dropHi = (InStr(excl, "hi") > 0)
    dropLo = (InStr(excl, "lo") > 0)
    
    ' Drop highest once if requested
    If dropHi And n > 1 Then
        idxMax = 1
        For i = 2 To n
            If dataVals(i) > dataVals(idxMax) Then idxMax = i
        Next i
        For i = idxMax To n - 1
            dataVals(i) = dataVals(i + 1)
            If hasWeights Then wtVals(i) = wtVals(i + 1)
        Next i
        n = n - 1
        If n = 0 Then
            TRIMAVG = CVErr(xlErrDiv0)
            Exit Function
        End If
        ReDim Preserve dataVals(1 To n)
        If hasWeights Then ReDim Preserve wtVals(1 To n)
    End If
    
    ' Drop lowest once if requested
    If dropLo And n > 1 Then
        idxMin = 1
        For i = 2 To n
            If dataVals(i) < dataVals(idxMin) Then idxMin = i
        Next i
        For i = idxMin To n - 1
            dataVals(i) = dataVals(i + 1)
            If hasWeights Then wtVals(i) = wtVals(i + 1)
        Next i
        n = n - 1
        If n = 0 Then
            TRIMAVG = CVErr(xlErrDiv0)
            Exit Function
        End If
        ReDim Preserve dataVals(1 To n)
        If hasWeights Then ReDim Preserve wtVals(1 To n)
    End If
    
    ' Compute average
    If Not hasWeights Then
        ' Simple average
        sumWX = 0#
        For i = 1 To n
            sumWX = sumWX + dataVals(i)
        Next i
        TRIMAVG = sumWX / n
    Else
        ' Weighted average: drop pairs with non-positive weights
        sumW = 0#: sumWX = 0#
        For i = 1 To n
            If wtVals(i) > 0 Then
                sumW = sumW + wtVals(i)
                sumWX = sumWX + wtVals(i) * dataVals(i)
            End If
        Next i
        If sumW = 0 Then
            TRIMAVG = CVErr(xlErrDiv0)
        Else
            TRIMAVG = sumWX / sumW
        End If
    End If
    Exit Function
    
Fail:
    TRIMAVG = CVErr(xlErrValue)
End Function

' lastNRow(rng, [N])
' - rng: a single-column range
' - N:  number of rows to return from the bottom of the filtered list.
'       Use "All" (default) to return all non-excluded rows.
' Exclusions:
'   * Empty cells
'   * Formulas that return ""
'   * Numeric 0
'   * Text "0"
'   * Error values (ignored)
' Returns a vertical 2D array suitable for spilling.
Public Function lastNRow(rng As Range, Optional n As Variant = "All") As Variant
    Dim arr As Variant
    Dim r As Long, rows As Long
    Dim keep() As Variant
    Dim k As Long
    Dim v As Variant
    Dim take As Long
    Dim outArr() As Variant
    
    ' Validate: must be a single column
    If rng Is Nothing Then
        lastNRow = CVErr(xlErrRef)
        Exit Function
    End If
    If rng.Columns.Count <> 1 Then
        lastNRow = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Read values
    arr = rng.Value
    rows = UBound(arr, 1)
    
    ' Pre-allocate max size; we'll fill from 1..k
    ReDim keep(1 To rows, 1 To 1)
    k = 0
    
    For r = 1 To rows
        v = arr(r, 1)
        
        ' Skip errors
        If IsError(v) Then GoTo NextR
        
        ' Skip true empties
        If IsEmpty(v) Then GoTo NextR
        
        ' Skip strikethrough formatting
        If rng.Cells(r).Font.Strikethrough = True Then GoTo NextR

        ' Skip formulas that return "", or strings that are just whitespace
        If VarType(v) = vbString Then
            If Len(Trim$(CStr(v))) = 0 Then GoTo NextR
            ' Treat "0" (as text) as zero
            If Trim$(CStr(v)) = "0" Then GoTo NextR
        End If
        
        ' Skip numeric zeros
        If IsNumeric(v) Then
            If CDbl(v) = 0 Then GoTo NextR
        End If
        
        ' Keep this Value
        k = k + 1
        keep(k, 1) = v
        
NextR:
    Next r
    
    ' If nothing to return
    If k = 0 Then
        lastNRow = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' Determine how many to take from the bottom
    If IsMissing(n) Then
        take = k
    ElseIf VarType(n) = vbString Then
        If UCase$(Trim$(CStr(n))) = "ALL" Or Len(Trim$(CStr(n))) = 0 Then
            take = k
        ElseIf IsNumeric(n) Then
            take = CLng(n)
        Else
            lastNRow = CVErr(xlErrValue)
            Exit Function
        End If
    ElseIf IsNumeric(n) Then
        take = CLng(n)
    Else
        lastNRow = CVErr(xlErrValue)
        Exit Function
    End If
    
    If take < 0 Then
        lastNRow = CVErr(xlErrValue)
        Exit Function
    ElseIf take = 0 Then
        lastNRow = CVErr(xlErrNA)
        Exit Function
    ElseIf take > k Then
        take = k
    End If
    
    ' Build the output (last N of the kept list)
    ReDim outArr(1 To take, 1 To 1)
    Dim i As Long
    For i = 1 To take
        outArr(i, 1) = keep(k - take + i, 1)
    Next i
    
    lastNRow = outArr
End Function





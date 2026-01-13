Attribute VB_Name = "Helper"
Option Private Module

Private Sub EnableDebugMode()
    debugMode = True
End Sub

Private Sub DisableDebugMode()
    debugMode = False
End Sub

Sub Check_DataProcess()
    qqq "-----------"
    qqq "errCount: " & errCount
    If pendingUpdate Then Debug.Print "pendingUpdate"
    If skipDataProcess Then Debug.Print "skipDataProcess!"
    If disableRequest Then Debug.Print "Request Disabled!"
    If disableWaitTime Then Debug.Print "WaitTime Disabled!"
End Sub

Public Function FormatYYYYMM_ToMmmYYYY(arr As Variant) As Variant
    Dim r As Long, c As Long
    Dim s As String
    Dim yyyy As Integer, mm As Integer

    If Not IsArray(arr) Then
        FormatYYYYMM_ToMmmYYYY = arr
        Exit Function
    End If

    For r = LBound(arr, 1) To UBound(arr, 1)
        For c = LBound(arr, 2) To UBound(arr, 2)
            If Not IsError(arr(r, c)) Then
                s = Trim$(CStr(arr(r, c)))
                If Len(s) = 6 And IsNumeric(s) Then
                    yyyy = CInt(Left$(s, 4))
                    mm = CInt(Right$(s, 2))
                    If mm >= 1 And mm <= 12 Then
                        arr(r, c) = Format$(DateSerial(yyyy, mm, 1), "mmm yyyy")
                    End If
                End If
            End If
        Next c
    Next r

    FormatYYYYMM_ToMmmYYYY = arr
End Function

Function RateLimited(ByVal key As String, _
                             Optional ByVal maxCalls As Long = 5, _
                             Optional ByVal windowSec As Double = 3#) As Boolean
    Static dict As Object  ' Scripting.Dictionary
    Dim nowT As Double, q As Collection, i As Long

    If dict Is Nothing Then Set dict = CreateObject("Scripting.Dictionary")

    nowT = Timer  ' seconds since midnight

    If Not dict.Exists(key) Then
        Set q = New Collection
        dict.Add key, q
    Else
        Set q = dict(key)
    End If

    ' Drop timestamps older than the window (rolling window)
    For i = q.Count To 1 Step -1
        If (nowT - CDbl(q(i))) > windowSec Then q.Remove i
    Next i

    ' If already at limit, block
    If q.Count >= maxCalls Then
        RateLimited = True
        Exit Function
    End If

    ' Otherwise record this call
    q.Add nowT
    RateLimited = False
End Function

Sub StatusBar(ByVal strMsg As String)
    Application.StatusBar = False
    DoEvents
    Application.StatusBar = strMsg
    DoEvents
End Sub

Public Sub WaitSec(ByVal seconds As Double)
    Dim t As Double
    t = Timer
    Do While Timer < t + seconds
        DoEvents
    Loop
End Sub

Public Function TransposeArray(arr As Variant) As Variant
    Dim outArr() As Variant
    Dim r As Long, c As Long
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long

    lb1 = LBound(arr, 1): ub1 = UBound(arr, 1)
    lb2 = LBound(arr, 2): ub2 = UBound(arr, 2)
    ReDim outArr(1 To ub2 - lb2 + 1, 1 To ub1 - lb1 + 1)

    For r = lb1 To ub1
        For c = lb2 To ub2
            outArr(c - lb2 + 1, r - lb1 + 1) = arr(r, c)
        Next c
    Next r

    TransposeArray = outArr
End Function

Public Sub BuildTriangle(ByVal rowCount As Long, ByVal colCount As Long)
    Dim i As Long
    Dim j As Long
    Dim stepSize As Double
    Dim rng As Range
    Dim startCell As Range

    Set startCell = ActiveCell
    
    stepSize = colCount / rowCount

    ' Loop through the rows
    For i = 0 To rowCount - 1
        ' Calculate the number of columns for this row
        Dim currentCols As Long
        currentCols = colCount - (stepSize * i)

        Set rng = startCell.Resize(1, currentCols).Offset(i, 0) ' Set the range to fill

        startCell.Copy
        rng.PasteSpecial Paste:=xlPasteFormulas
    Next i

    Application.CutCopyMode = False
End Sub

Function TimeMS() As String
    TimeMS = Format(Now, "hh:mm:ss") & _
        "(" & Right(Format(Timer - Int(Timer), "0.000"), 3) & ")"
End Function

Public Function WaitForFileReady(ByVal filePath As String, _
                                 ByVal maxWaitSeconds As Double) As Boolean
    Dim deadline As Date
    Dim ff As Integer
    Dim line As String
    
    Show_ufLoading
    
    deadline = DateAdd("s", maxWaitSeconds, Now)
    
    Do While Now < deadline
        If Dir(filePath) <> "" Then
            ff = FreeFile
            On Error Resume Next
            Open filePath For Input As #ff
            If Err.Number = 0 Then
                ' File exists and can be opened for reading
                Close #ff
                On Error GoTo 0
                WaitForFileReady = True
                Exit Function
            Else
                ' Exists but still locked / being written
                Err.Clear
                On Error GoTo 0
            End If
        End If
        DoEvents
    Loop
    
    ' Timed out
    WaitForFileReady = False
End Function

'Row-preserving: tri row 1 -> out row 1, tri row i -> out row i
'Latest = rightmost non-empty; DiagonalIndex=1 means 2nd rightmost non-empty, etc.
Public Function GetDiagonal(tri As Variant, Optional DiagonalIndex As Long = 0) As Variant
    Dim rLB As Long, rUB As Long, cLB As Long, cUB As Long
    Dim r As Long, c As Long
    Dim outArr() As Variant
    Dim cnt As Long, target As Long
    Dim v As Variant

    If DiagonalIndex < 0 Then DiagonalIndex = 0
    target = DiagonalIndex + 1

    'Expect 2D input (Range.Value2). If a 1D array is passed, fail to fallback.
    On Error GoTo Fallback
    rLB = LBound(tri, 1): rUB = UBound(tri, 1)
    cLB = LBound(tri, 2): cUB = UBound(tri, 2)
    On Error GoTo 0

    'Important: output is indexed the SAME way as tri's row index.
    ReDim outArr(rLB To rUB, 1 To 1)

    For r = rLB To rUB
        cnt = 0
        outArr(r, 1) = 0 'default if no value / not enough values
        qqq r
        For c = cUB To cLB Step -1
            v = tri(r, c)

            If IsNonEmptyValue(v) Then
                qqq v
                cnt = cnt + 1
                If cnt = target Then
                    outArr(r, 1) = v
                    Exit For
                End If
            End If
        Next c
    Next r

    GetDiagonal = outArr
    Exit Function

Fallback:
    ReDim outArr(1 To 1, 1 To 1)
    If DiagonalIndex = 0 And IsNonEmptyValue(tri) Then
        outArr(1, 1) = tri
    Else
        outArr(1, 1) = 0
    End If
    GetDiagonal = outArr
End Function

Private Function IsNonEmptyValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function

    If VarType(v) = vbString Then
        IsNonEmptyValue = (Len(Trim$(CStr(v))) > 0)
    Else
        IsNonEmptyValue = True
    End If
End Function

' ref like:  "Sheet Name!$B$4:$C$11"
Public Sub RefreshADASBlock(ByVal ref As String)
    Dim parts() As String
    Dim shName As String, addr As String
    Dim ws As Worksheet
    Dim rng As Range, topCell As Range
    
    parts = Split(ref, "!")
    If UBound(parts) <> 1 Then Exit Sub
    
    shName = parts(0)
    addr = parts(1)
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(shName)
    If ws Is Nothing Then Exit Sub
    
    ' ws.Activate
    
    Set rng = ws.Range(addr)
    If rng Is Nothing Then Exit Sub
    
    Set topCell = rng.Cells(1, 1)   ' works for single / CSE / spill
    'topCell.Select
    
    ' 1) Legacy CSE array: re-enter entire array
    If topCell.HasArray Then
        topCell.CurrentArray.FormulaArray = topCell.CurrentArray.FormulaArray
    
    ' 2) Dynamic spill or normal single formula:
    '    re-enter only the formula cell
    Else
        On Error Resume Next
        topCell.Formula2 = topCell.Formula2   ' Excel 365+
        If Err.Number <> 0 Then
            Err.Clear
            topCell.formula = topCell.formula ' fallback for older Excel
        End If
        On Error GoTo 0
    End If
End Sub

Public Function GetParamValue(ByVal fullStr As String, ByVal paramName As String) As String
    Dim parts() As String
    Dim i As Long, pair As String, p As Long

    parts = Split(fullStr, "#")
    
    For i = LBound(parts) To UBound(parts)
        pair = Trim(parts(i))
        p = InStr(1, pair, paramName & " = ", vbTextCompare)
        
        If p = 1 Then
            ' Extract the value after "="
            GetParamValue = Trim(Mid(pair, Len(paramName) + 4))
            Exit Function
        End If
    Next i
    
    GetParamValue = ""  ' Not found
End Function

Sub RemoveAllSheetsFromAddin()
    Dim ws As Worksheet
    Application.DisplayAlerts = False

    For Each ws In ThisWorkbook.Worksheets
        ws.Delete
    Next ws

    Application.DisplayAlerts = True
End Sub

Sub qqq(ByVal var1 As String)
    Debug.Print var1
End Sub


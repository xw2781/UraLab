Attribute VB_Name = "mod_SelectDataset"
Option Private Module
Option Explicit

' Read Category/Name columns (by header text) into arrays.
' Returns the number of rows loaded; arrays are 1-based.
Public Function LoadDatasetTypes( _
        ByVal Path As String, _
        ByVal sheetName As String, _
        ByRef names() As String, _
        ByRef cats() As String) As Long

    Dim wb As Workbook, ws As Worksheet
    Dim ur As Range, v As Variant
    Dim r As Long, cName As Long, cCat As Long
    Dim n As Long, lastRow As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Open invisibly to keep UI calm
    Set wb = Application.Workbooks.Open(Path, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    Set ws = wb.Worksheets(sheetName)
    Set ur = ws.UsedRange
    v = ur.Value

    ' Find headers in the first row of UsedRange
    cName = FindHeaderCol(v, "Name")
    cCat = FindHeaderCol(v, "Category")
    If cName = 0 Or cCat = 0 Then Err.Raise 5, , "Could not find 'Name' and/or 'Category' headers."

    lastRow = UBound(v, 1)
    ReDim names(1 To lastRow - 1)
    ReDim cats(1 To lastRow - 1)

    ' Rows after header
    For r = 2 To lastRow
        If Len(Trim$(CStr(v(r, cName)))) > 0 Then
            n = n + 1
            names(n) = CStr(v(r, cName))
            cats(n) = CStr(v(r, cCat))
        End If
    Next

    If n = 0 Then
        Erase names: Erase cats
        LoadDatasetTypes = 0
    Else
        ReDim Preserve names(1 To n)
        ReDim Preserve cats(1 To n)
        LoadDatasetTypes = n
    End If

CleanUp:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Function

ErrHandler:
    MsgBox "Load failed: " & Err.Description, vbExclamation
    Resume CleanUp
End Function

' Helper: find a header title (case-insensitive) in row 1 of a 2-D variant array
Private Function FindHeaderCol(ByVal v As Variant, ByVal header As String) As Long
    Dim j As Long, ub As Long
    ub = UBound(v, 2)
    For j = 1 To ub
        If StrComp(CStr(v(1, j)), header, vbTextCompare) = 0 Then
            FindHeaderCol = j
            Exit Function
        End If
    Next
    FindHeaderCol = 0
End Function

'-------------------------
' Argument splitting utils
'-------------------------

' Splits a function argument list into a 1-based Collection of strings.
' Respects quotes and nested parentheses (so commas inside "..." or (...) are ignored).
Private Function SplitArgs(ByVal s As String) As Collection
    Dim i As Long, ch As String * 1
    Dim inQuotes As Boolean, depth As Long
    Dim buf As String
    Dim col As New Collection
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case """"
                inQuotes = Not inQuotes
                buf = buf & ch
            Case "("
                depth = depth + 1
                buf = buf & ch
            Case ")"
                If depth > 0 Then depth = depth - 1
                buf = buf & ch
            Case ","
                If Not inQuotes And depth = 0 Then
                    col.Add Trim$(buf)
                    buf = ""
                Else
                    buf = buf & ch
                End If
            Case Else
                buf = buf & ch
        End Select
    Next i
    
    If Len(Trim$(buf)) > 0 Or Len(s) = 0 Then
        col.Add Trim$(buf)
    End If
    
    Set SplitArgs = col
End Function

' Joins a 1-based Collection of argument strings with ", "
Private Function JoinArgs(ByVal args As Collection) As String
    Dim i As Long, out As String
    For i = 1 To args.Count
        If i > 1 Then out = out & ", "
        out = out & args(i)
    Next
    JoinArgs = out
End Function

'=== Helpers ======================================================

Public Function IsADASFormula(ByVal f As String) As Boolean
    Dim u As String: u = UCase$(Trim$(f))
    IsADASFormula = (Left$(u, 9) = "=ADASTRI(") Or (Left$(u, 9) = "=ADASVEC(")
End Function

' Return the top-left “owner” cell if ActiveCell is inside a spilled ADAS output;
' otherwise Nothing.
Public Function FindADASOwnerForCell(ByVal c As Range) As Range
    On Error Resume Next

    ' 1) If c itself is an ADAS formula, it’s the owner
    If c.HasFormula Then
        If IsADASFormula(c.Formula2) Then
            Set FindADASOwnerForCell = c
            Exit Function
        End If
    End If

    ' 2) Look for a nearby spilled ADAS formula whose spill range contains c.
    '    Limit the search to CurrentRegion to keep it snappy.
    Dim region As Range
    If c.CurrentRegion Is Nothing Then
        Set region = c.Parent.UsedRange
    Else
        Set region = c.CurrentRegion
    End If

    Dim fc As Range, formulas As Range
    On Error Resume Next
    Set formulas = region.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not formulas Is Nothing Then
        For Each fc In formulas.Cells
            If IsADASFormula(fc.Formula2) Then
                If fc.HasSpill Then
                    Dim spill As Range
                    On Error Resume Next
                    Set spill = fc.SpillingToRange
                    On Error GoTo 0
                    If Not spill Is Nothing Then
                        If Not Application.Intersect(c, spill) Is Nothing Then
                            Set FindADASOwnerForCell = fc
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next fc
    End If
End Function

' Replace the N-th argument in the ADAS formula:
'   - If the N-th arg is a string literal, replace it in the formula text.
'   - If the N-th arg is a reference, write newVal into that reference and
'     leave the formula unchanged.
Public Function UpdateADASArg(ByVal n_th_arg As Long, _
                              ByVal owner As Range, _
                              ByVal newVal As String) As Boolean
    Dim f As String
    f = owner.Formula2
    If Not IsADASFormula(f) Then Exit Function
    If n_th_arg < 1 Then Exit Function

    Dim i As Long, lvl As Long, inQ As Boolean
    Dim ch As String * 1
    Dim openPos As Long, closePos As Long

    ' Find opening "(" of the function call
    openPos = InStr(1, f, "(", vbTextCompare)
    If openPos = 0 Then Exit Function

    ' Collect top-level commas and find the closing ")"
    Dim commas() As Long
    Dim commaCount As Long
    ReDim commas(1 To 1)

    For i = openPos + 1 To Len(f)
        ch = Mid$(f, i, 1)

        If ch = """" Then
            inQ = Not inQ

        ElseIf Not inQ Then
            If ch = "(" Then
                lvl = lvl + 1

            ElseIf ch = ")" Then
                If lvl = 0 Then
                    closePos = i
                    Exit For
                Else
                    lvl = lvl - 1
                End If

            ElseIf ch = "," And lvl = 0 Then
                commaCount = commaCount + 1
                If commaCount > UBound(commas) Then
                    ReDim Preserve commas(1 To commaCount)
                End If
                commas(commaCount) = i
            End If
        End If
    Next i

    If closePos = 0 Then Exit Function

    ' Number of arguments = commas + 1
    Dim numArgs As Long
    numArgs = commaCount + 1
    If n_th_arg > numArgs Then Exit Function

    ' Determine start/end of the N-th argument (1-based)
    Dim startPos As Long, endPos As Long

    If n_th_arg = 1 Then
        startPos = openPos + 1
    Else
        startPos = commas(n_th_arg - 1) + 1
    End If

    If n_th_arg <= commaCount Then
        endPos = commas(n_th_arg) - 1
    Else
        endPos = closePos - 1
    End If

    Dim argText As String, argTrim As String
    argText = Mid$(f, startPos, endPos - startPos + 1)
    argTrim = Trim$(argText)

    ' Case 1: N-th arg is a quoted string -> replace in formula text
    If Len(argTrim) >= 2 _
       And Left$(argTrim, 1) = """" _
       And Right$(argTrim, 1) = """" Then

        Dim esc As String
        esc = Replace(newVal, """", """""")  ' escape quotes

        Dim leftPart As String, rightPart As String
        leftPart = Left$(f, startPos - 1)
        rightPart = Mid$(f, endPos + 1)

        owner.Formula2 = leftPart & """" & esc & """" & rightPart
        UpdateADASArg = True
        Exit Function
    End If

    ' Case 2: N-th arg is a reference -> write the Value to that ref
    Dim refTxt As String
    Dim refRng As Range
    refTxt = argTrim

    On Error Resume Next
    Set refRng = owner.Parent.Evaluate(refTxt)
    If refRng Is Nothing Then Set refRng = Application.Evaluate(refTxt)
    On Error GoTo 0

    If Not refRng Is Nothing Then
        refRng.Value = newVal
        UpdateADASArg = True
    Else
        ' Fallback: if we can't resolve the ref, switch to string literal
        Dim esc2 As String
        Dim leftPart2 As String, rightPart2 As String

        esc2 = Replace(newVal, """", """""")
        leftPart2 = Left$(f, startPos - 1)
        rightPart2 = Mid$(f, endPos + 1)

        owner.Formula2 = leftPart2 & """" & esc2 & """" & rightPart2
        UpdateADASArg = True
    End If
End Function


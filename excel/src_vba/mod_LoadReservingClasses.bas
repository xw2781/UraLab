Attribute VB_Name = "mod_LoadReservingClasses"
Option Private Module
Option Explicit

'--- Parse a CSV line into exactly expectedCols fields (handles quotes)
Public Function ParseCSVLine(ByVal line As String, ByVal expectedCols As Long) As Variant
    Dim i As Long, ch As String * 1, inQuotes As Boolean
    Dim fields() As String, buf As String, colCount As Long
    
    ReDim fields(1 To expectedCols)
    colCount = 1
    
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        Select Case ch
            Case """"                    ' toggle quoting
                inQuotes = Not inQuotes
            Case ","                     ' next field if not quoted
                If Not inQuotes Then
                    fields(colCount) = buf: buf = ""
                    colCount = colCount + 1
                    If colCount > expectedCols Then Exit For
                Else
                    buf = buf & ch
                End If
            Case Else
                buf = buf & ch
        End Select
    Next
    If colCount <= expectedCols Then fields(colCount) = buf
    
    ' trim & unquote each field
    For i = 1 To expectedCols
        fields(i) = Trim$(fields(i))
        If Len(fields(i)) >= 2 Then
            If Left$(fields(i), 1) = """" And Right$(fields(i), 1) = """" Then
                fields(i) = Mid$(fields(i), 2, Len(fields(i)) - 2)
            End If
        End If
    Next
    
    ParseCSVLine = fields
End Function

'--- Load headers (row1), defaults (row2), and unique Value lists (rows 2..end)
'    headersOut(1..nCols) : strings from row 1
'    defaultsOut(1..nCols): strings from row 2 (or "" if missing)
'    colArraysOut(1..nCols): 0-based arrays of unique strings from rows 2..end
Public Sub LoadCSV5(ByVal Path As String, ByVal nCols As Long, _
                    ByRef headersOut As Variant, ByRef defaultsOut As Variant, _
                    ByRef colArraysOut As Variant)
    Dim fh As Integer, line As String, rowNum As Long, i As Long
    Dim dicts() As Object
    
    ReDim headersOut(1 To nCols)
    ReDim defaultsOut(1 To nCols)
    ReDim dicts(1 To nCols)
    For i = 1 To nCols
        Set dicts(i) = CreateObject("Scripting.Dictionary")
    Next
    
    fh = FreeFile
    Open Path For Input As #fh
    Do While Not EOF(fh)
        Line Input #fh, line
        If Len(line) > 0 Then
            rowNum = rowNum + 1
            Dim fields As Variant
            fields = ParseCSVLine(line, nCols)
            
            If rowNum = 1 Then
                ' headers
                For i = 1 To nCols
                    headersOut(i) = fields(i)
                Next
            Else
                ' defaults (second row) and unique lists
                If rowNum = 2 Then
                    For i = 1 To nCols
                        defaultsOut(i) = fields(i)
                    Next
                End If
                For i = 1 To nCols
                    If Len(fields(i)) > 0 Then
                        If Not dicts(i).Exists(fields(i)) Then dicts(i).Add fields(i), True
                    End If
                Next
            End If
        End If
    Loop
    Close #fh
    
    ReDim colArraysOut(1 To nCols)
    For i = 1 To nCols
        If dicts(i).Count > 0 Then
            colArraysOut(i) = dicts(i).Keys  ' 0-based
        Else
            colArraysOut(i) = Array()
        End If
    Next
End Sub



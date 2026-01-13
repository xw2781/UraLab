Attribute VB_Name = "Metadata"
Option Explicit

' Returns a 1x2 array:
'   (1) file size in KB
'   (2) last modified datetime
' If file not found (or error), returns #N/A
Public Function CsvFileInfoKB(ByVal filePath As String) As Variant
    On Error GoTo EH

    Dim p As String
    p = Trim$(filePath)

    If Len(p) = 0 Then GoTo EH
    If Len(Dir$(p, vbNormal)) = 0 Then GoTo EH

    Dim outArr(1 To 1, 1 To 2) As Variant
    outArr(1, 1) = FileLen(p) / 1024#
    outArr(1, 2) = FileDateTime(p)

    CsvFileInfoKB = outArr
    Exit Function

EH:
    CsvFileInfoKB = CVErr(xlErrNA)
End Function


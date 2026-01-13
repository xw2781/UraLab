Attribute VB_Name = "Module1"
Option Explicit

Private Const REPO_ROOT As String = "E:\ADAS\repos\ADAS-Actuarial-Data-Analysis-System"
Private Const OUT_DIR As String = "excel\src_vba"

Public Sub ExportProjectToRepo()
    Dim vbProj As Object ' VBIDE.VBProject
    Dim outPath As String

    Set vbProj = Application.VBE.ActiveVBProject
    outPath = CombinePath(REPO_ROOT, OUT_DIR)

    EnsureFolderExists outPath
    CleanExportFolder outPath

    ExportVBProject vbProj, outPath

    MsgBox "Export complete: " & outPath, vbInformation
End Sub

Private Sub ExportVBProject(ByVal vbProj As Object, ByVal outPath As String)
    Dim vbComp As Object ' VBIDE.VBComponent
    Dim ext As String
    Dim fileName As String
    Dim exportPath As String

    For Each vbComp In vbProj.VBComponents
        ext = ComponentExtension(vbComp.Type)
        If Len(ext) = 0 Then
            ' Skip unsupported types
        Else
            fileName = SafeFileName(vbComp.Name) & ext
            exportPath = CombinePath(outPath, fileName)

            On Error Resume Next
            Kill exportPath
            On Error GoTo 0

            vbComp.Export exportPath
        End If
    Next vbComp

    ExportDocumentModules vbProj, outPath
End Sub

' Some Excel host document modules may not export reliably via VBComponents in certain setups.
' This helper exports them as .cls using the CodeModule text as fallback.
Private Sub ExportDocumentModules(ByVal vbProj As Object, ByVal outPath As String)
    Dim doc As Object ' VBIDE.VBComponent
    Dim cm As Object ' VBIDE.CodeModule
    Dim codeText As String
    Dim filePath As String

    ' ThisWorkbook
    Set doc = vbProj.VBComponents("ThisWorkbook")
    If Not doc Is Nothing Then
        Set cm = doc.CodeModule
        codeText = cm.lines(1, cm.CountOfLines)
        filePath = CombinePath(outPath, "ThisWorkbook.cls")
        WriteTextFile filePath, codeText
    End If
End Sub

Private Function ComponentExtension(ByVal compType As Long) As String
    ' VBIDE.vbext_ComponentType:
    ' 1 = vbext_ct_StdModule  -> .bas
    ' 2 = vbext_ct_ClassModule -> .cls
    ' 3 = vbext_ct_MSForm -> .frm (+ .frx)
    ' 100 = vbext_ct_Document -> handled separately
    Select Case compType
        Case 1: ComponentExtension = ".bas"
        Case 2: ComponentExtension = ".cls"
        Case 3: ComponentExtension = ".frm"
        Case Else: ComponentExtension = vbNullString
    End Select
End Function

Private Sub CleanExportFolder(ByVal folderPath As String)
    Dim f As String
    f = Dir(CombinePath(folderPath, "*.*"))
    Do While Len(f) > 0
        If LCase$(f) <> ".gitkeep" Then
            On Error Resume Next
            Kill CombinePath(folderPath, f)
            On Error GoTo 0
        End If
        f = Dir()
    Loop
End Sub

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim parts() As String
    Dim i As Long
    Dim cur As String

    parts = Split(folderPath, "\")
    cur = parts(0) & "\"
    For i = 1 To UBound(parts)
        cur = cur & parts(i) & "\"
        If Len(Dir(cur, vbDirectory)) = 0 Then
            MkDir cur
        End If
    Next i
End Sub

Private Function CombinePath(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Then
        CombinePath = a & b
    Else
        CombinePath = a & "\" & b
    End If
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SafeFileName = s
    For i = LBound(badChars) To UBound(badChars)
        SafeFileName = Replace(SafeFileName, badChars(i), "_")
    Next i
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal text As String)
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, text
    Close #ff
End Sub


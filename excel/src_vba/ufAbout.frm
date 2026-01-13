VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAbout 
   Caption         =   "ADAS Excel Add-in"
   ClientHeight    =   2415
   ClientLeft      =   255
   ClientTop       =   1020
   ClientWidth     =   4560
   OleObjectBlob   =   "ufAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Label2.Caption = "Version: " & ADAS_VERSION
    Label3.ForeColor = RGB(100, 100, 100)
End Sub

Private Sub Label3_Click()
    OpenWordReadOnly "E:\ADAS\library\Version Track.docx"
End Sub

Sub OpenWordReadOnly(ByVal fullPath As String)
    Dim wdApp As Object
    Dim wdDoc As Object
    
    If Dir(fullPath) = "" Then
        MsgBox "File not found:" & vbCrLf & fullPath, vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")  ' attach to existing Word instance
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(fileName:=fullPath, ReadOnly:=True)
   
    wdApp.Activate
    wdApp.WindowState = 0   'wdWindowStateNormal
    
End Sub


Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label3.ForeColor = vbBlue
    ' Label3.Font.Name = "Aptos Display"
    Me.MousePointer = fmMousePointerHand   ' ? not available!
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label3.ForeColor = RGB(100, 100, 100)
    ' Label3.Font.Name = "Aptos Narrow"
    Me.MousePointer = fmMousePointerDefault
End Sub


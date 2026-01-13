VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufLoading 
   Caption         =   "ADAS"
   ClientHeight    =   1800
   ClientLeft      =   375
   ClientTop       =   1485
   ClientWidth     =   4320
   OleObjectBlob   =   "ufLoading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Reset()
    Me.Label1.Caption = "Loading dataset, please wait ..."
End Sub

Public Sub UpdateText(ByVal msg As String)
    Me.Label1.Caption = msg
    Me.Repaint
End Sub

Private Sub CommandButton1_Click()
    skipDataProcess = True
    Application.StatusBar = "Force stopped by user, use 'Load and Calculate' in ADAS menu to reset."
    Unload Me
End Sub

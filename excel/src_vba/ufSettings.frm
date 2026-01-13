VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSettings 
   Caption         =   "User Settings - ADAS"
   ClientHeight    =   7230
   ClientLeft      =   195
   ClientTop       =   795
   ClientWidth     =   6180
   OleObjectBlob   =   "ufSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
  ' Get Initial Values
    LoadConfig
    ComboBox1.Value = "E:\ADAS\Virtual Projects\ResQ - Channel.xlsx"
    ComboBox2.Value = "E:\ADAS\Team Profile\Actuarial_NJ.xlsm"
    
    LoadFilePaths ComboBox1, "E:\ADAS\Virtual Projects\"
    
    If removeData Then
        OptionButton1.Value = False
        OptionButton2.Value = True
    Else
        OptionButton1.Value = True
        OptionButton2.Value = False
    End If
    qqq disableProgressBar
    If disableProgressBar Then
        CheckBox5.Value = True
    Else
        CheckBox5.Value = False
    End If
End Sub

' +--------+
' | Page 1 |
' +--------+
Private Sub OptionButton1_Click()
    removeData = False
    UpdateConfigValue "removeData", "False"
    Label1.Visible = False
End Sub

Private Sub OptionButton2_Click()
    removeData = True
    UpdateConfigValue "removeData", "True"
    Label1.Visible = True
End Sub

Private Sub cmdb2_Click()
    Unload ufSettings
End Sub

' Disable UI Animations for Better Performance
Private Sub CheckBox5_Click()
    disableProgressBar = CheckBox5.Value
    UpdateConfigValue "disableProgressBar", CheckBox5.Value
End Sub

' +--------+
' | Page 2 |
' +--------+

' Virtual Project Settings

Private Sub CommandButton1_Click()
    OpenFileFromCombo Me.ComboBox1
End Sub

' Team Profile

Private Sub CommandButton2_Click()
    OpenFileFromCombo Me.ComboBox2
End Sub

' +--------+
' | Helper |
' +--------+

Public Sub LoadFilePaths(cb As MSForms.ComboBox, ByVal folderPath As String)
    Dim fileName As String
    
    ' Ensure trailing slash
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    cb.Clear
    
    fileName = Dir(folderPath & "*.*")
    Do While fileName <> ""
        cb.AddItem folderPath & fileName
        fileName = Dir
    Loop

End Sub


Public Sub OpenFileFromCombo(cb As MSForms.ComboBox)
    Dim f As String
    
    f = Trim(cb.Value)    ' selected file path
    
    If f = "" Then
        MsgBox "Please select a file first.", vbExclamation, "No File Selected"
        Exit Sub
    End If
    
    If Dir(f) = "" Then
        MsgBox "File not found:" & vbCrLf & f, vbCritical, "Error"
        Exit Sub
    End If
    
    ' Open using default associated application
    ThisWorkbook.FollowHyperlink f
    
End Sub


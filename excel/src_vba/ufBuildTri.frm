VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufBuildTri 
   Caption         =   "Triangle Tool"
   ClientHeight    =   2340
   ClientLeft      =   210
   ClientTop       =   885
   ClientWidth     =   3120
   OleObjectBlob   =   "ufBuildTri.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufBuildTri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Me.TextBox1.text = triangle_tool_row
    Me.TextBox2.text = triangle_tool_col
End Sub

Private Sub cmdSubmit_Click()

    Dim rows As Long
    Dim cols As Long
    
    rows = CLng(Me.TextBox1.text)
    cols = CLng(Me.TextBox2.text)
    
    triangle_tool_row = rows
    triangle_tool_col = cols

    Call BuildTriangle(rows, cols)
    Unload Me
    
End Sub



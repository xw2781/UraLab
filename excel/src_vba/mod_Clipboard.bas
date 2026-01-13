Attribute VB_Name = "mod_Clipboard"

Option Explicit

' ---------- WinAPI clipboard (Unicode) DECLARES (MUST be at top) ----------
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr

    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long

    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const GMEM_MOVEABLE As Long = &H2&
Private Const CF_UNICODETEXT As Long = 13&

' ---------- Public entry point ----------
Public Function CopyText(ByVal s As String) As Boolean
    ' Try MSForms (if present)
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("MSForms.DataObject")
    If Not o Is Nothing Then
        o.SetText s
        o.PutInClipboard
        CopyText = True
        Exit Function
    End If
    On Error GoTo 0

    ' Fallback: WinAPI method
    CopyText = CopyText_WinAPI(s)
End Function

' ---------- Implementation ----------
Private Function CopyText_WinAPI(ByVal s As String) As Boolean
    Dim cb As LongPtr, hGlobal As LongPtr, pGlobal As LongPtr, ok As Boolean

    cb = (Len(s) + 1) * 2  ' UTF-16 bytes incl. null
    hGlobal = GlobalAlloc(GMEM_MOVEABLE, cb)
    If hGlobal = 0 Then Exit Function

    pGlobal = GlobalLock(hGlobal)
    If pGlobal = 0 Then
        GlobalFree hGlobal
        Exit Function
    End If

    CopyMemory pGlobal, StrPtr(s), cb
    GlobalUnlock hGlobal

    If OpenClipboard(0) <> 0 Then
        If EmptyClipboard() <> 0 Then
            ok = (SetClipboardData(CF_UNICODETEXT, hGlobal) <> 0)
        End If
        CloseClipboard
    End If

    If Not ok Then GlobalFree hGlobal
    CopyText_WinAPI = ok
End Function



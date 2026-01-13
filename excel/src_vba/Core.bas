Attribute VB_Name = "Core"
Option Private Module
Option Explicit

Public Const ADAS_VERSION As String = "1.8.2"

' User Specific Config (C:\User\...\ADAS\config.txt)
Public configDir As String
Public configPath As String
Public removeData As Boolean
Public disable_ufLoading As Boolean
Public teamProfile As String
Public debugMode As Boolean
Public disableProgressBar As Boolean

' Internal Controls
Public disableRequest As Boolean
Public disableWaitTime As Boolean
Public skipDataProcess As Boolean
Public maxWaitTime As Single
Public errCount As Integer
Public lastRequestInfo As String

Public processedCells As New Collection
Public processedArrays As New Collection
Public cancelUpdate As Boolean
Public pendingUpdate As Boolean
Public doubleRefresh As Boolean
Public disableWatcher As Boolean

' UI Select Dataset
Public Const VP_SETTINGS_PATH As String = "E:\ADAS\Virtual Projects\ResQ - Channel.xlsx"
Public Const VP_SETTINGS_SHEET As String = "Dataset Types"

Public triangle_tool_row As Long
Public triangle_tool_col As Long

Public Function GetDataset(funcArgs As String)
' +---------------+
' | Main Function |
' +---------------+
    Dim dataPath As String
    Dim t1 As Double, t2 As Double
    Dim requestInfo As String
    Const MAX_WAIT_SEC As Double = 5
    On Error GoTo ErrHandler
    
    If skipDataProcess Then
        Exit Function
    End If
    
    ' t1 = Timer
    ' Debug.Print "Time - Start: " & TimeMS()
    
    dataPath = SetDataPath(funcArgs)
    requestInfo = funcArgs & "#DataPath = " & dataPath
    
    ' --- Case 1: reuse existing data if allowed ---
    If (Dir(dataPath) <> "") And (removeData = False) Then
        GetDataset = GetDataArray(dataPath)
        errCount = 0
        GoTo CleanExit
    End If
    
    ' --- Case 2: need fresh data ---
    ufLoading.UpdateText "Updating [" & GetParamValue(requestInfo, "DatasetName") & "]"
    
    If Dir(dataPath) <> "" Then
        Kill dataPath
    End If
    
    ' Send Request
    SendRequest requestInfo
    doubleRefresh = True
    
    ' Waiting for data...
    If disableWaitTime Then
        GetDataset = "(waiting for data)"
        Exit Function
    End If
    
    If Not WaitForFileReady(dataPath, MAX_WAIT_SEC) Then
        GetDataset = "request time out"
        GoTo CleanExit
    End If
    
    ' t2 = Timer
    ' Debug.Print "Time - End  : " & TimeMS()
    ' Debug.Print "Time - Spent: " & Format(t2 - t1, "0.000")
    
    If Dir(dataPath) <> "" Then
        GetDataset = GetDataArray(dataPath)
    Else
        Debug.Print "[error] - data path not found"
        GetDataset = "data path not found"
        GoTo CleanExit
    End If
    
    errCount = 0

CleanExit:
    Unload ufLoading
    ufLoading.Reset
    Exit Function
    
ErrHandler:
    Debug.Print "GetDataset error: "; Err.Number; Err.Description
    Resume CleanExit
    
End Function

Public Sub LoadConfig()
    Dim configDir As String
    Dim configPath As String
    Dim line As String, parts As Variant
    Dim fileVersion As String
    Dim f As Integer

    configDir = Environ$("USERPROFILE") & "\ADAS"
    configPath = configDir & "\config.txt"

    ' Ensure config dir
    If Dir(configDir, vbDirectory) = "" Then
        MkDir configDir
    End If

    ' -------------------------
    ' Check existing config version
    ' -------------------------
    If Dir(configPath) <> "" Then
        f = FreeFile
        Open configPath For Input As #f

        Do While Not EOF(f)
            Line Input #f, line
            line = Trim$(line)

            If InStr(line, "=") > 0 Then
                parts = Split(line, "=")
                If LCase$(Trim$(parts(0))) = "version" Then
                    fileVersion = Trim$(parts(1))
                    Exit Do
                End If
            End If
        Loop

        Close #f

        ' Version mismatch ? delete config
        If fileVersion <> ADAS_VERSION Then
            Kill configPath
        End If
    End If

    ' -------------------------
    ' Create config if missing
    ' -------------------------
    If Dir(configPath) = "" Then
        f = FreeFile
        Open configPath For Output As #f
        Print #f, "version = " & ADAS_VERSION
        Print #f, "removeData = False"
        Print #f, "disable_ufLoading = False"
        Print #f, "teamProfile = Default"
        Print #f, "debugMode = False"
        Print #f, "disableProgressBar = False"
        Close #f
    End If

    ' -------------------------
    ' Load config values
    ' -------------------------
    f = FreeFile
    Open configPath For Input As #f

    Do While Not EOF(f)
        Line Input #f, line
        line = Trim$(line)

        If InStr(line, "=") > 0 Then
            parts = Split(line, "=")

            Select Case LCase$(Trim$(parts(0)))
                Case "version"
                    ' ignore, already handled

                Case "removedata"
                    removeData = CBool(Trim$(parts(1)))

                Case "disable_ufLoading", "disable_ufLoading"
                    disable_ufLoading = CBool(Trim$(parts(1)))

                Case "teamprofile"
                    teamProfile = Trim$(parts(1))
                    
                Case "debugMode"
                    debugMode = CBool(Trim$(parts(1)))
                    
                Case "disableProgressBar"
                    disableProgressBar = CBool(Trim$(parts(1)))
                    
            End Select
        End If
    Loop

    Close #f
   
End Sub

Public Sub UpdateConfigValue(ByVal keyName As String, ByVal newValue As String)
    Dim configDir As String, configPath As String
    Dim lines() As String, temp As String
    Dim f As Integer, i As Long

    configDir = Environ$("USERPROFILE") & "\ADAS"
    configPath = configDir & "\config.txt"

    ' Read all lines
    f = FreeFile()
    Open configPath For Input As #f
    lines = Split(Input$(LOF(f), f), vbCrLf)
    Close #f

    ' Modify the specific key
    For i = LBound(lines) To UBound(lines)
        temp = Trim(lines(i))
        If InStr(temp, "=") > 0 Then
            If LCase$(Trim$(Split(temp, "=")(0))) = LCase$(keyName) Then
                lines(i) = keyName & " = " & newValue
            End If
        End If
    Next i

    ' Rewrite file
    f = FreeFile()
    Open configPath For Output As #f
    For i = LBound(lines) To UBound(lines)
        Print #f, lines(i)
    Next i
    Close #f
End Sub

Public Function SetDataPath(inputString As String) As String
    Dim s As String, proj As String
    Dim lines() As String, parts() As String
    Dim i As Long
    Dim key As String, val As String
    Dim fullName As String
    Dim basePath As String
    
    ' Normalize delimiters: allow either "#" or newlines between pairs
    s = inputString
    s = Replace(s, vbCrLf, "#")
    s = Replace(s, vbCr, "#")
    s = Replace(s, vbLf, "#")
    lines = Split(s, "#")
    
    ' Build the @-joined Value list, excluding ProjectName (captured separately)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            If InStr(1, lines(i), "=", vbTextCompare) > 0 Then
                ' Get key and Value (trim everything; be tolerant of spaces around "=")
                parts = Split(lines(i), "=")
                key = Trim$(parts(0))
                val = Trim$(Mid$(lines(i), InStr(1, lines(i), "=", vbTextCompare) + 1))
                
                If LCase$(key) = "projectname" Then
                    proj = val
                Else
                    If Len(fullName) > 0 Then fullName = fullName & "@"
                    fullName = fullName & val
                End If
            End If
        End If
    Next i
    
    ' Sanitize only the concatenated name (slashes become carets, as before)
    fullName = Replace(fullName, "\", "^")
    fullName = Replace(fullName, "/", "^")
    fullName = Replace(fullName, "*", "$star$")
    
    basePath = "E:\ADAS\data\"
    
    ' If ProjectName exists, use it as a subfolder and remove it from fullName (already excluded)
    If Len(proj) > 0 Then
        ' Replace Windows-invalid filename chars in folder name
        Dim invalidChars As Variant, ch As Variant
        invalidChars = Array(":", "*", "?", """", "<", ">", "|")
        For Each ch In invalidChars
            proj = Replace(proj, CStr(ch), "_")
        Next ch
        SetDataPath = basePath & proj & "\" & fullName & ".csv"
    Else
        SetDataPath = basePath & fullName & ".csv"
    End If

End Function

Public Function SetDefaultProject(ByVal ProjectName As String)
    If ProjectName = "Default" Then
        SetDefaultProject = ActiveWorkbook.Sheets("ResQ Settings").Range("B7").Value
    Else
        SetDefaultProject = ProjectName
    End If
End Function

Public Sub SendRequest(requestInfo As String)
    Dim lines() As String
    Dim aFile As Integer
    Dim currentTime As String
    Dim tempPath As String, finalPath As String
    Dim i As Long

    If disableRequest Then Exit Sub

    lines = Split(requestInfo, "#")

    currentTime = Format(Now, "yyyy-mm-dd_hh-mm-ss") & Format(Timer - Int(Timer), ".000")
    tempPath = "E:\ADAS\requests\request-" & currentTime & ".tmp"
    finalPath = "E:\ADAS\requests\request-" & currentTime & ".txt"

    aFile = FreeFile
    Open tempPath For Output As #aFile
        For i = LBound(lines) To UBound(lines)
            Print #aFile, lines(i)
        Next
        Print #aFile, "UserName = " & Environ$("USERNAME")
    Close #aFile

    ' --- overwrite protection ---
    If Dir(finalPath, vbNormal) <> "" Then
        On Error Resume Next
        Kill finalPath
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Kill tempPath     ' cleanup temp
            Exit Sub          ' abort safely
        End If
        On Error GoTo 0
    End If

    ' --- atomic publish ---
    Name tempPath As finalPath
End Sub

Public Function GetDataArray(dataPath As String)
' *----------------------------------------------*
' | Get the data array from an external csv file |
' *----------------------------------------------*
    Dim outputArray() As Variant
    Dim lines() As String
    Dim aFile As Integer
    Dim dateTimeString As String
    Dim data() As String
    Dim fileContent As String
    Dim i As Long, j As Long
    
    aFile = FreeFile
    Open dataPath For Input As #aFile
    fileContent = Input$(LOF(aFile), #aFile)
    Close #aFile

    lines = Split(fileContent, vbCrLf)
    ReDim outputArray(LBound(lines) To UBound(lines) - 1, 0)
    
    For i = LBound(lines) To UBound(lines) - 1
        data = Split(lines(i), ",")
        If UBound(data) > UBound(outputArray, 2) Then
            ReDim Preserve outputArray(LBound(lines) To UBound(lines) - 1, LBound(data) To UBound(data))
        End If
        For j = LBound(data) To UBound(data)
     
            dateTimeString = data(j)
            If InStr(dateTimeString, "+") > 0 Then
                dateTimeString = Left(dateTimeString, InStr(dateTimeString, "+") - 1)
            End If
            
            If IsNumeric(data(j)) Then
                outputArray(i, j) = CDbl(data(j))
            ElseIf IsDate(dateTimeString) Then
                outputArray(i, j) = CDbl(CDate(dateTimeString))
            Else
                outputArray(i, j) = data(j)
            End If
        Next j
    Next i
    
    GetDataArray = outputArray
End Function





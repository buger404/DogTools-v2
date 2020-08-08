Attribute VB_Name = "Core"
Public MusicList As GMusicList
Public user As String
Public Moni As New Monitor
Public Const DataPath As String = "D:\Program Files\DogTools"
Public Type BreakInfo
    ClassName As String
    Title As String
    ImName As String
    Data(4) As String
End Type
Public Type BreakList
    List() As BreakInfo
End Type
Public Breaks As BreakList
Public BreakTime As Long
Public BRI As String, BRT As String, BRC As String
Sub Log(ByVal func As String, ByVal Text As String)
    On Error Resume Next
    
    Open DataPath & "\Logs\" & func & "\" & year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日   " & Hour(Now) & "时.txt" For Append As #1
    Print #1, Now & "    " & Text
    Close #1
    
    If Err.Number <> 0 Then Err.Clear
End Sub
Sub CreateFolder(ByVal path As String)
    Dim temp() As String, NowPath As String, FSO As Object
    If Right(path, 1) <> "\" Then path = path & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    temp = Split(path, "\")
    For I = 0 To UBound(temp) - 1
        If I <> UBound(temp) - 1 Then
            If FSO.FolderExists(NowPath & temp(I)) = False Then Exit Sub
        End If
        NowPath = NowPath & temp(I) & "\"
        If Dir(NowPath, vbDirectory) = "" Then MkDir NowPath
    Next
End Sub
Public Function GetProcessPath(Hwnd As Long) As String
    On Error GoTo z
    
recheck:
    
    Dim PID As Long, Class As String * 255
    Dim cbNeeded As Long, szBuf(1 To 250) As Long, Ret As Long, szPathName As String, nSize As Long, hProcess As Long
    
    Class = "": PID = 0
    
    GetWindowThreadProcessId Hwnd, PID
    GetClassNameA Hwnd, Class, 255
    
    If UnSpace(Class) = "ApplicationFrameWindow" And Hwnd <> 0 Then 'UWP
        Hwnd = uwpFind(Hwnd)
        GoTo recheck
    End If
    
    hProcess = OpenProcess(&H400 Or &H10, 0, PID)
    If hProcess <> 0 Then
        szPathName = Space(260): nSize = 500
        Ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
        GetProcessPath = Left(szPathName, Ret)
    End If
    
    Ret = CloseHandle(hProcess)
    If GetProcessPath = "" Then
        GetProcessPath = "System"
    End If
    
    Exit Function
z:
End Function

Public Function UnSpace(ByVal Str As String) As String
    If InStr(Str, Chr(0)) <> 0 Then
        UnSpace = Left(Str, InStr(Str, Chr(0)) - 1)
    Else
        UnSpace = Str
    End If
End Function
Sub Main()
    If App.LogMode = 0 Then GoTo SkipRoot '调试模式下禁用提权
    'GoTo SkipRoot
    
    If Trim(Command$) = "-root" Then
        'ShellExecuteA 0, "open", """" & App.path & "\ps\psexec64.exe""", "-i -d -s " & """" & App.path & "\" & App.EXEName & ".exe"" ""-rootok""", "", SW_SHOW      '提权到Admin
        'Log "Tools", "取得系统权限"
        'End
    ElseIf Command$ = "" Then
        ShellExecuteA 0, "runas", """" & App.path & "\" & App.EXEName & ".exe""", "-root", "", SW_SHOW      '提权到Admin
        Log "Tools", "取得管理员权限"
        End
    End If
    
SkipRoot:
    
    If Dir(DataPath & "\", vbDirectory) = "" Then CreateFolder DataPath & "\"
    If Dir(DataPath & "\Logs\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\"
    If Dir(DataPath & "\Logs\Breaker\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Breaker\"
    If Dir(DataPath & "\Logs\Monitor\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Monitor\"
    If Dir(DataPath & "\Logs\Tools\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Tools\"
    If Dir(DataPath & "\Monitor\", vbDirectory) = "" Then CreateFolder DataPath & "\Monitor\"
    If Dir(DataPath & "\Logs\Keyboard\", vbDirectory) = "" Then CreateFolder DataPath & "\Logs\Keyboard\"
    
    AppWindow.Show
    
    Log "Tools", "成功启动"
End Sub

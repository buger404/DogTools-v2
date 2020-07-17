VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Private Type SignFile
        Owner As String
        KeyString As String
        SignDots() As Single
        SignKey As String
        GenerateTime As Date
        Permission As String
    End Type
    Dim Page As GPage
    Dim Scan As String, ScanTime As Long, Working As Boolean
    Dim Faces() As String, Score As Long
    Dim RefreshTime As Long, LastDC As Long, PassTime As Long
    Dim Key As SignFile
'======================================
Public Function VerifyUSB(Dr As String) As Boolean
    Dim Name As String, size As Long, Serial As Long, MaxComp As Long, FileFlags As Long, filename As String
    Name = Space(255): filename = Space(255): size = 255

    Dim DriveS As String
    DriveS = Dr
    GetVolumeInformationA DriveS, Name, size, Serial, MaxComp, FileFlags, filename, size

    Open DriveS & "dogtools-v2-key.dt2k" For Binary As #1
    Get #1, , Key
    Close #1
    
    Dim MD5 As New MD5, con As Boolean

    con = Key.KeyString = MD5.Md5_String_Calc("copied with " & Key.Owner & " by " & BMEA(size, Key.Permission) & "&&" & BMEA(Serial, Key.Permission) & "&&" & BMEA(MaxComp, Key.Permission) & "&&" & BMEA(FileFlags, Key.Permission) & "&&" & BMEA(filename, Key.Permission))
    If Not con Then Exit Function
    
    Dim temp As String, TX As Long
    For I = 1 To UBound(Key.SignDots)
        TX = Sqr(Key.SignDots(I) / 500)
        temp = temp & BMEA(Key.SignDots(I), Key.KeyString)
        If Len(temp) > 64 Then temp = MD5.Md5_String_Calc(temp)
        ECore.Display
    Next
    
    con = Key.SignKey = MD5.Md5_String_Calc(Key.Owner & " signed " & temp & " (" & UBound(Key.SignDots) & ")")
    If Not con Then Exit Function
    
    VerifyUSB = True
End Function
Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制代码请放在此处）
    
    Page.Clear argb(255, 255, 255, 255) '清空画布

    Dim pro As Single
    If PassTime <> 0 Then pro = (GetTickCount - PassTime) / 1000
    If pro > 1 Then pro = 1
    pro = Cubic(pro, 0, 1, 1, 1)
    
    If GetTickCount - ScanTime >= 1000 And Working = False And PassTime = 0 Then
        Working = True
        AppWindow.Drive.Refresh
        For I = 0 To AppWindow.Drive.ListCount - 1
            If VerifyUSB(Split(AppWindow.Drive.List(I), ":")(0) & ":\") Then
                Scan = Key.Owner
                PassTime = GetTickCount
                Exit For
            End If
        Next
        ScanTime = GetTickCount
        Working = False
    End If
    
    If Scan <> "" Then
        user = Scan
        Page.Writes Scan, GW / 2 + 1 - pro * 100, GH / 2 - 17, 24, argb(255, 0, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
        Page.Writes "Welcome", GW / 2 + 1 + pro * 100, GH / 2 - 17, 24, argb(pro * 255, 0, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
        Page.Paint 0, GW / 2 + 1 + pro * 50, GH / 2 + 25, pro * 105, 7, argb(pro * 255, 0, 0, 0)
    End If
    
    If pro = 0 Then Page.DrawImage "USB.png", GW / 2, GH / 2 - 60 + pro * GH, alpha:=1 - pro, pos:=posOnCenter
    If pro = 0 Then Page.ShowLoading GW / 2 - 30 - pro * 100, GH / 2 + 160, 60, 60, 8, argb(255, 0, 0, 0), argb(160, 0, 0, 0), argb(80, 0, 0, 0)

    If GetTickCount - PassTime >= 2000 And PassTime <> 0 Then
        Log "Tools", user & "登陆了工具。"
        Scan = "": PassTime = 0: Working = False
        ECore.NewTransform transFadeIn, 1000, "MainPage"
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    ScanTime = GetTickCount
    PassTime = 0
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage  '创建页面绘制器
    Page.Create Me  '将绘制器与控制器绑定
    Page.Res.NewImages App.path & "\assets"  '导入资源
    ECore.Add Page, "AppPage"  '创建页面
    LastDC = CreateCDC(GW, GH)
    Dim Face As String
    Face = Dir(App.path & "\source\")
    ReDim Faces(0)
    Do While Face <> ""
        ReDim Preserve Faces(UBound(Faces) + 1)
        Faces(UBound(Faces)) = Face
        Face = Dir()
    Loop
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

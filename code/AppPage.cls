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
    Dim Page As GPage
    Dim Scan As String, ScanTime As Long, Working As Boolean
    Dim Faces() As String, Score As Long
    Dim RefreshTime As Long, LastDC As Long, PassTime As Long
'======================================

Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制代码请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布

    Page.Writes "摄像头" & NowWebcam & "未连接", GW / 2, GH / 2 - 17, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold

    If GetTickCount - RefreshTime <= 100 Or Scan <> "" Then
        PaintDC LastDC, Page.CDC: GoTo skipCatch
    End If
    
    RefreshTime = GetTickCount
    
    Call CatchWebCam(Page, GetTickCount - ScanTime >= 1000 And Not Working)
    
    If GetTickCount - ScanTime >= 1000 And Not Working Then
        ScanTime = GetTickCount: Working = True
        Dim Code As Long, PID As Long, hid As Long
        Scan = ""
        For I = 1 To UBound(Faces)
            PID = Shell(App.path & "\engine\facecheck.exe " & """" & App.path & "\source\" & Faces(I) & """" & " " & """" & App.path & "\tar.bmp" & """", vbHide)
            hid = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
            Do
                GetExitCodeProcess hid, Code
                ECore.Display: DoEvents
            Loop While Code = STATUS_PENDING
            CloseHandle hid
            If Code > 78 Then
                Score = Code
                Scan = Split(Faces(I), ".")(0)
                ECore.NewTransform transFadeIn, 1000
                PassTime = GetTickCount
                Exit Sub
            End If
        Next
        Working = False
        Exit Sub
    End If
    
    PaintDC Page.CDC, LastDC
    
skipCatch:
    
    Dim pro As Single
    If PassTime <> 0 Then pro = (GetTickCount - PassTime) / 1000
    If pro > 1 Then pro = 1
    pro = Cubic(pro, 0, 1, 1, 1)
    
    If Scan <> "" Then
        user = Scan
        Page.Paint 0, 0, 0, GW, GH, argb(210, 0, 0, 0)
        Page.Writes Scan, GW / 2 + 1 - pro * 100, GH / 2 - 17, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        Page.Writes "Welcome", GW / 2 + 1 + pro * 100, GH / 2 - 17, 24, argb(pro * 255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        Page.Writes Int(pro * Score) & "%", GW / 2, GH - 80, 24, argb(pro * 200, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        Page.Paint 0, GW / 2 + 1 + pro * 50, GH / 2 + 25, pro * 105, 7, argb(pro * 255, 255, 255, 255)
    Else
        Page.Paint 0, 0, 0, GW, GH, argb(60, 0, 0, 0)
    End If
    
    If Working Then Page.ShowLoading GW / 2 - 50 - pro * 100, GH / 2 - 50, 100, 100, 10, argb(255, 255, 255, 255), argb(160, 255, 255, 255), argb(80, 255, 255, 255)

    Page.Writes "switch the camera", GW / 2, GH - 80 + 200 * pro, 24, argb(200 - 200 * pro, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    If CheckMouse(GW / 2 - 200, GH - 120, 400, 80) >= mMouseDown And pro = 0 Then
        SwitchWebCam
    End If
    
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

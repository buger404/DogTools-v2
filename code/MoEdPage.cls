VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoEdPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim PI As Integer
'======================================

Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制代码请放在此处）
    
    Page.Clear '清空画布

    Page.Writes "breaker", GW / 2, 100, 32, argb(255, 253, 78, 38), align:=StringAlignmentCenter
    Page.Writes "avoid something that is not expected to be opened", GW / 2, 160, 20, argb(120, 0, 0, 0), align:=StringAlignmentCenter

    For I = PI * 5 + 1 To PI * 5 + 5
        If I > UBound(Breaks.List) Then Exit For
        If I < 0 Then
            PI = 0: Exit For
        End If
        With Breaks.List(I)
            Page.Writes .Title & " | " & .ClassName & " | " & .ImName, 50, 200 + (I - PI * 5) * 45, 20, argb(200, 0, 0, 0)
        End With
        Page.Writes "Delete", GW - 50, 200 + (I - PI * 5) * 45, 20, argb(230, 253, 78, 38), align:=StringAlignmentFar
        If CheckMouse(GW - 50 - 100, 200 + (I - PI * 5) * 45, 100, 40) = mMouseUp Then
            If ECore.SimpleMsg("Delete: " & Breaks.List(I).Title & " | " & Breaks.List(I).ClassName & " | " & Breaks.List(I).ImName, "Breaker", StrArray("Delete", "Cancel"), Radius:=20) = 0 Then
                Mouse.button = 0: Mouse.State = 0
                For S = I To UBound(Breaks.List) - 1
                    Breaks.List(S) = Breaks.List(S + 1)
                Next
                ReDim Preserve Breaks.List(UBound(Breaks.List) - 1)
                Open DataPath & "\break.db" For Binary As #1
                Put #1, , Breaks
                Close #1
                TrayBalloon AppWindow, "A rule deleted just now.", "BREAKER", NIIF_ICON_MASK
            End If
        End If
    Next
    
    Dim min As Long, max As Long
    min = 0
    max = Int(UBound(Breaks.List) / 6)
    If PI > max Then PI = max
    
    Page.Writes "< Prev", 50, 485, 20, argb(180, 255, 83, 57)
    If CheckMouse(50, 480, 75, 35) = mMouseUp And PI > min Then
        Mouse.button = 0: Mouse.State = 0
        PI = PI - 1
    End If
    Page.Writes PI + 1 & "/" & max + 1, GW / 2, 485, 20, argb(180, 255, 83, 57), align:=StringAlignmentCenter
    Page.Writes "Next >", GW - 50, 485, 20, argb(180, 255, 83, 57), align:=StringAlignmentFar
    If CheckMouse(GW - 50 - 75, 480, 75, 35) = mMouseUp And PI < max Then
        Mouse.button = 0: Mouse.State = 0
        PI = PI + 1
    End If
    
    Page.Paint 0, 0, GH - 149, GW, 150, argb(255, 242, 242, 242)
    If CheckMouse2 = mMouseUp Then
        If ECore.SimpleMsg("Are you wanna to avoid a new window?", "Breaker", StrArray("Yes", "No"), Radius:=20) = 0 Then
            BreakTime = GetTickCount
        End If
    End If
    Page.Writes IIf(BreakTime <> 0, Int(6 - (GetTickCount - BreakTime) / 1000) & "s left ...", "TARGET"), GW / 2, GH - 100, 24, argb(255, 0, 0, 0), align:=StringAlignmentCenter

    If CheckMouse(0, 0, GW, GH) = mMouseUp And Mouse.button = 2 Then
        Mouse.button = 0: Mouse.State = 0
        ECore.NewTransform transDarkReturn, 500, "MainPage"
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
    ECore.Add Page, "MoEdPage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

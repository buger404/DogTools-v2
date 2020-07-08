VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Dim Page As GPage
'======================================

Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制代码请放在此处）
    
    Page.Clear argb(255, 255, 255, 255) '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim pro As Single
    For I = 1 To 4
        pro = Cubic(((GetTickCount + I * 500) Mod 3000) / 3000, 0, 1, 1, 1)
        Page.Paint 1, GW / 2, GH / 2 - 50, pro * 500, pro * 500, argb(150 - 150 * pro, 195, 157, 147), pos:=posOnCenter
    Next
    Page.DrawImage "Icon.png", GW / 2, GH / 2 - 50, alpha:=1, pos:=posOnCenter
    
    Page.Writes "WELCOME " & user, GW / 2 + 1, GH / 2 + 80, 28, argb(255, 0, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
    
    If ESave.GetData("setuptime") = "" Then ESave.PutData "setuptime", Now
    
    Dim otime As Date
    otime = Now - CDate(ESave.GetData("setuptime"))
    
    Page.Writes "Dog Tools -v2 has get along with you " & year(otime) Mod 1899 & "y" & Month(otime) Mod 12 & "m" & Day(otime) Mod 30 & "d" & Hour(otime) & "h" & Minute(otime) & "min" & Second(otime) & "s", GW / 2 + 1, GH / 2 + 240, 20, argb(180, 0, 0, 0), align:=StringAlignmentCenter
    
    If CheckMouse(0, 0, GW, GH) = mMouseDown And Mouse.button = 1 Then
        Mouse.button = 0: Mouse.State = 0
        ECore.NewTransform transDarkTo, 500, "MoEdPage"
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
    ECore.Add Page, "MainPage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

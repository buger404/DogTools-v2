VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RulePage"
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
    
    Page.Clear '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    Page.Paint 0, 0, 120, GW + 1, GH - 120 + 1, argb(255, 242, 242, 242)
    
    Page.Writes "Rule arrangement", 40, 40, 28, argb(255, 255, 83, 57)
    
    Page.Writes "Title", 40, 150, 20, argb(255, 16, 16, 16)
    Page.ShowEdit BRT, 0, 150, 150, 600, 30, argb(255, 64, 64, 64), argb(255, 232, 232, 232), argb(255, 232, 232, 235), argb(255, 255, 83, 57), size:=20

    Page.Writes "Class", 40, 200, 20, argb(255, 16, 16, 16)
    Page.ShowEdit BRC, 0, 150, 200, 600, 30, argb(255, 64, 64, 64), argb(255, 232, 232, 232), argb(255, 232, 232, 235), argb(255, 255, 83, 57), size:=20

    Page.Writes "Image", 40, 250, 20, argb(255, 16, 16, 16)
    Page.ShowEdit BRI, 0, 150, 250, 600, 30, argb(255, 64, 64, 64), argb(255, 232, 232, 232), argb(255, 232, 232, 235), argb(255, 255, 83, 57), size:=20

    If Page.ShowColorButton(2, GW / 2 - 70, GH - 100, 140, 40, "confirm", argb(255, 255, 255, 255), argb(255, 253, 78, 38), argb(255, 253, 78, 38), 999, size:=20) = mMouseUp Then
        Moni.ConfirmRule
        ECore.NewTransform , , "MoEdPage"
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
    ECore.Add Page, "RulePage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

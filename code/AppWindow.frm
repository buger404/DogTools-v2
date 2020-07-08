VERSION 5.00
Begin VB.Form AppWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dog Tools -v2 Made by CZY"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   Icon            =   "AppWindow.frx":0000
   LinkTopic       =   "AppWindow"
   MaxButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer MoniTimer 
      Interval        =   100
      Left            =   9048
      Top             =   780
   End
   Begin VB.PictureBox WebCam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   324
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   该类模块是由Emerald创建的 界面容器窗口（应用窗口） 模板
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面控制器类模块声明语句
    Dim AppPage As AppPage
    Dim MainPage As MainPage
    Dim MoEdPage As MoEdPage
'==================================================

Private Sub Form_Load()
    CreateFolder "C:\DogTools\Logs\Monitor\"
    CreateFolder "C:\DogTools\Logs\Keyboard\"
    CreateFolder "C:\DogTools\Logs\Tools\"
    CreateFolder "C:\DogTools\Logs\Breaker\"
    If App.LogMode <> 0 Then StartKeyboard
    
    StartEmerald Me.Hwnd, 1200, 700  '初始化Emerald（在此处可以修改窗口大小）
    WebCam.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Call StartWebCam(WebCam)
    
    Set EF = New GFont
    EF.AddFont App.path & "\font.ttf"
    EF.MakeFont "Selawik"  '创建字体
   
    Set EC = New GMan   '创建页面管理器
    
    '创建存档（可选），存档key的问题请查看Emerald的wiki
    Set ESave = New GSaving
    ESave.Create "DogTools-v2.Keys", "213C5BC91FC31B4446D590D1CF68566"
    DisableLOGO = 1: HideLOGO = 1
    
    '创建音乐列表（可选）
    'Set MusicList = New GMusicList
    'MusicList.Create App.path & "\music"

    '开始显示界面
    Me.Show
    DrawTimer.Enabled = True
    
    '在此处实例化你的页面控制器
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set AppPage = New AppPage
        Set MainPage = New MainPage
        Set MoEdPage = New MoEdPage
    '=============================================

    '设置活动页面（在此处设置则为你的启动页面）
    EC.ActivePage = "AppPage"
    
    TrayAddIcon Me, "Dog Tools -v2", NIIF_NONE
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    If ECore.ActivePage = "AppPage" Then
        Log "Tools", "密码校验失败"
        TrayBalloon AppWindow, "I won't quit unless you login your account.", "STILL HERE", NIIF_ERROR
        Me.Hide
    Else
        If MsgBox("Lock your account anyway?", 48 + vbYesNo) = vbYes Then
            Log "Tools", user & "锁定了工具。"
            TrayBalloon AppWindow, "You log-out your account successfully.", "ACCOUNT LOG-OUT", NIIF_INFO
            user = "": Me.Hide
        End If
    End If
    Exit Sub
    If App.LogMode <> 0 Then EndKeyboard
    TrayRemoveIcon
    Call StopWebCam
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
    End
End Sub

Private Sub DrawTimer_Timer()
    EC.Display
End Sub

'============================================================
' 事件映射
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    '发送鼠标信息
    UpdateMouse X, Y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    '发送鼠标信息
    If Me.Visible = False Then
        If button = 2 Then
            ECore.ActivePage = IIf(user <> "", "MainPage", "AppPage")
            Me.WindowState = 0
            Me.Show
        End If
        Exit Sub
    End If

    If Mouse.State = 0 Then
        UpdateMouse X, Y, 0, button
    Else
        Mouse.X = X: Mouse.Y = Y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    '发送鼠标信息
    UpdateMouse X, Y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub MoniTimer_Timer()
    Moni.Update
End Sub

'============================================================
Private Sub WebCam_Click()

End Sub

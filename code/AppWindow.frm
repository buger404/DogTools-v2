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
   StartUpPosition =   2  '��Ļ����
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
'   ����ģ������Emerald������ �����������ڣ�Ӧ�ô��ڣ� ģ��
'==================================================
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ���������ģ���������
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
    
    StartEmerald Me.Hwnd, 1200, 700  '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�С��
    WebCam.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Call StartWebCam(WebCam)
    
    Set EF = New GFont
    EF.AddFont App.path & "\font.ttf"
    EF.MakeFont "Selawik"  '��������
   
    Set EC = New GMan   '����ҳ�������
    
    '�����浵����ѡ�����浵key��������鿴Emerald��wiki
    Set ESave = New GSaving
    ESave.Create "DogTools-v2.Keys", "213C5BC91FC31B4446D590D1CF68566"
    DisableLOGO = 1: HideLOGO = 1
    
    '���������б���ѡ��
    'Set MusicList = New GMusicList
    'MusicList.Create App.path & "\music"

    '��ʼ��ʾ����
    Me.Show
    DrawTimer.Enabled = True
    
    '�ڴ˴�ʵ�������ҳ�������
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set AppPage = New AppPage
        Set MainPage = New MainPage
        Set MoEdPage = New MoEdPage
    '=============================================

    '���ûҳ�棨�ڴ˴�������Ϊ�������ҳ�棩
    EC.ActivePage = "AppPage"
    
    TrayAddIcon Me, "Dog Tools -v2", NIIF_NONE
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    If ECore.ActivePage = "AppPage" Then
        Log "Tools", "����У��ʧ��"
        TrayBalloon AppWindow, "I won't quit unless you login your account.", "STILL HERE", NIIF_ERROR
        Me.Hide
    Else
        If MsgBox("Lock your account anyway?", 48 + vbYesNo) = vbYes Then
            Log "Tools", user & "�����˹��ߡ�"
            TrayBalloon AppWindow, "You log-out your account successfully.", "ACCOUNT LOG-OUT", NIIF_INFO
            user = "": Me.Hide
        End If
    End If
    Exit Sub
    If App.LogMode <> 0 Then EndKeyboard
    TrayRemoveIcon
    Call StopWebCam
    '��ֹ����
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    EndEmerald
    End
End Sub

Private Sub DrawTimer_Timer()
    EC.Display
End Sub

'============================================================
' �¼�ӳ��
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
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
    '���������Ϣ
    UpdateMouse X, Y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub MoniTimer_Timer()
    Moni.Update
End Sub

'============================================================
Private Sub WebCam_Click()

End Sub

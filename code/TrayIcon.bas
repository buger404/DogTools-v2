Attribute VB_Name = "TrayIcon"
'使用高分辨率图标所用的API
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const IMAGE_ICON = 1
'系统托盘
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5
Private Const WM_USER As Long = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
Private Const NOTIFYICON_VERSION = 3
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_COMMAND As Long = &H111
Private Const WM_CLOSE As Long = &H10
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

Public Enum bFlag
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_GUID = &H5
    NIIF_ICON_MASK = &HF
    NIIF_NOSOUND = &H10 '关闭提示音标志
End Enum

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutAndVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

'鼠标事件
Public Enum TrayRetunEventEnum
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
    BalloonClick = (WM_USER + 5)
End Enum

Public ni As NOTIFYICONDATA
Public Sub TrayAddIcon(ByVal MyForm As Form, ByVal ToolTip As String, Optional ByVal bFlag As bFlag)

    With ni
        .cbSize = Len(ni)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = MyForm.Icon
        .szTip = ToolTip & vbNullChar
    End With
    
    Call Shell_NotifyIcon(NIM_ADD, ni)

End Sub

Public Sub TrayRemoveIcon()

    Shell_NotifyIcon NIM_DELETE, ni
    
End Sub

Public Sub TrayBalloon(ByVal MyForm As Form, ByVal sBaloonText As String, sBallonTitle As String, Optional ByVal bFlag As bFlag)

    With ni
        .cbSize = Len(ni)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_INFO
        .dwInfoFlags = bFlag
        .szInfoTitle = sBallonTitle & vbNullChar
        .szInfo = sBaloonText & vbNullChar
    End With

    Shell_NotifyIcon NIM_MODIFY, ni

End Sub

Public Sub TrayTip(ByVal MyForm As Form, ByVal sTipText As String)

    With ni
        .cbSize = Len(ni)
        .hwnd = MyForm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .szTip = sTipText & vbNullChar
    End With

    Shell_NotifyIcon NIM_MODIFY, ni

End Sub

Attribute VB_Name = "WebCamCore"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054
Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private CapHwnd As Long, CapBox As PictureBox
Public NowWebcam As Integer
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOZORDER = &H4
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SM_CYCAPTION = 4
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_STYLE = (-16)
'为窗体设置值
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function lStrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Declare Function lStrCpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As Any, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByVal hpvDest As Long, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub hmemcpy Lib "kernel32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'这个函数为窗口指定个个新位置和状态。它也可改变窗口在内部窗口列表中的位置
Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'关闭窗体及子窗体
Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'在结构中为指定的窗口设置信息
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String) As Long
Public lwndC As Long '窗体句柄
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Function ReleaseCapture Lib "user32" () As Long
'**********************************'保存窗口最前
Public Const WM_USER = &H400 '偏移地址
Type POINTAPI
X As Long
y As Long
End Type
'调用一个窗口的窗口函数，将一条消息发给那个窗口。直到消息被处理完毕，该函数才会返回
'hwnd（long）要接收消息的那个窗口的句柄、 wmsg（long）消息的标识符 、wparam（long）具体取决于消息 iparam（ANY）具体取决于消息
Declare Function SendMessageS Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long
Public Const WM_CAP_START = WM_USER '开始址
Public Const WM_CAP_GET_CAPSTREAMPTR = WM_CAP_START + 1 '
Public Const WM_CAP_SET_CALLBACK_ERROR = WM_CAP_START + 2 '在程序设定当发生错误时调用的回调函数
Public Const WM_CAP_SET_CALLBACK_STATUS = WM_CAP_START + 3 '在程序中设定当状态改变时调用的回调函数
Public Const WM_CAP_SET_CALLBACK_YIELD = WM_CAP_START + 4 '在程序中设定当程序让位时调用的回调函数
Public Const WM_CAP_SET_CALLBACK_FRAME = WM_CAP_START + 5 '在程序中设定当预览帧被捕捉时调用的加调函数
Public Const WM_CAP_SET_CALLBACK_VIDEOSTREAM = WM_CAP_START + 6 '在程序中设定当一个新的视频缓冲区可以时调用的回调函数
Public Const WM_CAP_SET_CALLBACK_WAVESTREAM = WM_CAP_START + 7 '在程序中设定当一个新的音频缓冲区可以时调用的回调函数
Public Const WM_CAP_GET_USER_DATA = WM_CAP_START + 8 '把数据关联到一个捕捉窗口，可以获取一个长整型数据
Public Const WM_CAP_SET_USER_DATA = WM_CAP_START + 9 '把数据关联到一个捕捉窗口，'可以设置一个长整型数据
Public Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10 '把捕捉窗口连接到一个捕捉设备
Public Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11 ''用来断开捕捉驱动器和捕捉窗口之间的连接
Public Const WM_CAP_DRIVER_GET_NAME = WM_CAP_START + 12 '连接到'来得到已连接到某个捕捉窗口的捕捉设备驱动程序的名字
Public Const WM_CAP_DRIVER_GET_VERSION = WM_CAP_START + 13 '得到驱动程序的版本号
Public Const WM_CAP_DRIVER_GET_CAPS = WM_CAP_START + 14 '来得到捕捉窗口的硬件的性能。
'捕捉文件和缓存
Public Const WM_CAP_FILE_SET_CAPTURE_FILE = WM_CAP_START + 20 '可以指定另一个路径和文件名。这个消息指定文件名，但不创建文件，也不打开文件或为文件申请空间
Public Const WM_CAP_FILE_GET_CAPTURE_FILE = WM_CAP_START + 21 '来得到当前的捕捉文件
Public Const WM_CAP_FILE_ALLOCATE = WM_CAP_START + 22 '为捕捉文件预分配空间,从而可以减少被漏掉的帧
Public Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23 '将捕捉文件保存为另一个用户指定的文件。这个消息不会改变捕捉文件的名字和内容,
'由于捕捉文件保留它最初的文件名，因此必须指定个新的文件的文件名来保存
Public Const WM_CAP_FILE_SET_INFOCHUNK = WM_CAP_START + 24 '可以把信息块例如文本或者自定义数据插入avi文件。同样用这个消息也可以清除avi文件中的信息块
Public Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25 '把从帧缓存中复制出图像存为设备无关位图书馆（DIB),应用程序也可以使用这两个单帧捕捉消息来编辑帧序列，
'或者创建一个慢速摄影序列
Public Const WM_CAP_EDIT_COPY = WM_CAP_START + 30 '一旦捕捉到图像，把缓存中图像复制到剪贴板中
Public Const WM_CAP_SET_AUDIOFORMAT = WM_CAP_START + 35 '设置音频格式。设置时传入一个WAVEFORMAT、WAVEFORMATEX、或PCMWAVEOFMAT结构的指针
Public Const WM_CAP_GET_AUDIOFORMAT = WM_CAP_START + 36 '来得到音频数据的格式和该格式结构体的大小。默认的捕捉音频格式是mono、8-bit和11kHZ PCM
Public Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41 '对数字化后的视频帧的大小和图像深度，以及被捕捉视频的数据的压缩方式的选择
Public Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42 '视频对话框，用来枚举连接视频源的捕捉卡的信号种类，并且
'控制颜色、对比度、饱和度的改变，如果视频驱动程序支技，可以用这个对话框
Public Const WM_CAP_DLG_VIDEODISPLAY = WM_CAP_START + 43 '视频显示对话框控制视频捕捉过程中视频在显示器上的显示。对捕捉数据无影响，但会影响数了信号表达式
Public Const WM_CAP_GET_VIDEOFORMAT = WM_CAP_START + 44 '给捕捉窗口来得到视频格式的结构和该结构的大小。
Public Const WM_CAP_SET_VIDEOFORMAT = WM_CAP_START + 45 '用来设置视频格式
Public Const WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46 ' 视频压缩对话框
Public Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50 '发送给捕捉窗口来使预览模式有效或者失效
Public Const WM_CAP_SET_OVERLAY = WM_CAP_START + 51 '使窗口处于叠加模式。使叠加模式有效也会自动地使预览模式失效
Public Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52 '发送给捕捉窗口来设置在预览模式下帧的显示频率
Public Const WM_CAP_SET_SCALE = WM_CAP_START + 53 '来使预览模式的缩放有效或者无效
Public Const WM_CAP_GET_STATUS = WM_CAP_START + 54 '得到捕捉窗口的当前状态
Public Const WM_CAP_SET_SCROLL = WM_CAP_START + 55 '如果是在预览模式或者叠加模式，还可以通过本消息发送给窗口，
'在窗口里的用户区域设置视频帧的滚动条的位置
'定义结束时响应信息
Public Const WM_CAP_SET_CALLBACK_CAPCONTROL = WM_CAP_START + 85
Public Const WM_CAP_END = WM_CAP_SET_CALLBACK_CAPCONTROL
'// 导入avicap32.dll连接库下的两个函数
Declare Function capCreateCaptureWindowA Lib "avicap32.dll" ( _
ByVal lpszWindowName As String, _
ByVal dwStyle As Long, _
ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, _
ByVal hWndParent As Long, ByVal nID As Long) As Long
Declare Function capGetDriverDescriptionA Lib "avicap32.dll" ( _
ByVal wDriver As Integer, _
ByVal lpszName As String, _
ByVal cbName As Long, _
ByVal lpszVer As String, _
ByVal cbVer As Long) As Boolean
Function capDriverConnect(ByVal lwnd As Long, ByVal I As Integer) As Boolean
'把捕捉窗口连接到一个捕捉设备
capDriverConnect = SendMessage(lwnd, WM_CAP_DRIVER_CONNECT, I, 0)
End Function
Function capDriverDisconnect(ByVal lwnd As Long) As Boolean
''用来断开捕捉驱动器和捕捉窗口之间的连接
capDriverDisconnect = SendMessage(lwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0)
End Function
Function capDriverGetName(ByVal lwnd As Long, ByVal szName As Long, ByVal wSize As Integer) As Boolean
'获得驱动程序名字
capDriverGetName = SendMessage(lwnd, YOURCONSTANTMESSAGE, wSize, szName)
End Function
Function capDriverGetCaps(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
'来得到捕捉窗口的硬件的性能
capDriverGetCaps = SendMessage(lwnd, WM_CAP_DRIVER_GET_CAPS, wSize, S)
End Function
Function capPreview(ByVal lwnd As Long, ByVal f As Boolean) As Boolean
'发送给捕捉窗口来使预览模式有效或者失效
capPreview = SendMessage(lwnd, WM_CAP_SET_PREVIEW, f, 0)
End Function
Function capPreviewRate(ByVal lwnd As Long, ByVal wMS As Integer) As Boolean
'发送给捕捉窗口来设置在预览模式下帧的显示频率
capPreviewRate = SendMessage(lwnd, WM_CAP_SET_PREVIEWRATE, wMS, 0)
End Function
Function capPreviewScale(ByVal lwnd As Long, ByVal f As Boolean) As Boolean
'来使预览模式的缩放有效或者无效
capPreviewScale = SendMessage(lwnd, WM_CAP_SET_SCALE, f, 0)
End Function
Public Sub StartWebCam(RenderBox As PictureBox)
    Set CapBox = RenderBox
    CapHwnd = capCreateCaptureWindow("Capture Window", 0, 0, 0, CapBox.Width, CapBox.Height, 0, 0)
    SendMessage CapHwnd, CONNECT, 0, 0
    'SendMessage CapHwnd, 1024 + 50, 1, 0
    'SendMessage CapHwnd, 1024 + 52, 1, 0
End Sub
Public Sub SwitchWebCam()
    SendMessage CapHwnd, DISCONNECT, NowWebcam, NowWebcam
    NowWebcam = IIf(NowWebcam = 0, 1, 0)
    SendMessage CapHwnd, CONNECT, NowWebcam, NowWebcam
End Sub
Public Sub StopWebCam()
    SendMessage CapHwnd, DISCONNECT, 0, 0
    DestroyWindow CapHwnd
End Sub
Public Sub CatchWebCam(Page As GPage, Out As Boolean)
    On Error Resume Next
    SendMessage CapHwnd, GET_FRAME, StrPtr(path), 0
    SendMessage CapHwnd, COPY, 0, 0
    CapBox.Picture = Clipboard.GetData
    Clipboard.Clear
    Dim Image As Long
    GdipCreateBitmapFromHBITMAP CapBox.Picture.handle, CapBox.Picture.hpal, Image
    GdipDrawImageRect Page.GG, Image, 0, 0, GW, GH
    GdipDisposeImage Image
    If Out Then
        'GdipSaveImageToFile image, StrPtr(App.path & "\tar.png"), ImageFormatPNG, ImageEncoderPNG
        SavePicture CapBox.Picture, App.path & "\tar.bmp"
    End If
End Sub

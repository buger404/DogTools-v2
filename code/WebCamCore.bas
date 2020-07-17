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
'Ϊ��������ֵ
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function lStrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Declare Function lStrCpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As Any, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByVal hpvDest As Long, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub hmemcpy Lib "kernel32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'�������Ϊ����ָ��������λ�ú�״̬����Ҳ�ɸı䴰�����ڲ������б��е�λ��
Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'�رմ��弰�Ӵ���
Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'�ڽṹ��Ϊָ���Ĵ���������Ϣ
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String) As Long
Public lwndC As Long '������
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Function ReleaseCapture Lib "user32" () As Long
'**********************************'���洰����ǰ
Public Const WM_USER = &H400 'ƫ�Ƶ�ַ
Type POINTAPI
X As Long
y As Long
End Type
'����һ�����ڵĴ��ں�������һ����Ϣ�����Ǹ����ڡ�ֱ����Ϣ��������ϣ��ú����Ż᷵��
'hwnd��long��Ҫ������Ϣ���Ǹ����ڵľ���� wmsg��long����Ϣ�ı�ʶ�� ��wparam��long������ȡ������Ϣ iparam��ANY������ȡ������Ϣ
Declare Function SendMessageS Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long
Public Const WM_CAP_START = WM_USER '��ʼַ
Public Const WM_CAP_GET_CAPSTREAMPTR = WM_CAP_START + 1 '
Public Const WM_CAP_SET_CALLBACK_ERROR = WM_CAP_START + 2 '�ڳ����趨����������ʱ���õĻص�����
Public Const WM_CAP_SET_CALLBACK_STATUS = WM_CAP_START + 3 '�ڳ������趨��״̬�ı�ʱ���õĻص�����
Public Const WM_CAP_SET_CALLBACK_YIELD = WM_CAP_START + 4 '�ڳ������趨��������λʱ���õĻص�����
Public Const WM_CAP_SET_CALLBACK_FRAME = WM_CAP_START + 5 '�ڳ������趨��Ԥ��֡����׽ʱ���õļӵ�����
Public Const WM_CAP_SET_CALLBACK_VIDEOSTREAM = WM_CAP_START + 6 '�ڳ������趨��һ���µ���Ƶ����������ʱ���õĻص�����
Public Const WM_CAP_SET_CALLBACK_WAVESTREAM = WM_CAP_START + 7 '�ڳ������趨��һ���µ���Ƶ����������ʱ���õĻص�����
Public Const WM_CAP_GET_USER_DATA = WM_CAP_START + 8 '�����ݹ�����һ����׽���ڣ����Ի�ȡһ������������
Public Const WM_CAP_SET_USER_DATA = WM_CAP_START + 9 '�����ݹ�����һ����׽���ڣ�'��������һ������������
Public Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10 '�Ѳ�׽�������ӵ�һ����׽�豸
Public Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11 ''�����Ͽ���׽�������Ͳ�׽����֮�������
Public Const WM_CAP_DRIVER_GET_NAME = WM_CAP_START + 12 '���ӵ�'���õ������ӵ�ĳ����׽���ڵĲ�׽�豸�������������
Public Const WM_CAP_DRIVER_GET_VERSION = WM_CAP_START + 13 '�õ���������İ汾��
Public Const WM_CAP_DRIVER_GET_CAPS = WM_CAP_START + 14 '���õ���׽���ڵ�Ӳ�������ܡ�
'��׽�ļ��ͻ���
Public Const WM_CAP_FILE_SET_CAPTURE_FILE = WM_CAP_START + 20 '����ָ����һ��·�����ļ����������Ϣָ���ļ��������������ļ���Ҳ�����ļ���Ϊ�ļ�����ռ�
Public Const WM_CAP_FILE_GET_CAPTURE_FILE = WM_CAP_START + 21 '���õ���ǰ�Ĳ�׽�ļ�
Public Const WM_CAP_FILE_ALLOCATE = WM_CAP_START + 22 'Ϊ��׽�ļ�Ԥ����ռ�,�Ӷ����Լ��ٱ�©����֡
Public Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23 '����׽�ļ�����Ϊ��һ���û�ָ�����ļ��������Ϣ����ı䲶׽�ļ������ֺ�����,
'���ڲ�׽�ļ�������������ļ�������˱���ָ�����µ��ļ����ļ���������
Public Const WM_CAP_FILE_SET_INFOCHUNK = WM_CAP_START + 24 '���԰���Ϣ�������ı������Զ������ݲ���avi�ļ���ͬ���������ϢҲ�������avi�ļ��е���Ϣ��
Public Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25 '�Ѵ�֡�����и��Ƴ�ͼ���Ϊ�豸�޹�λͼ��ݣ�DIB),Ӧ�ó���Ҳ����ʹ����������֡��׽��Ϣ���༭֡���У�
'���ߴ���һ��������Ӱ����
Public Const WM_CAP_EDIT_COPY = WM_CAP_START + 30 'һ����׽��ͼ�񣬰ѻ�����ͼ���Ƶ���������
Public Const WM_CAP_SET_AUDIOFORMAT = WM_CAP_START + 35 '������Ƶ��ʽ������ʱ����һ��WAVEFORMAT��WAVEFORMATEX����PCMWAVEOFMAT�ṹ��ָ��
Public Const WM_CAP_GET_AUDIOFORMAT = WM_CAP_START + 36 '���õ���Ƶ���ݵĸ�ʽ�͸ø�ʽ�ṹ��Ĵ�С��Ĭ�ϵĲ�׽��Ƶ��ʽ��mono��8-bit��11kHZ PCM
Public Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41 '�����ֻ������Ƶ֡�Ĵ�С��ͼ����ȣ��Լ�����׽��Ƶ�����ݵ�ѹ����ʽ��ѡ��
Public Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42 '��Ƶ�Ի�������ö��������ƵԴ�Ĳ�׽�����ź����࣬����
'������ɫ���Աȶȡ����Ͷȵĸı䣬�����Ƶ��������֧��������������Ի���
Public Const WM_CAP_DLG_VIDEODISPLAY = WM_CAP_START + 43 '��Ƶ��ʾ�Ի��������Ƶ��׽��������Ƶ����ʾ���ϵ���ʾ���Բ�׽������Ӱ�죬����Ӱ�������źű��ʽ
Public Const WM_CAP_GET_VIDEOFORMAT = WM_CAP_START + 44 '����׽�������õ���Ƶ��ʽ�Ľṹ�͸ýṹ�Ĵ�С��
Public Const WM_CAP_SET_VIDEOFORMAT = WM_CAP_START + 45 '����������Ƶ��ʽ
Public Const WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46 ' ��Ƶѹ���Ի���
Public Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50 '���͸���׽������ʹԤ��ģʽ��Ч����ʧЧ
Public Const WM_CAP_SET_OVERLAY = WM_CAP_START + 51 'ʹ���ڴ��ڵ���ģʽ��ʹ����ģʽ��ЧҲ���Զ���ʹԤ��ģʽʧЧ
Public Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52 '���͸���׽������������Ԥ��ģʽ��֡����ʾƵ��
Public Const WM_CAP_SET_SCALE = WM_CAP_START + 53 '��ʹԤ��ģʽ��������Ч������Ч
Public Const WM_CAP_GET_STATUS = WM_CAP_START + 54 '�õ���׽���ڵĵ�ǰ״̬
Public Const WM_CAP_SET_SCROLL = WM_CAP_START + 55 '�������Ԥ��ģʽ���ߵ���ģʽ��������ͨ������Ϣ���͸����ڣ�
'�ڴ�������û�����������Ƶ֡�Ĺ�������λ��
'�������ʱ��Ӧ��Ϣ
Public Const WM_CAP_SET_CALLBACK_CAPCONTROL = WM_CAP_START + 85
Public Const WM_CAP_END = WM_CAP_SET_CALLBACK_CAPCONTROL
'// ����avicap32.dll���ӿ��µ���������
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
'�Ѳ�׽�������ӵ�һ����׽�豸
capDriverConnect = SendMessage(lwnd, WM_CAP_DRIVER_CONNECT, I, 0)
End Function
Function capDriverDisconnect(ByVal lwnd As Long) As Boolean
''�����Ͽ���׽�������Ͳ�׽����֮�������
capDriverDisconnect = SendMessage(lwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0)
End Function
Function capDriverGetName(ByVal lwnd As Long, ByVal szName As Long, ByVal wSize As Integer) As Boolean
'���������������
capDriverGetName = SendMessage(lwnd, YOURCONSTANTMESSAGE, wSize, szName)
End Function
Function capDriverGetCaps(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
'���õ���׽���ڵ�Ӳ��������
capDriverGetCaps = SendMessage(lwnd, WM_CAP_DRIVER_GET_CAPS, wSize, S)
End Function
Function capPreview(ByVal lwnd As Long, ByVal f As Boolean) As Boolean
'���͸���׽������ʹԤ��ģʽ��Ч����ʧЧ
capPreview = SendMessage(lwnd, WM_CAP_SET_PREVIEW, f, 0)
End Function
Function capPreviewRate(ByVal lwnd As Long, ByVal wMS As Integer) As Boolean
'���͸���׽������������Ԥ��ģʽ��֡����ʾƵ��
capPreviewRate = SendMessage(lwnd, WM_CAP_SET_PREVIEWRATE, wMS, 0)
End Function
Function capPreviewScale(ByVal lwnd As Long, ByVal f As Boolean) As Boolean
'��ʹԤ��ģʽ��������Ч������Ч
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

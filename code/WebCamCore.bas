Attribute VB_Name = "WebCamCore"
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054
Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private CapHwnd As Long, CapBox As PictureBox
Public NowWebcam As Integer
Public Sub StartWebCam(RenderBox As PictureBox)
    Set CapBox = RenderBox
    CapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, CapBox.Width, CapBox.Height, 0, 0)
    SendMessage CapHwnd, CONNECT, 0, 0
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
    Dim image As Long
    GdipCreateBitmapFromHBITMAP CapBox.Picture.handle, CapBox.Picture.hpal, image
    GdipDrawImageRect Page.GG, image, 0, 0, GW, GH
    GdipDisposeImage image
    If Out Then
        'GdipSaveImageToFile image, StrPtr(App.path & "\tar.png"), ImageFormatPNG, ImageEncoderPNG
        SavePicture CapBox.Picture, App.path & "\tar.bmp"
    End If
End Sub

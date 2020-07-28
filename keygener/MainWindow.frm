VERSION 5.00
Begin VB.Form MainWindow 
   BackColor       =   &H00FBFBFB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DogTools Key Generate"
   ClientHeight    =   5916
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   13104
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1092
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox SignBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00422E00&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      ForeColor       =   &H00FFFFFF&
      Height          =   2664
      Left            =   468
      ScaleHeight     =   2664
      ScaleWidth      =   6096
      TabIndex        =   6
      Top             =   2184
      Width           =   6096
   End
   Begin VB.TextBox Logs 
      Appearance      =   0  'Flat
      Height          =   5316
      Left            =   7020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   312
      Width           =   5784
   End
   Begin VB.TextBox Sign 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2028
      TabIndex        =   3
      Text            =   "YOUR NAME"
      Top             =   1560
      Width           =   4536
   End
   Begin VB.DriveListBox Drive 
      Height          =   336
      Left            =   468
      TabIndex        =   0
      Top             =   780
      Width           =   6096
   End
   Begin VB.Label Confirm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDB1A&
      Caption         =   "��֤"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Left            =   4524
      TabIndex        =   7
      Top             =   5148
      Width           =   2040
   End
   Begin VB.Label Btn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDB1A&
      Caption         =   "��Ȩ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Left            =   468
      TabIndex        =   4
      Top             =   5148
      Width           =   2040
   End
   Begin VB.Label Tips 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Կ������ǩ����"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   468
      TabIndex        =   2
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label DriveTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ��һ��������������DogTools-v2��Կ��"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   468
      TabIndex        =   1
      Top             =   312
      Width           =   3468
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SignFile
    Owner As String
    KeyString As String
    SignDots() As Single
    SignKey As String
    GenerateTime As Date
    Permission As String
End Type
Dim Dots() As Single
Dim lX As Long, lY As Long

Private Sub Btn_Click()
    LogW "========================="
    LogW "��ʼ��Ȩ"
    
    Dim Name As String, Size As Long, Serial As Long, MaxComp As Long, FileFlags As Long, FileName As String
    Name = Space(255): FileName = Space(255): Size = 255

    Dim DriveS As String
    DriveS = Split(Drive.List(Drive.ListIndex), ":")(0) & ":\"
    LogW "Ŀ���豸��" & DriveS
    GetVolumeInformationA DriveS, Name, Size, Serial, MaxComp, FileFlags, FileName, Size
    LogW "�豸������" & Size & "," & Serial & "," & MaxComp & "," & FileFlags & "," & FileName
    LogW ""
    
    Dim Key As SignFile
    Key.Permission = "Buger404(" & BMEA_Engine.GetBMKey & ")"
    LogW "��Ȩ��Ϣ��" & Key.Permission
    
    Key.GenerateTime = Now
    LogW "��Ȩ���ڣ�" & Key.GenerateTime
    
    Key.Owner = Sign.Text
    LogW "��Կ�����ߣ�" & Key.Owner
    
    Dim md5 As New md5

    LogW "תд�豸����..."
    Key.KeyString = md5.Md5_String_Calc("copied with " & Key.Owner & " by " & BMEA(Size, Key.Permission) & "&&" & BMEA(Serial, Key.Permission) & "&&" & BMEA(MaxComp, Key.Permission) & "&&" & BMEA(FileFlags, Key.Permission) & "&&" & BMEA(FileName, Key.Permission))
    LogW "�ɹ���" & Key.KeyString
    
    LogW "תд��дǩ��(�ܼ�" & UBound(Dots) & "������)..."
    ReDim Key.SignDots(UBound(Dots))
    
    For I = 1 To UBound(Dots)
        Key.SignDots(I) = Dots(I)
        If I Mod __(2)__ = 0 Then LogW "����תдǩ���ĵ�" & I & "�����㣺" & Dots(I) & "(�ܼ�" & UBound(Dots) & "������)..."
        Key.SignKey = Key.SignKey & BMEA(Dots(I), Key.KeyString)
        If Len(Key.SignKey) > 64 Then Key.SignKey = md5.Md5_String_Calc(Key.SignKey)
    Next
    
    LogW "���ںϲ�ǩ��..."
    Key.SignKey = md5.Md5_String_Calc(Key.Owner & " signed " & Key.SignKey & " (" & UBound(Dots) & ")")
    LogW "�ɹ���" & Key.SignKey
    
    LogW "������Կ->""" & DriveS & "dogtools-v2-key.dt2k"""
    If Dir(DriveS & "dogtools-v2-key.dt2k") <> "" Then Kill DriveS & "dogtools-v2-key.dt2k"
    Open DriveS & "dogtools-v2-key.dt2k" For Binary As #1
    Put #1, , Key
    Close #1
    
    LogW "��Ȩ�ɹ���"
    LogW "========================="
End Sub

Public Sub LogW(ByVal str As String)
    Logs.Text = Logs.Text & str & vbCrLf
    Logs.SelStart = Len(Logs.Text) - Logs.SelLength - 1
End Sub

Private Sub Confirm_Click()
    SignBox.Cls

    LogW "========================="
    LogW "��ʼ��֤"
    
    Dim Name As String, Size As Long, Serial As Long, MaxComp As Long, FileFlags As Long, FileName As String
    Name = Space(255): FileName = Space(255): Size = 255

    Dim DriveS As String
    DriveS = Split(Drive.List(Drive.ListIndex), ":")(0) & ":\"
    LogW "Ŀ���豸��" & DriveS
    GetVolumeInformationA DriveS, Name, Size, Serial, MaxComp, FileFlags, FileName, Size
    LogW "�豸������" & Size & "," & Serial & "," & MaxComp & "," & FileFlags & "," & FileName
    LogW ""
    
    Dim Key As SignFile
    Open DriveS & "dogtools-v2-key.dt2k" For Binary As #1
    Get #1, , Key
    Close #1
    
    LogW "��Ȩ��Ϣ��" & Key.Permission
    LogW "��Ȩ���ڣ�" & Key.GenerateTime
    LogW "��Կ�����ߣ�" & Key.Owner
    
    Dim md5 As New md5, con As Boolean

    con = Key.KeyString = md5.Md5_String_Calc("copied with " & Key.Owner & " by " & BMEA(Size, Key.Permission) & "&&" & BMEA(Serial, Key.Permission) & "&&" & BMEA(MaxComp, Key.Permission) & "&&" & BMEA(FileFlags, Key.Permission) & "&&" & BMEA(FileName, Key.Permission))
    LogW "�豸�������飺" & IIf(con, "�ɹ�", "ʧ��")
    If Not con Then Exit Sub
    
    Dim temp As String, TX As Long
    For I = 1 To UBound(Key.SignDots)
        TX = Sqr(Key.SignDots(I) / 500)
        SignBox.Circle (TX + I * 10, Key.SignDots(I) / 250000 + Sin(I) * SignBox.Height), 100, RGB(255, 255 * (TX / SignBox.Width) * (I / UBound(Key.SignDots)), 255 * (ty / SignBox.Height))
        temp = temp & BMEA(Key.SignDots(I), Key.KeyString)
        If Len(temp) > __(3)__ Then temp = md5.Md5_String_Calc(temp)
        DoEvents
    Next
    
    con = Key.SignKey = md5.Md5_String_Calc(Key.Owner & " signed " & temp & " (" & UBound(Key.SignDots) & ")")
    LogW "ǩ�����飺" & IIf(con, "�ɹ�", "ʧ��")
    If Not con Then Exit Sub
    
    LogW "��Ȩ������"
    LogW "========================="
End Sub

Private Sub Form_Load()
    ReDim Dots(0)
End Sub

Private Sub SignBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lX = 0 And lY = 0 Then lX = X: lY = Y
        SignBox.Line (lX, lY)-(X, Y), RGB(255, 255, 255)
        lX = X: lY = Y
        ReDim Preserve Dots(UBound(Dots) + 1)
        Dim b1() As Byte, b2(3) As Byte, Dotish As Single
        ReDim b1(3)
        Dotish = 1
        CopyMemory b1(0), lX, 4: CopyMemory b2(0), lY, 4
        ReDim Preserve b1(7)
        For I = __(1)__ To 7
            b1(I) = b2(I - 4)
        Next
        For I = 0 To 7
            Dotish = Dotish * IIf(b1(I) = 0, I, b1(I))
        Next
        Dots(UBound(Dots)) = Dotish
    End If
End Sub

Private Sub SignBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lX = 0: lY = 0
    If Button = 2 Then SignBox.Cls: ReDim Dots(0)
End Sub

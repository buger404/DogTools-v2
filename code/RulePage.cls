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
'   ����ģ������Emerald������ ҳ������� ģ��
'======================================
'   ҳ�������
    Dim Page As GPage
'======================================

Public Sub Update()
    'Ӧ�ý���ÿһ֡ˢ�»���õķ�����ҳ����ƴ�������ڴ˴���
    
    Page.Clear '��ջ���

    '================================================================
    '   ʾ������������
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
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
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(id As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage  '����ҳ�������
    Page.Create Me  '�����������������
    Page.Res.NewImages App.path & "\assets"  '������Դ
    ECore.Add Page, "RulePage"  '����ҳ��
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub
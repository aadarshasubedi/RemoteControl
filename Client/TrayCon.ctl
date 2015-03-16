VERSION 5.00
Begin VB.UserControl TrayCon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   ClipBehavior    =   0  '��
   FillStyle       =   4  'Upward Diagonal
   HitBehavior     =   0  '��
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   Picture         =   "TrayCon.ctx":0000
   ScaleHeight     =   510
   ScaleWidth      =   465
End
Attribute VB_Name = "TrayCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Download by http://www.codefans.net
Option Explicit
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

'Tip�ϵ�ͼ��
Private Const NIIF_NONE = &H0      'û��ͼ��
Private Const NIIF_WARNING = &H2      '�����桱ͼ�꣨��ɫ�ġ�������
Private Const NIIF_ERROR = &H3      '������ͼ�꣨��ɫ�ġ�X����
Private Const NIIF_INFO = &H1      '����Ϣ��ͼ�꣨��ɫ�ġ�i����
 
Private frmcls As New trayCls

' *********************************************
' ��ϵͳ������ɾ��ͼ���ǳ���Ҫ.
' *********************************************
Public Sub RemoveFromTray()
Call RemoveFromTray_cls
End Sub

Public Sub SetMenu(mnu As Object)
    If mnu Is Nothing Then Else Set TheMenu = mnu
End Sub
' *********************************************
' �����µ���ʾ����
' *********************************************
Public Sub SetTrayTip(tip As String)
    With TheData
        .szTip = tip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' *********************************************
' �����µ�������ʾ���ֱ���
' *********************************************
Public Sub SetTrayTitle(strTitle As String)
    With TheData
        .szInfoTitle = strTitle & vbNullChar
        .uFlags = NIF_INFO Or NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' *********************************************
' �����µ�������ʾ������Ϣ
' *********************************************
Public Sub SetTrayInfo(strInfo As String)
    With TheData
        .szInfo = strInfo & vbNullChar
        .uFlags = NIF_INFO
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' *********************************************
' �����µ�������ʾͼ��
' *********************************************
Public Sub SetTrayInfoFlags(InfoFlags As String)
    With TheData
        .dwInfoFlags = InfoFlags
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

' *********************************************
' �����µ�ͼ��
' *********************************************
Public Sub SetTrayIcon(Optional pic As Picture)
    If pic.Type <> vbPicTypeIcon Then Exit Sub '������Ǹ���Ч��ͼ�꣬������
    With TheData '����
        .hIcon = pic.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

' *********************************************
' ���ϵͳ����
' ��һ��������Ҫ���ڶ���������ѡ��������������ѡ
' ���ڶ�������û��ʱ����ͼ��Ĭ��Ϊ�����icon��
' ������������ֻ�������Ҽ������Ĳ˵���Ҳ����Ϊ��
' *********************************************
Public Sub AddToTray(Frm As Object, Optional mnu As Object, Optional pic As Picture, Optional tip As String = "", Optional strTitle As String = "", Optional strInfo As String = "")
    ' ShowInTaskbar must be set to False at
    ' �����״ֻ̬��������ʱ����д
    frmcls.setForm Frm '����������from
    ' ���ô���Ͳ˵�
    Set TheForm = Frm

    If mnu Is Nothing Then Else Set TheMenu = mnu
    OldWindowProc = SetWindowLong(Frm.hwnd, _
        GWL_WNDPROC, AddressOf NewWindowProc)
    
    ' Install the form's icon in the tray.
    With TheData
    
    .uId = 0
    .hwnd = Frm.hwnd
    .cbSize = Len(TheData)
    If pic Is Nothing Then
    .hIcon = Frm.Icon.Handle
    Else
    .hIcon = pic
    End If
    .uCallBackMessage = TRAY_CALLBACK
    .cbSize = Len(TheData)
    
    .uFlags = NIF_INFO Or NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .szTip = tip & vbNullChar
    .szInfoTitle = strTitle & vbNullChar
    .szInfo = strInfo & vbNullChar
    
    .dwInfoFlags = NIIF_INFO

    End With
    Shell_NotifyIcon NIM_ADD, TheData
    
    'SetTrayTip tip
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 450
UserControl.Height = 450

End Sub

Attribute VB_Name = "APIStuff"
' Download by http://www.codefans.net
Option Explicit
Public OldWindowProc As Long
Public TheForm As Object
Public TheMenu As Object
Public laststate As Integer

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'�򴰿ڷ�����Ϣ
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Private Const WM_USER = &H400
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONUP = &H205
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_USERDATA = (-21)
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const SC_RESTORE = &HF120&      '��ԭ����
Public Const WM_SYSCOMMAND = &H112      'ϵͳ��Ϣ
Public Type NOTIFYICONDATA
    cbSize  As Long     'NOTIFYICONDATA���͵Ĵ�С����Len(������)��ü���
    hwnd  As Long       '���������
    uId  As Long        'ͼ����Դ��ID�ţ�ͨ��ʹ��  vbNull
    uFlags  As Long     'ʹ��Щ������Ч��������ö�������е�  NIF_INFO  Or  NIF_ICON  Or  NIF_TIP  Or  NIF_MESSAGE  �ĸ����������
    uCallBackMessage  As Long   '������Ϣ���¼�
    hIcon  As Long      'ͼ������
    szTip  As String * 128     '�����ͣ����ͼ����ʱ��ʾ��Tip�ı�
    dwState  As Long    'ͨ��Ϊ  0
    dwStateMask  As Long 'ͨ��Ϊ  0
    szInfo  As String * 256         'Tip�ı�����
    uTimeoutOrVersion  As Long      '����VB��û��Union���ͣ�ֻ����Long�ʹ���
    szInfoTitle  As String * 64     'Tip�ı��ı���
    dwInfoFlags  As Long
End Type

Public TheData As NOTIFYICONDATA
' *********************************************
' ��������ϵͳ����
' *********************************************
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
On Error Resume Next
    If Msg = TRAY_CALLBACK Then
        ' �û�����ͼ��
        Select Case lparam
            Case WM_LBUTTONUP  '���������ʾ����
                If TheForm.WindowState = 1 Then
                    TheForm.Show
                   SendMessage hwnd, WM_SYSCOMMAND, SC_RESTORE, 0      '��ԭ���ڴ�С
                    Exit Function
                End If
            Case WM_RBUTTONUP   '�����Ҽ���ʾ�˵�
                TheForm.PopupMenu TheMenu
                Exit Function
            Case 515 '˫�����
                frmmain.Show
                    frmmain.chkautohide.Value = 0
                   AutoHide = 0
                    WriteINI "Local", "AutoHide", AutoHide, GetPath & "data\set.kac"
                Exit Function
        End Select
    End If
    ' ������ϵͳ
    NewWindowProc = CallWindowProc( _
        OldWindowProc, hwnd, Msg, _
        wParam, lparam)
End Function
' *********************************************
' ɾ������
' *********************************************
Public Sub RemoveFromTray_cls()
    ' ɾ��ͼ��
    With TheData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData
     ' ��ϵͳɾ��
    SetWindowLong TheForm.hwnd, GWL_WNDPROC, _
        OldWindowProc
End Sub

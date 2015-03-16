Attribute VB_Name = "HotKeys"
Option Explicit

'�ڴ��ڽṹ��Ϊָ���Ĵ���������Ϣ
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'��ָ�����ڵĽṹ��ȡ����Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'����ָ���Ľ���
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��ϵͳע��һ��ָ�����ȼ�
Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'ȡ���ȼ����ͷ�ռ�õ���Դ
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long
'�������API������ע��ϵͳ���ȼ�������ģ�����ʵ�ֹ����������ʾ

  '�ȼ���־����,�����жϵ����̰���������ʱ�Ƿ������������趨���ȼ�
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)

'����ϵͳ���ȼ�,ԭ�жϱ�ʾ,�����ص���Ŀ���
Public preWinProc As Long, MyhWnd As Long, uVirtKey As Long
Private CountKey As Integer
'�ȼ����ع���
Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then     '������ص��ȼ���־����
        If wParam = 1 Then      '��������ǵĶ�����ȼ�...
            HideDone            'ִ�����������ָ��Ŀ
        End If
      End If
    '��������ȼ�,���߲����������õ��ȼ�,��������Ȩ��ϵͳ,��������ȼ�
    WndProc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function

'��ؼ�����Ŀ���ع���
Public Sub HideDone()
    CountKey = CountKey + 1
    If CountKey = 5 Then
        CountKey = 0
        If frmmain.Visible Then
            frmmain.Hide
            frmmain.chkautohide.Value = 1
        Else
            frmmain.Show
            frmmain.chkautohide.Value = 0
        End If
        AutoHide = frmmain.chkautohide.Value
        WriteINI "Local", "AutoHide", AutoHide, GetPath & "data\set.kac"
    End If
End Sub


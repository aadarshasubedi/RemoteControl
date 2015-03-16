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
'向窗口发送消息
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
Public Const SC_RESTORE = &HF120&      '还原窗口
Public Const WM_SYSCOMMAND = &H112      '系统消息
Public Type NOTIFYICONDATA
    cbSize  As Long     'NOTIFYICONDATA类型的大小，用Len(变量名)获得即可
    hwnd  As Long       '窗体的名柄
    uId  As Long        '图标资源的ID号，通常使用  vbNull
    uFlags  As Long     '使哪些参数有效它是以下枚举类型中的  NIF_INFO  Or  NIF_ICON  Or  NIF_TIP  Or  NIF_MESSAGE  四个常数的组合
    uCallBackMessage  As Long   '接受消息的事件
    hIcon  As Long      '图标名柄
    szTip  As String * 128     '当鼠标停留在图标上时显示的Tip文本
    dwState  As Long    '通常为  0
    dwStateMask  As Long '通常为  0
    szInfo  As String * 256         'Tip文本正文
    uTimeoutOrVersion  As Long      '由于VB中没有Union类型，只能用Long型代替
    szInfoTitle  As String * 64     'Tip文本的标题
    dwInfoFlags  As Long
End Type

Public TheData As NOTIFYICONDATA
' *********************************************
' 交给操作系统处理
' *********************************************
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
On Error Resume Next
    If Msg = TRAY_CALLBACK Then
        ' 用户单击图标
        Select Case lparam
            Case WM_LBUTTONUP  '按下左键显示窗体
                If TheForm.WindowState = 1 Then
                    TheForm.Show
                   SendMessage hwnd, WM_SYSCOMMAND, SC_RESTORE, 0      '还原窗口大小
                    Exit Function
                End If
            Case WM_RBUTTONUP   '按下右键显示菜单
                TheForm.PopupMenu TheMenu
                Exit Function
            Case 515 '双击左键
                frmmain.Show
                    frmmain.chkautohide.Value = 0
                   AutoHide = 0
                    WriteINI "Local", "AutoHide", AutoHide, GetPath & "data\set.kac"
                Exit Function
        End Select
    End If
    ' 交还给系统
    NewWindowProc = CallWindowProc( _
        OldWindowProc, hwnd, Msg, _
        wParam, lparam)
End Function
' *********************************************
' 删除托盘
' *********************************************
Public Sub RemoveFromTray_cls()
    ' 删除图标
    With TheData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData
     ' 从系统删除
    SetWindowLong TheForm.hwnd, GWL_WNDPROC, _
        OldWindowProc
End Sub

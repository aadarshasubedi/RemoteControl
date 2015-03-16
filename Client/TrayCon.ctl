VERSION 5.00
Begin VB.UserControl TrayCon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   ClipBehavior    =   0  '无
   FillStyle       =   4  'Upward Diagonal
   HitBehavior     =   0  '无
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

'Tip上的图标
Private Const NIIF_NONE = &H0      '没有图标
Private Const NIIF_WARNING = &H2      '“警告”图标（黄色的“！”）
Private Const NIIF_ERROR = &H3      '“错误”图标（红色的“X”）
Private Const NIIF_INFO = &H1      '“消息”图标（蓝色的“i”）
 
Private frmcls As New trayCls

' *********************************************
' 从系统托盘里删除图表，非常重要.
' *********************************************
Public Sub RemoveFromTray()
Call RemoveFromTray_cls
End Sub

Public Sub SetMenu(mnu As Object)
    If mnu Is Nothing Then Else Set TheMenu = mnu
End Sub
' *********************************************
' 设置新的提示文字
' *********************************************
Public Sub SetTrayTip(tip As String)
    With TheData
        .szTip = tip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' *********************************************
' 设置新的气泡提示文字标题
' *********************************************
Public Sub SetTrayTitle(strTitle As String)
    With TheData
        .szInfoTitle = strTitle & vbNullChar
        .uFlags = NIF_INFO Or NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' *********************************************
' 设置新的气泡提示文字信息
' *********************************************
Public Sub SetTrayInfo(strInfo As String)
    With TheData
        .szInfo = strInfo & vbNullChar
        .uFlags = NIF_INFO
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' *********************************************
' 设置新的气泡提示图标
' *********************************************
Public Sub SetTrayInfoFlags(InfoFlags As String)
    With TheData
        .dwInfoFlags = InfoFlags
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

' *********************************************
' 设置新的图标
' *********************************************
Public Sub SetTrayIcon(Optional pic As Picture)
    If pic.Type <> vbPicTypeIcon Then Exit Sub '如果不是个有效的图标，则跳出
    With TheData '更新
        .hIcon = pic.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

' *********************************************
' 添加系统托盘
' 第一个参数必要，第二个参数可选，第三个参数可选
' 当第二个参数没有时，则图表默认为窗体的icon，
' 第三个参数是只电击鼠标右键弹出的菜单，也可以为空
' *********************************************
Public Sub AddToTray(Frm As Object, Optional mnu As Object, Optional pic As Picture, Optional tip As String = "", Optional strTitle As String = "", Optional strInfo As String = "")
    ' ShowInTaskbar must be set to False at
    ' 在设计状态只读，运行时可以写
    frmcls.setForm Frm '用类来处理from
    ' 设置窗体和菜单
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

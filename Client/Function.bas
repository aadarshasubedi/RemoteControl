Attribute VB_Name = "Function"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpszOp As String, _
ByVal lpszFile As String, ByVal lpszParams As String, _
ByVal LpszDir As String, ByVal FsShowCmd As Long) _
As Long
Public Function GetPath() As String
    If Right(App.Path, 1) = "/" Then
        GetPath = App.Path
    Else
        GetPath = App.Path & "/"
    End If
End Function

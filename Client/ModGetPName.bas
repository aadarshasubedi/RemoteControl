Attribute VB_Name = "ModGetPName"
Option Explicit
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Function GetPName() As String
Dim i As Long, j As Long
Dim tStr As String * 254
On Error GoTo 10
i = GetForegroundWindow
j = GetWindowText(i, tStr, Len(tStr) + 1)
GetPName = Left(tStr, j)
Exit Function
10:
GetPName = ""
End Function

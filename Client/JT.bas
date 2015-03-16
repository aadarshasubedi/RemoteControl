Attribute VB_Name = "JT"
Option Explicit

Public Type GUID
     Data1 As Long
     Data2 As Integer
     Data3 As Integer
     Data4(0 To 7) As Byte
End Type

Public Type GdiplusStartupInput
     GdiplusVersion As Long
     DebugEventCallback As Long
     SuppressBackgroundThread As Long
     SuppressExternalCodecs As Long
End Type

Public Type EncoderParameter
     GUID As GUID
     NumberOfValues As Long
     type As Long
     Value As Long
End Type

Public Type EncoderParameters
     Count As Long
     Parameter As EncoderParameter
End Type

Public Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Public Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Public Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Public Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Public Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, ID As GUID) As Long
Public Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, Bitmap As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const theScreen = 0
Public Const TheForms = 1


Public Sub SaveScreenJPG(ByVal TargetFileName As String, Optional ByVal quality As Byte = 80)

     Dim hDCscr As Long
     Dim hDCmem As Long
     Dim hBmp As Long

     Dim hBmpPrev As Long
     Dim lWidth As Long, lHeight As Long
     lWidth = GetSystemMetrics(SM_CXSCREEN)
     lHeight = GetSystemMetrics(SM_CYSCREEN)

     '常规抓图代码，得到一个hBmp：
     hDCscr = GetDC(0)
     hDCmem = CreateCompatibleDC(hDCscr)
     hBmp = CreateCompatibleBitmap(hDCscr, lWidth, lHeight)
     hBmpPrev = SelectObject(hDCmem, hBmp)
     BitBlt hDCmem, 0, 0, lWidth, lHeight, hDCscr, 0, 0, SRCCOPY
     SelectObject hDCmem, hBmpPrev
     DeleteDC hDCmem
     ReleaseDC 0, hDCscr

    '通过GDI+将hBmp存为JPG文件
     Dim tSI As GdiplusStartupInput
     Dim lRes As Long
     Dim lGDIP As Long
     Dim lBitmap As Long
    
     tSI.GdiplusVersion = 1
     lRes = GdiplusStartup(lGDIP, tSI, 0)
     If lRes = 0 Then
     lRes = GdipCreateBitmapFromHBITMAP(hBmp, 0, lBitmap)
     If lRes = 0 Then
         Dim tJpgEncoder As GUID
         Dim tParams As EncoderParameters
         CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder

         tParams.Count = 1
         With tParams.Parameter
            CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
            .NumberOfValues = 1
            .type = 4
            .Value = VarPtr(quality)
         End With
         lRes = GdipSaveImageToFile(lBitmap, StrPtr(TargetFileName), tJpgEncoder, tParams)
         GdipDisposeImage lBitmap
         End If
         GdiplusShutdown lGDIP
     End If
End Sub




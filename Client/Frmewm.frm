VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmewm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "二维条码"
   ClientHeight    =   7260
   ClientLeft      =   9150
   ClientTop       =   3360
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6555
   Begin MSComDlg.CommonDialog cmmd 
      Left            =   2400
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save this picture"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   6720
      Width           =   1935
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      Index           =   3
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Text encoding"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form1.frx":47062
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      Index           =   2
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Mask type"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      Index           =   1
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Error correction level"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      Index           =   0
      ItemData        =   "Form1.frx":4706D
      Left            =   600
      List            =   "Form1.frx":4706F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Version"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "创建"
      Default         =   -1  'True
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmewm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByRef lpUsedDefaultChar As Any) As Long
Private Const CP_UTF8 As Long = 65001

Private obj As New clsQRCode

Private Sub Command1_Click()
    Dim b2() As Byte
    Dim s As String
    Dim i As Long, m As Long
    For i = 0 To cmb1.UBound
        If cmb1(i).ListIndex < 0 Then Exit Sub
    Next i
    Select Case cmb1(3).ListIndex
        Case 1
        s = Text1.Text
        m = Len(s)
        i = m * 3 + 64
        ReDim b2(i)
        m = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(s), m, b2(0), i, ByVal 0, ByVal 0)
        Case Else
        s = StrConv(Text1.Text, vbFromUnicode)
        b2 = s
        m = LenB(s)
    End Select
    Set Image1.Picture = obj.Encode(b2, m, cmb1(0).ListIndex, cmb1(1).ListIndex + 1, cmb1(2).ListIndex - 1)
End Sub

Private Sub Command2_Click()
    cmmd.Filter = "BMP图像|*.bmp"
    cmmd.ShowSave
    If cmmd.FileName = "" Then Exit Sub
    With Picture1
        .AutoRedraw = True
        .Width = Image1.Width
        .Height = Image1.Height
         .PaintPicture Image1.Picture, 0, 0, Image1.Width, Image1.Height
        SavePicture Picture1.Image, cmmd.FileName
    End With
End Sub

Private Sub Form_Load()
    Dim i As Long
    cmb1(0).AddItem "自动"
    For i = 1 To 40
        cmb1(0).AddItem CStr(i)
    Next i
    cmb1(0).ListIndex = 0
    cmb1(1).AddItem "L - 7%"
    cmb1(1).AddItem "M - 15%"
    cmb1(1).AddItem "Q - 25%"
    cmb1(1).AddItem "H - 30%"
    cmb1(1).ListIndex = 2
    cmb1(2).AddItem "自动"
    For i = 0 To 7
        cmb1(2).AddItem CStr(i)
    Next i
    cmb1(2).ListIndex = 0
    cmb1(3).AddItem "ANSI"
    cmb1(3).AddItem "UTF-8"
    cmb1(3).ListIndex = 1
End Sub


Private Sub Text1_Change()
    Command1_Click
End Sub

Private Sub Text1_DblClick()
  Text1 = ""
End Sub

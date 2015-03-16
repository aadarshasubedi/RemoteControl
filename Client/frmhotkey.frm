VERSION 5.00
Begin VB.Form frmhotkey 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "热键设置"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "取消热键"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "当前热键"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   " 您可以自定义您的热键，在您连续按5此热键后，窗体会弹出或隐藏，热键必须是A-Z中的一个字母，重启后生效"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmhotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "设置无效", vbCritical
        Exit Sub
    End If
    PauseHotKey = Check1.Value
    WriteINI "Local", "PauseHotKey", PauseHotKey, GetPath & "data\set.kac"
     WriteINI "Local", "HotKey", Left(UCase(Text1.Text), 1), GetPath & "data\set.kac"
     MsgBox "设置成功，重启后生效", vbInformation
End Sub

Private Sub Form_Load()
    Check1.Value = PauseHotKey
    Label1.Caption = "您可以自定义您的热键，在您连续按5此热键后，窗体会弹出或隐藏，热键必须是A-Z中的一个字母，重启后生效"
    Text1.Text = Chr(HotKey)
End Sub

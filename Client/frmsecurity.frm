VERSION 5.00
Begin VB.Form frmsecurity 
   Caption         =   "������֤������"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text3 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "��֤����"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "�ظ�����"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "��������"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmsecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Countspass As Integer
Private Sub Command1_Click()
    If Not LoginState Then
        If Text1.Text = "" Then
            MsgBox "���벻��Ϊ��", vbCritical
            Exit Sub
        Else
            If Text1.Text <> Text2.Text Then
                MsgBox "�������벻һ�£����飡", vbCritical
                Exit Sub
            End If
            Dim w As Object
            Set w = CreateObject("wscript.shell")
            LoginPassword = MD5.DigestStrToHexStr(Text1.Text)
            w.regwrite "HKCU\Software\Starainrt\RC\Password", LoginPassword, "REG_SZ"
        End If
        If Me.Caption <> "�޸�����" Then
            Load frmmain
            If AutoHide = 0 Then frmmain.Show Else frmmain.Hide
        Else
            MsgBox "�޸ĳɹ���", vbInformation
            Unload Me
        End If
    Else
        If LoginPassword <> MD5.DigestStrToHexStr(Text3.Text) Then
            MsgBox "���벻��ȷ", vbCritical
            Countspass = Countspass + 1
            If Countspass > 4 Then
                MsgBox "����������࣡����", vbCritical
                End
            End If
            Exit Sub
        Else
               Load frmmain
                If AutoHide = 0 Then frmmain.Show Else frmmain.Hide
        End If
    End If
    App.TaskVisible = False
    Unload Me
End Sub

Private Sub Form_Load()
    If Not LoginState Then
        Label2.Visible = True
        Text1.Visible = True
        Label3.Visible = True
        Text2.Visible = True
        Text3.Visible = False
        Label4.Visible = False
    Else
        Label2.Visible = False
        Text1.Visible = False
        Label3.Visible = False
        Text2.Visible = False
        Text3.Visible = True
        Label4.Visible = True
    End If
    App.TaskVisible = True
End Sub

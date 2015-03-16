VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Frmpost 
   Caption         =   "HTML"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Frmpost"
   ScaleHeight     =   5310
   ScaleWidth      =   7395
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "App.Path"
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   960
      TabIndex        =   17
      Top             =   4200
      Width           =   5295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Put File"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4680
      TabIndex        =   14
      Text            =   "102400"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "stop"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLS"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   2520
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Text            =   "5000"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Text            =   "utf-8"
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "http://www.baidu.com"
      Top             =   360
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Post"
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6360
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Caption         =   "FileName"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "byta num"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "state"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Timeout"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "infor"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "website"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Frmpost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()

Dim myurl As String, send_data As String, my_head As String
On Error GoTo errhand:
Inet1.RequestTimeout = Text4.Text

myurl = Text1.Text

send_data = Trim(Text2.Text)

my_head = "Content-Type: application/x-www-form-urlencoded"

Inet1.Execute myurl, "POST", send_data, my_head
Exit Sub
errhand:
MsgBox Err.Description
End Sub

Private Sub Command2_Click()
On Error GoTo errhand:
Dim myurl As String, send_data As String, my_head As String
Inet1.RequestTimeout = Text4.Text
myurl = Text1.Text

send_data = Trim(Text2.Text)

my_head = "Content-Type: application/x-www-form-urlencoded"

Inet1.Execute myurl, "GET"
Exit Sub
errhand:
MsgBox Err.Description
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
End Sub

Private Sub Command4_Click()
    If Inet1.StillExecuting Then
        Inet1.Cancel
    End If
End Sub

Private Sub Command5_Click()
Text7.Text = App.Path & "\Temp.Temp"
End Sub

Private Sub Form_Load()
Text7.Text = App.Path & "\Temp.Temp"
End Sub

Private Sub Inet1_StateChanged(ByVal state As Integer)
On Error GoTo errhand:
Dim a As String
    Text5.Text = state
    Dim get_data() As Byte
    If state = 12 Then
        get_data = Inet1.GetChunk(Val(Text6.Text), icByteArray)
        If chk.Value = 0 Then
            a = BytesToBstr(get_data, Text3.Text)
            ' MsgBox a
            ' Open App.Path & "\read.txt" For Output As #1
        '         Print #1, a
        ' Close #1
            Frmpostshow.Text1.Text = ""
            Frmpostshow.Text1.Text = a
            Frmpostshow.Show
        Else
            Dim filenames As String
            filenames = Trim(Text7.Text)
            Dim nfso As New FileSystemObject
            If nfso.FileExists(filenames) Then
                If MsgBox("是否替换原文件？", vbYesNo) = vbYes Then
                    Kill filenames
                Else
                    Exit Sub
                End If
            End If
            Open filenames For Binary As #1
                Put #1, 1, get_data
            Close #1
            If MsgBox("是否打开？", vbQuestion + vbYesNo) = vbYes Then
                ShellExecute Me.hwnd, "open", filenames, "", "", 1
            End If
        End If
    End If
    If state = 11 Then
        MsgBox "获取失败"
    End If
    Exit Sub
errhand:
    MsgBox Err.Number & " " & Err.Description
End Sub

Function BytesToBstr(strBody, CodeBase)
Dim ObjStream
Set ObjStream = CreateObject("Adodb.Stream")
With ObjStream
.Type = 1
.Mode = 3
.Open
.Write strBody
.Position = 0
.Type = 2
.Charset = CodeBase
BytesToBstr = .ReadText
.Close
End With
Set ObjStream = Nothing
End Function


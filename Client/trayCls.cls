VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "trayCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Download by http://www.codefans.net
Option Explicit
'如何识别TrayForm是FROM还是MDIFORM???
'识别不了只能用OBJECT.设
'设Public TrayForm As form 只能用在FROM
'设Public TrayForm As mdiform 只能用在mdifrom
'设Public TrayForm As Object 在from mdifrom通用
'但是TrayForm_Unload,及TrayForm_Resize用不了,站长可否解决一下
Private WithEvents TrayForm As Form
Attribute TrayForm.VB_VarHelpID = -1
Private WithEvents TrayMdiForm As MDIForm
Attribute TrayMdiForm.VB_VarHelpID = -1

Public Sub setForm(Frm As Object)
Dim i As String
  On Error GoTo err
  i = Frm.AutoShowChildren
  Set TrayMdiForm = Frm
Exit Sub
err:
  If err.Number = 438 Then
  Set TrayForm = Frm
End If
End Sub
' 处理事件，相信大家都看得懂，如果不懂，请到www.play78.com提问
Public Sub TrayForm_Resize()
If laststate = 1 Then laststate = 0
    If TrayForm.WindowState <> vbMinimized Then _
        laststate = TrayForm.WindowState
End Sub

' 当窗体退出时，删除系统托盘
Private Sub TrayForm_Unload(Cancel As Integer)
    RemoveFromTray_cls
End Sub

Public Sub TrayMdiForm_Resize()
If laststate = 1 Then laststate = 0
    If TrayMdiForm.WindowState <> vbMinimized Then _
        laststate = TrayMdiForm.WindowState
End Sub

' 当窗体退出时，删除系统托盘
Private Sub TrayMdiForm_Unload(Cancel As Integer)
    TrayForm_Unload Cancel
End Sub


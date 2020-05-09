VERSION 5.00
Begin VB.Form xinxixiugai 
   Caption         =   "信息修改"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   15195
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "xinxixiugai.frx":0000
      Left            =   3960
      List            =   "xinxixiugai.frx":000A
      TabIndex        =   15
      Text            =   "普通"
      Top             =   3840
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "xinxixiugai.frx":0019
      Left            =   3960
      List            =   "xinxixiugai.frx":002F
      TabIndex        =   13
      Text            =   "剧情"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   615
      Left            =   5400
      TabIndex        =   12
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存修改"
      Height          =   615
      Left            =   2760
      TabIndex        =   11
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   3
      Left            =   3960
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   2
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   1
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   5
      Left            =   2640
      TabIndex        =   0
      Top             =   4920
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "类型"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "联系电话"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "电子邮箱"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户偏好"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "生日"
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "个人介绍"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   4440
      Width           =   735
   End
End
Attribute VB_Name = "xinxixiugai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs7 As New ADODB.Recordset
Dim sql As String
Dim a(6) As String
Private Sub Command1_Click()
If Combo2.Text = "VIP" Then
    c = MsgBox("是否升级？确定支付5元/月？", vbYesNo)
 If c = vbYes Then
    MsgBox "支付成功"
    a(6) = Combo2.Text
Else
    Combo2.Text = "普通"
    a(6) = Combo2.Text
End If
 End If
If Len(Text1(1).Text) <> 11 Then MsgBox "请输入11位手机号"
i = 1
For i = 1 To 11
 b = Right(Text1(1).Text, i)
 If Asc(b) < 48 Or Asc(b) > 57 Then MsgBox ("号码有非数字字符错误")
Next i
For i = 0 To 3
a(i) = Text1(i).Text
Next i
a(4) = Combo1.Text
a(5) = Text1(5).Text
sql = "update user set userphone='" & a(1) & "',usermail='" & a(2) & "',userbir='" & a(3) & "',userprefer='" & a(4) & "',userresume='" & a(5) & "',usertype='" & a(6) & "' where userid='" & a(0) & "' "
cnmovie.Execute sql
  MsgBox "修改完毕！"
For i = 0 To 3
  Text1(i).Enabled = False
  Next i
  Combo1.Enabled = False
  Text1(5).Enabled = False
End Sub

Private Sub Command2_Click()
b = MsgBox("确定不再修改？", vbYesNo)
 If b = vbYes Then movie.Show
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
Text1(0).Text = uid
rs7.CursorLocation = adUseClient
rs7.Open "select  *  from user where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
Text1(0).Enabled = False
Text1(1).Text = rs7.Fields("userphone")
Text1(2).Text = rs7.Fields("usermail")
Text1(3).Text = rs7.Fields("userbir")
Text1(5).Text = rs7.Fields("userresume")
Combo1.Text = rs7.Fields("userprefer")
Combo2.Text = rs7.Fields("usertype")
End Sub
Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

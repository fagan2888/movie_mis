VERSION 5.00
Begin VB.Form guanliyonghu 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   15165
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command9 
      Caption         =   "取消"
      Height          =   495
      Left            =   9360
      TabIndex        =   27
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "保存修改"
      Height          =   495
      Left            =   7920
      TabIndex        =   26
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "删除用户"
      Height          =   495
      Left            =   6360
      TabIndex        =   25
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "修改权限"
      Height          =   495
      Left            =   4920
      TabIndex        =   24
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   8040
      TabIndex        =   13
      Top             =   3600
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   5160
      TabIndex        =   12
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   5160
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   5160
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   0
      Left            =   5160
      TabIndex        =   9
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   2
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   3
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   4
      Left            =   5160
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一条"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "前一条"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "后一条"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "末一条"
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "返回"
      Height          =   495
      Left            =   9360
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "说明：管理员没有资格添加用户或者修改除权限外的信息"
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   7200
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "个人介绍"
      Height          =   255
      Left            =   7920
      TabIndex        =   23
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "头像"
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "生日"
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   21
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户偏好"
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   20
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "性别"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   19
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "类型"
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   18
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "电子邮箱"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "联系电话"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "密码"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "guanliyonghu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs6 As New ADODB.Recordset
Private Sub Command6_Click()
Text1(5).Enabled = True
MsgBox "请修改为普通或VIP"
  Command5.Enabled = False
  Command8.Enabled = True
  Command9.Enabled = True
  Command7.Enabled = False
  Command1.Enabled = False
  Command2.Enabled = False
  Command3.Enabled = False
  Command4.Enabled = False
  Command6.Enabled = False
End Sub

Private Sub Command7_Click()
  b = MsgBox("是否要删除该记录？", vbYesNo)
 If b = vbYes Then
  a = "delete from user where userid='"
  a = a + Text1(0).Text + "'"
  cnn.Execute a
  rs6.Close
  sql = "select * from user"
  rs6.Open sql, cnmovie, adOpenDynamic, adLockOptimistic
     If rs.BOF And rs.EOF Then
       MsgBox "表中无记录！"
     Else
       rs6.MoveFirst
     Call viewdata
   End If
 End If
End Sub

Private Sub Command8_Click()
a = "update user set usertype='" & Text1(5).Text & "' where userid='" & Text1(0).Text & "'"
  cnn.Execute a
  MsgBox "修改完毕！"
  Command9.Enabled = False
  Command8.Enabled = False
  Command5.Enabled = True
  Command6.Enabled = True
  Command7.Enabled = True
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  Command4.Enabled = True
End Sub

Private Sub Command9_Click()
rs6.CancelUpdate
  rs6.MoveFirst
  Call viewdata
  Command9.Enabled = False
  Command8.Enabled = False
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  Command4.Enabled = True
  Command5.Enabled = True
  Command6.Enabled = True
  Command7.Enabled = True
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
For i = 0 To 7
  Text1(i).Enabled = False
Next i
Text2.Enabled = False
rs6.CursorLocation = adUseClient
rs6.Open "select  *  from user", cnmovie, adOpenDynamic, adLockOptimistic
Call viewdata

End Sub
Private Sub Command1_Click()
rs6.MoveFirst
Call viewdata
End Sub

Private Sub Command2_Click()
rs6.MovePrevious
If rs6.BOF Then rs6.MoveFirst
Call viewdata
End Sub

Private Sub Command3_Click()
rs6.MoveNext
If rs6.EOF Then rs6.MoveLast
Call viewdata
End Sub

Private Sub Command4_Click()
rs6.MoveLast
Call viewdata
End Sub

Private Sub Command5_Click()
rs6.Close
Set rs6 = Nothing
guanliyuan.Show
End Sub
Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub
Private Sub viewdata()
For i = 0 To 7
  Text1(i).Enabled = False
Next i
Text2.Enabled = False

For i = 0 To 7
  Text1(i).Text = rs6.Fields(i)
Next i
Text2.Text = rs6.Fields("userresume")
Image1.Picture = LoadPicture(App.Path + "\..\photo\用户头像\" + rs6.Fields("userphoto"))
End Sub


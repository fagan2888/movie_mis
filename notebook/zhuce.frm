VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form zhuce 
   Caption         =   "注册"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   15015
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "上传头像"
      Height          =   495
      Left            =   13200
      TabIndex        =   41
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   0
      ItemData        =   "zhuce.frx":0000
      Left            =   4320
      List            =   "zhuce.frx":000A
      TabIndex        =   35
      Text            =   "普通"
      Top             =   3600
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13320
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   615
      Left            =   6600
      TabIndex        =   25
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   615
      Left            =   4440
      TabIndex        =   24
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   23
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   5520
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   3
      ItemData        =   "zhuce.frx":0019
      Left            =   4320
      List            =   "zhuce.frx":0029
      TabIndex        =   21
      Text            =   "你的管信老师是？"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   2
      ItemData        =   "zhuce.frx":0071
      Left            =   4320
      List            =   "zhuce.frx":0087
      TabIndex        =   20
      Text            =   "剧情"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   1
      ItemData        =   "zhuce.frx":00AF
      Left            =   4320
      List            =   "zhuce.frx":00B9
      TabIndex        =   19
      Text            =   "男"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   4
      Left            =   4320
      TabIndex        =   18
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   3
      Left            =   4320
      TabIndex        =   17
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   2
      Left            =   4320
      TabIndex        =   16
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "zhuce.frx":00C5
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "VIP点播电影免费，月套餐5元/月"
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "给老师的说明：管理员账号为系统后台管理人员分配的，不能注册，因为实际生活中管理员账号随意注册会引起混乱"
      Height          =   1215
      Left            =   600
      TabIndex        =   39
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "我们需要您的完整邮箱地址"
      Height          =   255
      Left            =   6960
      TabIndex        =   38
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "注意注意，标星号为必填项哦"
      Height          =   255
      Left            =   4080
      TabIndex        =   37
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "限于时间精力，类型不是特别全，望大家谅解"
      Height          =   375
      Left            =   6840
      TabIndex        =   36
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "我们的格式是2011/1/1"
      Height          =   255
      Left            =   6960
      TabIndex        =   34
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "在这里填写11位手机号"
      Height          =   255
      Left            =   6960
      TabIndex        =   33
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "在这里展示你的风采吧，也可以使用默认头像,"
      Height          =   1095
      Left            =   13080
      TabIndex        =   32
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   4
      Left            =   6480
      TabIndex        =   31
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   3
      Left            =   6480
      TabIndex        =   30
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   2
      Left            =   6480
      TabIndex        =   29
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   1
      Left            =   6480
      TabIndex        =   28
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   0
      Left            =   6480
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "用户注册："
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "密码"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "联系电话"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "电子邮箱"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "类型"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "性别"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户偏好"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "密保问题"
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "密保答案"
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   6
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "生日"
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "头像"
      Height          =   255
      Index           =   10
      Left            =   8880
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户介绍"
      Height          =   255
      Index           =   11
      Left            =   8760
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "确认密码"
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "zhuce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim str As String
Dim i As Integer, c As String, d As String
Dim a(12) As String
Private Sub Command1_Click()
If Combo1(0).Text = "VIP" Then
    b = MsgBox("确定支付5元/月？", vbYesNo)
 If b = vbYes Then MsgBox "支付成功"
 End If
For i = 0 To 3
  If Text1(i).Text = "" Then
   MsgBox "该项为必填项哦！"
   Text1(i).SetFocus
   Exit Sub
  End If
  Next i
  If Len(Text1(2).Text) <> 11 Then
  MsgBox "请输入11位手机号"
  Exit Sub
  Text1(2).SetFocus
  End If
  If Text1(1).Text <> Text4.Text Then MsgBox "两次输入密码不相同，请检查！"
i = 1
For i = 1 To 11
 d = Right(Text1(2).Text, i)
 If Asc(d) < 48 Or Asc(d) > 57 Then MsgBox ("号码有非数字字符错误")
Next i
str = "select * from user where userid='" & Trim(Text1(0).Text) & "';"
rs2.Open str, cnmovie, adOpenDynamic, adLockPessimistic
   If Not rs2.EOF Then
     MsgBox "用户名已存在，请重新输入"
     Text1(0).Text = ""
     Text1(0).SetFocus
     rs2.Close
   Exit Sub
   End If
i = 0
  For i = 0 To 4
    a(i) = Text1(i).Text
    Next i
i = 0
For i = 0 To 3
    a(i + 5) = Combo1(i).Text
    Next i
    a(9) = Text3.Text
    a(11) = Text2.Text
    If Command3.Enabled = False Then
    a(10) = Text1(0).Text & ".jpg"
    FileCopy CommonDialog1.FileName, App.Path & "\..\photo\用户头像\" & a(10)
    Else
     a(10) = "moren.jpg"
    End If
 c = "insert into user(userid,userpassword,userphone,usermail,userbir,usertype,usersex,userprefer,userquestion,useranswer,userphoto,userresume)  values('" & a(0) & "','" & a(1) & "','" & a(2) & "','" & a(3) & "','" & a(4) & "','" & a(5) & "','" & a(6) & "','" & a(7) & "','" & a(8) & "','" & a(9) & "','" & a(10) & "','" & a(11) & "')"
  cnmovie.Execute c
  MsgBox "注册成功！"
End Sub

Private Sub Command2_Click()
b = MsgBox("别走，还没完成注册呢，真的要离开么？", vbYesNo)
 If b = vbYes Then shouye.Show
End Sub


Private Sub Command3_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Command3.Enabled = False
Image1.Picture = LoadPicture(CommonDialog1.FileName)
End If
'CommonDialog1.ShowSave
'FileCopy CommonDialog1.FileName
'App.Path
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
Image1.Picture = LoadPicture(App.Path & "\..\photo\用户头像\moren.jpg")
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

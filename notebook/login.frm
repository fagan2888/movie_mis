VERSION 5.00
Begin VB.Form login 
   Caption         =   "login"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   14715
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "注册"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "检查密码"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   9
      Text            =   "admin"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Text            =   "admin"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "忘记密码"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "管理员"
      Height          =   615
      Left            =   7680
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "用户登录"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "还没有账号？快去注册试试吧~"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "密码："
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "用户名："
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "真正的说明：普通用户，admin，密码admin，并没有使用期限因为不会设置。。。管理员用户：guanli001,密码guanli001,希望老师使用愉快~"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   6240
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   $"login.frx":0000
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim n As Integer
Private Sub Command3_Click()
 b = MsgBox("不要走嘛，要不再看看？", vbYesNo)
 If b = vbNo Then shouye.Show
End Sub

Private Sub Command5_Click()
find.Show
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.PasswordChar = ""
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.PasswordChar = "*"
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
 MsgBox "用户名不能为空，别着急"
 Text1.SetFocus
 Exit Sub
End If
If Text2.Text = "" Then
 MsgBox "密码不能为空，别着急"
 Text2.Text = ""
 Exit Sub
End If
rs1.CursorLocation = adUseClient
rs1.Open "select  *  from user where userid='" & Text1.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
If rs1.RecordCount = 0 Then
   MsgBox "好像不太对，请注册或重新输入用户名"
   n = n + 1
         If n = 3 Then
              MsgBox "您输入的错误次数已经达到3次，请检查大小写或者再回忆一下，稍后在进行登陆"
              End
         End If
    Text1.Text = ""
    Text1.SetFocus
    rs1.Close
    Set rs1 = Nothing
    Exit Sub
Else
    If Trim(Text2.Text) <> rs1.Fields("userpassword") Then
         MsgBox "您输入的帐号或密码错误，请重新输入"
         Text2.Text = ""
         Text2.SetFocus
         rsclothes.Close
         Set rsclothes = Nothing
         Exit Sub
         n = n + 1
         If n = 3 Then
             MsgBox "您输入的帐号或密码错误次数已经达到3次，要不要检查一下密码？忘记密码可以找回哦，过会再来吧！"
             rs1.Close
             Set rs1 = Nothing
             End
         End If
     End If
End If

rs1.Close
Set rs1 = Nothing
MsgBox "登陆成功，快来体验吧！"
uid = Text1.Text
movie.Show
End Sub

Private Sub Command6_Click()
zhuce.Show
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then
 MsgBox "用户名不能为空，别着急"
 Text1.SetFocus
 Exit Sub
End If
If Text2.Text = "" Then
 MsgBox "密码不能为空，别着急"
 Text2.Text = ""
 Exit Sub
End If
rs1.CursorLocation = adUseClient
rs1.Open "select  *  from management where userid='" & Text1.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
If rs1.RecordCount = 0 Then
   MsgBox "好像不太对，请注册或重新输入用户名"
   n = n + 1
         If n = 3 Then
              MsgBox "您输入的错误次数已经达到3次，请检查大小写或者再回忆一下，稍后在进行登陆"
              End
         End If
    Text1.Text = ""
    Text1.SetFocus
    rs1.Close
    Set rs1 = Nothing
    Exit Sub
Else
    If Trim(Text2.Text) <> rs1.Fields("userpassword") Then
         MsgBox "您输入的帐号或密码错误，请重新输入"
         Text2.Text = ""
         Text2.SetFocus
         rs1.Close
         Set rsclothes = Nothing
         Exit Sub
         n = n + 1
         If n = 3 Then
             MsgBox "您输入的帐号或密码错误次数已经达到3次，要不要检查一下密码？忘记密码可以找回哦，过会再来吧！"
             rs1.Close
             Set rs1 = Nothing
             End
         End If
     End If
End If

rs1.Close
Set rs1 = Nothing
MsgBox "登陆成功，快来体验吧！"
uid1 = Text1.Text
guanliyuan.Show
End Sub


VERSION 5.00
Begin VB.Form chakanxinxi 
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   15045
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "返回"
      Height          =   495
      Left            =   8160
      TabIndex        =   23
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "末一条"
      Height          =   495
      Left            =   6600
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "后一条"
      Height          =   495
      Left            =   5160
      TabIndex        =   21
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "前一条"
      Height          =   495
      Left            =   3720
      TabIndex        =   20
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一条"
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   3
      Left            =   3960
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   0
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   3960
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   6840
      TabIndex        =   0
      Top             =   4080
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "密码"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "联系电话"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "电子邮箱"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "类型"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "性别"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   13
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户偏好"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   12
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "生日"
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "头像"
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "个人介绍"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "chakanxinxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs6 As New ADODB.Recordset
Dim i As Integer

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

VERSION 5.00
Begin VB.Form shouye 
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   Picture         =   "shouye.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   15120
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   8760
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "注册"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登陆"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "电影分享"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   5880
      TabIndex        =   0
      Top             =   1800
      Width           =   3495
   End
End
Attribute VB_Name = "shouye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
login.Show
End Sub

Private Sub Command2_Click()
zhuce.Show
End Sub

Private Sub Command3_Click()
jieshao.Show
End Sub

Private Sub Command4_Click()
 b = MsgBox("不要走嘛，要不再看看？", vbYesNo)
 If b = vbNo Then End
End Sub

Private Sub Form_Load()
Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

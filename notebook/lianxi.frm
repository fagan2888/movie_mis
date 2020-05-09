VERSION 5.00
Begin VB.Form lianxi 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   10560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label5 
      Caption         =   "返回"
      Height          =   255
      Left            =   8400
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "联系电话：15600611363"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "电子邮箱：869688716@qq.com"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "联系地址：清华东路十七号中国农业          大学东校区一号公寓B座"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "茫茫人海中遇到您，是我们的小幸运~  有什么可以帮到您的，请尽管联系我们！"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   6135
   End
End
Attribute VB_Name = "lianxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub


Private Sub Label5_Click()
 b = MsgBox("不要走嘛，要不再看看？", vbYesNo)
 If b = vbNo Then movie.Show
End Sub

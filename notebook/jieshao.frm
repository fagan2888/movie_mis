VERSION 5.00
Begin VB.Form jieshao 
   Caption         =   "介绍"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   15240
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "功能介绍："
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎来到电影分享的世界"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "jieshao"
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



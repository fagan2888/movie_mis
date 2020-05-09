VERSION 5.00
Begin VB.Form moviechakan 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   15090
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "返回"
      Height          =   615
      Left            =   11400
      TabIndex        =   19
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "管理此电影"
      Height          =   615
      Left            =   9840
      TabIndex        =   18
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "末一条"
      Height          =   615
      Left            =   8280
      TabIndex        =   17
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一条"
      Height          =   615
      Left            =   6840
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一条"
      Height          =   615
      Left            =   5160
      TabIndex        =   15
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一条"
      Height          =   615
      Left            =   3240
      TabIndex        =   14
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "moviechakan.frx":0000
         Top             =   3720
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Text            =   "Text7"
         Top             =   4920
         Width           =   2775
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Text            =   "Text8"
         Top             =   5520
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "人观看"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   13
         Top             =   3120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   3255
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "人推荐"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "评分"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "时长"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "导演"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "主演"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   5520
         Width           =   615
      End
   End
End
Attribute VB_Name = "moviechakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs4 As New ADODB.Recordset

Private Sub Command1_Click()
rs4.MoveFirst
Call viewdata
End Sub

Private Sub Command2_Click()
rs4.MovePrevious
If rs4.BOF Then rs4.MoveFirst
Call viewdata
End Sub

Private Sub Command3_Click()
rs4.MoveNext
If rs4.EOF Then rs4.MoveLast
Call viewdata
End Sub

Private Sub Command4_Click()
rs4.MoveLast
Call viewdata
End Sub

Private Sub Command5_Click()
mnum1 = rs4.Fields("movnum")
movieguanli.Show
End Sub

Private Sub Command6_Click()
guanliyuan.Show
End Sub

Private Sub Form_Load() '显示窗体的时候显示影片信息
biaoji = True
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
n = 1
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select movnum,movname,movphoto,movlove,movlook,movdirector,movactor,movgrade,movtime,movbir from movie", cnmovie, adOpenDynamic, adLockOptimistic
rs4.MoveFirst
Call viewdata
End Sub
Private Sub Form_Resize() '背景图片大小随窗体变动
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

Private Sub viewdata()
Image1.Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs4.Fields("movphoto"))
Text2 = rs4.Fields("movlove")
Text3 = rs4.Fields("movlook")
Text4 = rs4.Fields("movnum") + rs4.Fields("movname")
Text5 = rs4.Fields("movgrade")
Text6 = rs4.Fields("movtime")
Text7 = rs4.Fields("movdirector")
Text8 = rs4.Fields("movactor")
End Sub

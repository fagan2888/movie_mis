VERSION 5.00
Begin VB.Form movieshow 
   Caption         =   "电影精选"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14985
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "movieshow.frx":0000
      Left            =   9720
      List            =   "movieshow.frx":0013
      TabIndex        =   50
      Text            =   "剧情"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   2
      Left            =   9840
      TabIndex        =   36
      Top             =   1440
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
         Index           =   2
         Left            =   120
         TabIndex        =   43
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
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   42
         Text            =   "movieshow.frx":0035
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
         Index           =   2
         Left            =   720
         TabIndex        =   41
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
         Index           =   2
         Left            =   2880
         TabIndex        =   40
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
         Index           =   2
         Left            =   2040
         TabIndex        =   39
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
         Index           =   2
         Left            =   840
         TabIndex        =   38
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
         Index           =   2
         Left            =   840
         TabIndex        =   37
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
         Index           =   2
         Left            =   3000
         TabIndex        =   49
         Top             =   3120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   3255
         Index           =   2
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
         Index           =   2
         Left            =   1200
         TabIndex        =   48
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
         Index           =   2
         Left            =   0
         TabIndex        =   47
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
         Index           =   2
         Left            =   2160
         TabIndex        =   46
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
         Index           =   2
         Left            =   0
         TabIndex        =   45
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
         Index           =   2
         Left            =   0
         TabIndex        =   44
         Top             =   5520
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   1
      Left            =   5040
      TabIndex        =   22
      Top             =   1440
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
         Index           =   1
         Left            =   120
         TabIndex        =   29
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
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "movieshow.frx":003B
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
         Index           =   1
         Left            =   720
         TabIndex        =   27
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
         Index           =   1
         Left            =   2880
         TabIndex        =   26
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
         Index           =   1
         Left            =   2040
         TabIndex        =   25
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
         Index           =   1
         Left            =   840
         TabIndex        =   24
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
         Index           =   1
         Left            =   840
         TabIndex        =   23
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
         Index           =   1
         Left            =   3000
         TabIndex        =   35
         Top             =   3120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   3255
         Index           =   1
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
         Index           =   1
         Left            =   1200
         TabIndex        =   34
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
         Index           =   1
         Left            =   0
         TabIndex        =   33
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
         Index           =   1
         Left            =   2160
         TabIndex        =   32
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
         Index           =   1
         Left            =   0
         TabIndex        =   31
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
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Top             =   5520
         Width           =   615
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "返回"
      Height          =   375
      Left            =   8400
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "下一页"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "上一页"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
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
         Index           =   0
         Left            =   840
         TabIndex        =   19
         Text            =   "Text8"
         Top             =   5520
         Width           =   2775
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
         Index           =   0
         Left            =   840
         TabIndex        =   18
         Text            =   "Text7"
         Top             =   4920
         Width           =   2775
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
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   3120
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
         Index           =   0
         Left            =   2880
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   4440
         Width           =   975
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
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   4440
         Width           =   975
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
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "movieshow.frx":0041
         Top             =   3720
         Width           =   3615
      End
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
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   3120
         Width           =   1095
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
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   5520
         Width           =   615
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
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   4920
         Width           =   735
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
         Index           =   0
         Left            =   2160
         TabIndex        =   17
         Top             =   4440
         Width           =   735
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
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   4440
         Width           =   615
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
         TabIndex        =   15
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   3255
         Index           =   0
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
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
         Index           =   0
         Left            =   3000
         TabIndex        =   11
         Top             =   3120
         Width           =   855
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "分类查看"
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "按观看人数排序"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "按评分排序"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "请输入电影名"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "搜索"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "movieshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs4 As New ADODB.Recordset
Dim n As Integer '当前页
Dim m As Integer '记录总数
Dim page As Integer '页数
Dim biaoji As Boolean

Private Sub Form_Load() '显示窗体的时候显示影片信息
biaoji = True
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
n = 1
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select movnum,movname,movphoto,movlove,movlook,movdirector,movactor,movgrade,movtime,movbir from movie", cnmovie, adOpenDynamic, adLockOptimistic
m = rs4.RecordCount
If m Mod 3 <> 0 Then
  page = Int(m / 3) + 1
Else
  page = Int(m / 3)
  Command5.Enabled = False
If page > 1 Then Command6.Enabled = True
End If
Call dangqianye
End Sub
Private Sub Form_Resize() '背景图片大小随窗体变动
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

Private Sub dangqianye()
If biaoji = True Then
      If n < page Then
            For i = 0 To 2
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs4.Fields("movphoto"))
                Text2(i) = rs4.Fields("movlove")
                Text3(i) = rs4.Fields("movlook")
                Text4(i) = rs4.Fields("movnum") + rs4.Fields("movname")
                Text5(i) = rs4.Fields("movgrade")
                Text6(i) = rs4.Fields("movtime")
                Text7(i) = rs4.Fields("movdirector")
                Text8(i) = rs4.Fields("movactor")
                rs4.MoveNext
           Next i
        Else
            For i = 0 To rs4.RecordCount - (page - 1) * 3 - 1
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs4.Fields("movphoto"))
                Text2(i) = rs4.Fields("movlove")
                Text3(i) = rs4.Fields("movlook")
                Text4(i) = rs4.Fields("movnum") + rs4.Fields("movname")
                Text5(i) = rs4.Fields("movgrade")
                Text6(i) = rs4.Fields("movtime")
                Text7(i) = rs4.Fields("movdirector")
                Text8(i) = rs4.Fields("movactor")
                rs4.MoveNext
                If rs4.EOF Then rs4.MoveLast
            Next i
            For i = rs4.RecordCount - (page - 1) * 3 To 2
                Image1(i).Picture = Nothing
                Text2(i) = ""
                Text3(i) = ""
                Text4(i) = ""
                Text5(i) = ""
                Text6(i) = ""
                Text7(i) = ""
                Text8(i) = ""
                Label1(i).Caption = ""
                Label2(i).Caption = ""
                Label3(i).Caption = ""
                Label4(i).Caption = ""
                Label5(i).Caption = ""
                Label6(i).Caption = ""
            Next i
       End If
Else
     If n = page - 1 Then
          rs4.Move (page - 2) * 3 - rs4.RecordCount
          If rs4.BOF Then rs4.MoveFirst
          For i = 0 To 2
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs4.Fields("movphoto"))
                Text2(i) = rs4.Fields("movlove")
                Text3(i) = rs4.Fields("movlook")
                Text4(i) = rs4.Fields("movnum") + rs4.Fields("movname")
                Text5(i) = rs4.Fields("movgrade")
                Text6(i) = rs4.Fields("movtime")
                Text7(i) = rs4.Fields("movdirector")
                Text8(i) = rs4.Fields("movactor")
                rs4.MoveNext
          Next i
     Else
         rs4.Move -6
        If rs4.BOF Then rs4.MoveFirst
        For i = 0 To 2
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs4.Fields("movphoto"))
                Text2(i) = rs4.Fields("movlove")
                Text3(i) = rs4.Fields("movlook")
                Text4(i) = rs4.Fields("movnum") + rs4.Fields("movname")
                Text5(i) = rs4.Fields("movgrade")
                Text6(i) = rs4.Fields("movtime")
                Text7(i) = rs4.Fields("movdirector")
                Text8(i) = rs4.Fields("movactor")
                rs4.MoveNext
        Next i
    End If
End If
End Sub

Private Sub Command2_Click()
n = 1
For i = 0 To 2
Image1(i).Visible = True
Text2(i).Visible = True
Text3(i).Visible = True
Text4(i).Visible = True
Text5(i).Visible = True
Text6(i).Visible = True
Text7(i).Visible = True
Text8(i).Visible = True
Label1(i).Visible = True
Label2(i).Visible = True
Label3(i).Visible = True
Label4(i).Visible = True
Label5(i).Visible = True
Label6(i).Visible = True
Next i
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select movnum,movname,movphoto,movlove,movlook,movdirector,movactor,movgrade,movtime,movbir from movie order by movgrade asc", cnmovie, adOpenDynamic, adLockOptimistic
m = rs4.RecordCount
If Im Mod 3 <> 0 Then
  page = Int(m / 3) + 1
Else
  page = Int(m / 3)
  Command5.Enabled = False
If page > 1 Then Command6.Enabled = True
End If
Call dangqianye
End Sub
Private Sub Command3_Click()
n = 1
For i = 0 To 2
Image1(i).Visible = True
Text2(i).Visible = True
Text3(i).Visible = True
Text4(i).Visible = True
Text5(i).Visible = True
Text6(i).Visible = True
Text7(i).Visible = True
Text8(i).Visible = True
Label1(i).Visible = True
Label2(i).Visible = True
Label3(i).Visible = True
Label4(i).Visible = True
Label5(i).Visible = True
Label6(i).Visible = True
Next i
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select movnum,movname,movphoto,movlove,movlook,movdirector,movactor,movgrade,movtime,movbir from movie order by movlook asc", cnmovie, adOpenDynamic, adLockOptimistic
m = rs4.RecordCount
If m Mod 3 <> 0 Then
  page = Int(m / 3) + 1
Else
  page = Int(m / 3)
  Command5.Enabled = False
If page > 1 Then Command6.Enabled = True
End If
Call dangqianye
End Sub
Private Sub Command4_Click()
n = 1
For i = 0 To 2
Image1(i).Visible = True
Text2(i).Visible = True
Text3(i).Visible = True
Text4(i).Visible = True
Text5(i).Visible = True
Text6(i).Visible = True
Text7(i).Visible = True
Text8(i).Visible = True
Label1(i).Visible = True
Label2(i).Visible = True
Label3(i).Visible = True
Label4(i).Visible = True
Label5(i).Visible = True
Label6(i).Visible = True
Next i
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select movnum,movname,movphoto,movlove,movlook,movdirector,movactor,movgrade,movtime,movbir from movie where movtype='" & Combo1.Text & " '", cnmovie, adOpenDynamic, adLockOptimistic
m = rs4.RecordCount
If m Mod 3 <> 0 Then
  page = Int(m / 3) + 1
Else
  page = Int(m / 3)
  Command5.Enabled = False
If page > 1 Then Command6.Enabled = True
End If
Call dangqianye
End Sub
Private Sub Command5_Click()
biaoji = False
Command6.Enabled = True
If n > 1 Then
n = n - 1
End If
If n = 1 Then
Command5.Enabled = False
End If

Call dangqianye

End Sub
Private Sub Command6_Click()
biaoji = True
Command5.Enabled = True
n = n + 1
If n = page Then
Command6.Enabled = False
End If

Call dangqianye

End Sub
Private Sub Command7_Click()
b = MsgBox("不要走嘛，要不再看看？", vbYesNo)
 If b = vbNo Then movie.Show
End Sub

Private Sub Frame1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
mnum = Left(Text4(Index), 6)
moviexiangqing.Show
End Sub

Private Sub Image1_Click(Index As Integer)
mnum = Left(Text4(Index), 6)
moviexiangqing.Show
End Sub

Private Sub Label7_Click()
n = 1
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select movnum,movname,movphoto,movlove,movlook,movdirector,movactor,movgrade,movtime,movbir from movie where movname='" & Text1.Text & " '", cnmovie, adOpenDynamic, adLockOptimistic
m = rs4.RecordCount
If rs4.EOF Then
MsgBox "您所搜索的电影不存在哦，谢谢您的使用"
Else
page = 1
Image1(0).Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs4.Fields("movphoto"))
Text2(0) = rs4.Fields("movlove")
Text3(0) = rs4.Fields("movlook")
Text4(0) = rs4.Fields("movnum") + rs4.Fields("movname")
Text5(0) = rs4.Fields("movgrade")
Text6(0) = rs4.Fields("movtime")
Text7(0) = rs4.Fields("movdirector")
Text8(0) = rs4.Fields("movactor")
For i = 1 To 2
Image1(i).Visible = False
Text2(i).Visible = False
Text3(i).Visible = False
Text4(i).Visible = False
Text5(i).Visible = False
Text6(i).Visible = False
Text7(i).Visible = False
Text8(i).Visible = False
Label1(i).Visible = False
Label2(i).Visible = False
Label3(i).Visible = False
Label4(i).Visible = False
Label5(i).Visible = False
Label6(i).Visible = False
Next i
End If
rs4.Close
End Sub

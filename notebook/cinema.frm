VERSION 5.00
Begin VB.Form cinema 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   15210
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   375
      Left            =   10200
      TabIndex        =   40
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   2
      Left            =   10560
      TabIndex        =   28
      Top             =   1080
      Width           =   4215
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
         Left            =   1080
         TabIndex        =   34
         Text            =   "Text8"
         Top             =   5160
         Width           =   2415
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
         Left            =   2760
         TabIndex        =   33
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
         Index           =   2
         Left            =   720
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   4560
         Width           =   3375
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
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   4080
         Width           =   3375
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
         Height          =   375
         Index           =   2
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "cinema.frx":0000
         Top             =   3600
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
         Index           =   2
         Left            =   600
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   615
         Index           =   2
         Left            =   1920
         TabIndex        =   35
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "人均消费"
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
         TabIndex        =   39
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "电话"
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
         Left            =   0
         TabIndex        =   38
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "地址"
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
         TabIndex        =   37
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "城市"
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
         Left            =   0
         TabIndex        =   36
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   3255
         Index           =   2
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   1
      Left            =   5520
      TabIndex        =   16
      Top             =   1080
      Width           =   4215
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
         Left            =   1080
         TabIndex        =   22
         Text            =   "Text8"
         Top             =   5160
         Width           =   2415
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
         Left            =   2760
         TabIndex        =   21
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
         Index           =   1
         Left            =   720
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   4560
         Width           =   3375
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
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   4080
         Width           =   3375
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
         Height          =   375
         Index           =   1
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "cinema.frx":0006
         Top             =   3600
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
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   615
         Index           =   1
         Left            =   1920
         TabIndex        =   23
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
      Begin VB.Label Label6 
         Caption         =   "人均消费"
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
         TabIndex        =   27
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "电话"
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
         Left            =   0
         TabIndex        =   26
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "地址"
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
         TabIndex        =   25
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "城市"
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
         Left            =   0
         TabIndex        =   24
         Top             =   3120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1200
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
         Index           =   0
         Left            =   600
         TabIndex        =   10
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
         Height          =   375
         Index           =   0
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "cinema.frx":000C
         Top             =   3600
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
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   4080
         Width           =   3375
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
         Left            =   720
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   4560
         Width           =   3375
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
         Left            =   2760
         TabIndex        =   6
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
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Text            =   "Text8"
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label1 
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
         Height          =   615
         Index           =   0
         Left            =   1920
         TabIndex        =   15
         Top             =   3120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   3255
         Index           =   0
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "城市"
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
         Left            =   0
         TabIndex        =   14
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "地址"
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
         TabIndex        =   13
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "电话"
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
         Left            =   0
         TabIndex        =   12
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "人均消费"
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
         TabIndex        =   11
         Top             =   5160
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一页"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一页"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "搜索"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "输入你所在的城市"
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "cinema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs4 As New ADODB.Recordset
Dim n As Integer '当前页
Dim m As Integer '记录总数
Dim page As Integer '页数
Dim biaoji As Boolean
Private Sub Command1_Click()
n = 1
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select * from cinema where cincity='" & Text1.Text & " '", cnmovie, adOpenDynamic, adLockOptimistic
m = rs4.RecordCount
If m Mod 3 <> 0 Then
  page = Int(m / 3) + 1
Else
  page = Int(m / 3)
  Command2.Enabled = False
If page > 1 Then Command3.Enabled = True
End If
Call dangqianye
End Sub

Private Sub Command3_Click()
biaoji = True
Command2.Enabled = True
n = n + 1
If n = page Then
Command3.Enabled = False
End If

Call dangqianye

End Sub

Private Sub Command2_Click()
biaoji = False
Command3.Enabled = True
If n > 1 Then
n = n - 1
End If
If n = 1 Then
Command2.Enabled = False
End If

Call dangqianye

End Sub

Private Sub Command4_Click()
movie.Show
End Sub

Private Sub Form_Load() '显示窗体的时候显示影院信息
biaoji = True
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
n = 1
If rs4.State = adStateOpen Then rs4.Close
rs4.CursorLocation = adUseClient
rs4.PageSize = 3
rs4.Open "select * from cinema", cnmovie, adOpenDynamic, adLockOptimistic
m = rs4.RecordCount
If m Mod 3 <> 0 Then
  page = Int(m / 3) + 1
Else
  page = Int(m / 3)
  Command2.Enabled = False
If page > 1 Then Command3.Enabled = True
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
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影院\" + rs4.Fields("cinphoto"))
                Text2(i) = rs4.Fields("cincity")
                Text3(i) = rs4.Fields("cingrade")
                Text4(i) = rs4.Fields("cinid") + rs4.Fields("cinnam")
                Text5(i) = rs4.Fields("cinloc")
                Text6(i) = rs4.Fields("cinphone")
                Text7(i) = rs4.Fields("cinprice")
                rs4.MoveNext
           Next i
        Else
            For i = 0 To rs4.RecordCount - (page - 1) * 3 - 1
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影院\" + rs4.Fields("cinphoto"))
                Text2(i) = rs4.Fields("cincity")
                Text3(i) = rs4.Fields("cingrade")
                Text4(i) = rs4.Fields("cinid") + rs4.Fields("cinnam")
                Text5(i) = rs4.Fields("cinloc")
                Text6(i) = rs4.Fields("cinphone")
                Text7(i) = rs4.Fields("cinprice")
                rs4.MoveNext
                If rs4.EOF Then rs4.MoveLast
            Next i
            For i = rs4.RecordCount - (page - 1) * 3 To 2
                Image1(i).Picture = Nothing
                Text2(i).Text = ""
                Text3(i).Text = ""
                Text4(i).Text = ""
                Text5(i).Text = ""
                Text6(i).Text = ""
                Text7(i).Text = ""
                Label1(i).Caption = ""
                Label2(i).Caption = ""
                Label3(i).Caption = ""
                Label4(i).Caption = ""
                Label6(i).Caption = ""
            Next i
       End If
Else
     If n = page - 1 Then
          rs4.Move (page - 2) * 3 - rs4.RecordCount
          If rs4.BOF Then rs4.MoveFirst
          For i = 0 To 2
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影院\" + rs4.Fields("cinphoto"))
                Text2(i) = rs4.Fields("cincity")
                Text3(i) = rs4.Fields("cingrade")
                Text4(i) = rs4.Fields("cinid") + rs4.Fields("cinnam")
                Text5(i) = rs4.Fields("cinloc")
                Text6(i) = rs4.Fields("cinphone")
                Text7(i) = rs4.Fields("cinprice")
                rs4.MoveNext
          Next i
     Else
         rs4.Move -6
        If rs4.BOF Then rs4.MoveFirst
        For i = 0 To 2
                Image1(i).Picture = LoadPicture(App.Path + "\..\photo\电影院\" + rs4.Fields("cinphoto"))
                Text2(i) = rs4.Fields("cincity")
                Text3(i) = rs4.Fields("cingrade")
                Text4(i) = rs4.Fields("cinid") + rs4.Fields("cinnam")
                Text5(i) = rs4.Fields("cinloc")
                Text6(i) = rs4.Fields("cinphone")
                Text7(i) = rs4.Fields("cinprice")
                rs4.MoveNext
        Next i
    End If
End If
End Sub

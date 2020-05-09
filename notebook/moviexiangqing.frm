VERSION 5.00
Begin VB.Form moviexiangqing 
   Caption         =   "Form2"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15135
   LinkTopic       =   "Form2"
   ScaleHeight     =   7875
   ScaleWidth      =   15135
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   10200
      TabIndex        =   28
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "点播"
      Height          =   375
      Left            =   10200
      TabIndex        =   27
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   9
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   6600
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   10
      Left            =   7080
      TabIndex        =   24
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   7080
      TabIndex        =   22
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   7080
      TabIndex        =   21
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   7080
      TabIndex        =   20
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   19
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   18
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   17
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   16
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   15
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "电影简介"
      Height          =   255
      Index           =   10
      Left            =   1560
      TabIndex        =   25
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "推荐"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "看过"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "时长"
      Height          =   255
      Index           =   9
      Left            =   6000
      TabIndex        =   9
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "获奖信息"
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   8
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "国家"
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "上映日期"
      Height          =   255
      Index           =   6
      Left            =   6000
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "类型"
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "主演"
      Height          =   255
      Index           =   4
      Left            =   6000
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "编剧"
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "导演"
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "电影名"
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "电影编号"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "moviexiangqing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs12 As New ADODB.Recordset
Dim rs13 As New ADODB.Recordset
Dim ty As String

Private Sub Command1_Click()
rs13.CursorLocation = adUseClient
rs13.Open "select  *  from user where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
ty = rs13.Fields("usertype")
rs13.Close
rs13.Open "select  *  from dianbo where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
If ty = "VIP" Then
   MsgBox "会员可免费观看此片"
   m = 0
Else
      If rs12.Fields("movlocation") = "中国" Then
         Select Case rs12.Fields("movtime")
         Case Is <= 100
            m = 0.5
            n = MsgBox("一百分钟以下国产电影的资费为0.5元，是否支付？", vbYesNo)
            If n = vbYes Then MsgBox "支付成功"
         Case Is <= 200
            m = 1
            n = MsgBox("一百分钟以上，两百以下国产电影的资费为1元，是否支付？", vbYesNo)
            If n = vbYes Then MsgBox "支付成功"
        Case Else
            m = 2
            n = MsgBox("二百分钟以上国产电影的资费为2元，是否支付？", vbYesNo)
            If n = vbYes Then MsgBox "支付成功"
        End Select
      Else
        Select Case rs12.Fields("movtime")
         Case Is <= 100
            m = 1
            n = MsgBox("一百分钟以下国产电影的资费为1元，是否支付？", vbYesNo)
            If n = vbYes Then MsgBox "支付成功"
         Case Is <= 200
            m = 2
            n = MsgBox("一百分钟以上，两百以下国产电影的资费为2元，是否支付？", vbYesNo)
            If n = vbYes Then MsgBox "支付成功"
        Case Else
            m = 3
            n = MsgBox("一百分钟以上，两百以下国产电影的资费为3元，是否支付？", vbYesNo)
            If n = vbYes Then MsgBox "支付成功"
        End Select
      End If
End If
 c = "insert into dianbo(movnum,userid,price)  values('" & mnum & "','" & uid & "'," & m & ")"
 cnmovie.Execute c
 list.Show
End Sub

Private Sub Command2_Click()
movie.Show
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
rs12.CursorLocation = adUseClient
rs12.Open "select  *  from movie where movnum='" & mnum & "'", cnmovie, adOpenDynamic, adLockOptimistic
For i = 0 To 10
 Text1(i).Enabled = False
 Text1(i).Text = rs12.Fields(i)
Next i
Image1.Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs12.Fields("movphoto"))
Text4.Text = rs12.Fields("movgrade")
Text2.Text = rs12.Fields("movlook")
Text3.Text = rs12.Fields("movlove")

End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

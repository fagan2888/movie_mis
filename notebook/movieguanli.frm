VERSION 5.00
Begin VB.Form movieguanli 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   15225
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "取消"
      Height          =   375
      Left            =   9840
      TabIndex        =   32
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   10
      Left            =   6720
      TabIndex        =   30
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   375
      Left            =   9840
      TabIndex        =   29
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "保存"
      Height          =   375
      Left            =   9840
      TabIndex        =   28
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      Height          =   375
      Left            =   9840
      TabIndex        =   27
      Top             =   3960
      Width           =   1215
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
      Left            =   0
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   11
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   10
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   6720
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   6720
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   6720
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   9
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6240
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "时长"
      Height          =   375
      Left            =   5640
      TabIndex        =   31
      Top             =   5400
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "电影编号"
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   26
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "电影名"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   25
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "导演"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   24
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "编剧"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   23
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "主演"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   22
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "类型"
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   21
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "上映日期"
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   20
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "国家"
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   19
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "获奖信息"
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   18
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "看过"
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "推荐"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "电影简介"
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   15
      Top             =   5880
      Width           =   735
   End
End
Attribute VB_Name = "movieguanli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs12 As New ADODB.Recordset
Dim edtif As Integer
Dim a(14) As String
Private Sub Command1_Click()
For i = 0 To 10
 Text1(i).Enabled = True
 Text1(i).Text = ""
Next i
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Image1.Enabled = True
Text4.Text = ""
Text2.Text = ""
Text3.Text = ""
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Command6.Enabled = True
edtif = 0
End Sub

Private Sub Command2_Click()
For i = 0 To 10
 Text1(i).Enabled = True
Next i
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Image1.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Command6.Enabled = True
edtif = 1
End Sub

Private Sub Command3_Click()
b = MsgBox("是否要删除该记录？", vbYesNo)
 If b = vbYes Then
  a = "delete from movie where movnum='"
  a = a + Text1(0).Text + "'"
  cnmovie.Execute a
  rs12.Close
  MsgBox "删除成功！"
 End If
End Sub

Private Sub Command4_Click()
Select Case editf
 Case 0
  If Text1(0).Text = "" Then
   MsgBox "电影编号不能为空！"
   Text1(0).SetFocus
    Exit Sub
  End If
 Dim rstemp As New ADODB.Recordset
 Dim strtemp As String
strtemp = "select * from movie where movnum='" & (Text1(0).Text) & "';"
rstemp.Open strtemp, cnmovie, adOpenDynamic, adLockPessimistic
   If Not rstemp.EOF Then
     MsgBox "电影编号不唯一，重新输入！"
     Text1(0).Text = ""
     Text1(0).SetFocus
   rstemp.Close
   Exit Sub
   End If
rstemp.Close
For i = 0 To 10
  Text1(i).Enabled = False
  a(i) = Text1(i).Text
Next i
a(11) = Text4.Text
a(12) = Text2.Text
a(13) = Text3.Text
a(14) = Text1(0).Text + ".jpg"
 a = "insert into movie values('" & a(0) & "','" & a(1) & "','" & a(2) & "','" & a(3) & "','" & a(4) & "','" & a(5) & "','" & a(6) & "','" & a(7) & "','" & a(8) & "','" & a(9) & "','" & a(10) & "','" & a(14) & "','" & a(11) & "','" & a(12) & "','" & a(13) & "')"
  cnmovie.Execute a
  MsgBox "保存成功！"
Case 1
  For i = 0 To 10
  Text1(i).Enabled = False
  a(i) = Text1(i).Text
Next i
a(11) = Text4.Text
a(12) = Text2.Text
a(13) = Text3.Text
a(14) = Text1(0).Text + ".jpg"
  a = "update movie set movnum='" & a(0) & "',movname='" & a(1) & "',movdirector='" & a(2) & "',movauthor='" & a(3) & "',movactor='" & a(4) & "',movtype='" & a(5) & "',movbir='" & a(6) & "',movlocation='" & a(7) & "',movawards='" & a(8) & "',movresume='" & a(9) & "',movtime='" & a(10) & "',movphoto='" & a(14) & "',movgrade='" & a(11) & "',movlook='" & a(12) & "',movlove='" & a(13) & "' where movnum='" + Text1(0).Text + "'"
  cnmovie.Execute a
  MsgBox "修改完毕！"
  Text1(0).Locked = False
  editf = 0
  End Select
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  Command4.Enabled = False
  Command6.Enabled = False
End Sub

Private Sub Command5_Click()
moviechakan.Show
End Sub

Private Sub Command6_Click()
rs12.CancelUpdate
For i = 0 To 10
 Text1(i).Enabled = False
 Text1(i).Text = rs12.Fields(i)
Next i
Image1.Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs12.Fields("movphoto"))
Text4.Text = rs12.Fields("movgrade")
Text2.Text = rs12.Fields("movlook")
Text3.Text = rs12.Fields("movlove")
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  Command4.Enabled = False
  Command6.Enabled = False
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
rs12.CursorLocation = adUseClient
rs12.Open "select  *  from movie where movnum='" & mnum1 & "'", cnmovie, adOpenDynamic, adLockOptimistic
For i = 0 To 10
 Text1(i).Enabled = False
 Text1(i).Text = rs12.Fields(i)
Next i
Image1.Picture = LoadPicture(App.Path + "\..\photo\电影海报\" + rs12.Fields("movphoto"))
Text4.Text = rs12.Fields("movgrade")
Text2.Text = rs12.Fields("movlook")
Text3.Text = rs12.Fields("movlove")
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub



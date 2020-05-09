VERSION 5.00
Begin VB.Form yingping 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   15195
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "保存修改"
      Height          =   615
      Left            =   9000
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "修改影评"
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "查看影评"
      Height          =   615
      Left            =   5280
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      Height          =   615
      Left            =   9480
      TabIndex        =   10
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "提交影评"
      Height          =   615
      Left            =   6240
      TabIndex        =   9
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除该影评"
      Height          =   615
      Left            =   10680
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Text            =   "请只输入影片名或只输入影片精选页面影片编号，如美丽人生或010101"
      Top             =   3600
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "yingping.frx":0000
      Top             =   4440
      Width           =   6255
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5040
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   "影评内容"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "影片名"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "写影评"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "我的影评"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "yingping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs9 As New ADODB.Recordset
Dim d As String, e As Long

Private Sub Command5_Click() '修改影评
Text1.Enabled = True
End Sub

Private Sub Command6_Click()
If Text1.Text = "" Then
  MsgBox "影评为空，不能提交"
Else
  sql = "update review set review='" & Text1.Text & "' where revid=" & e & ""
  cnmovie.Execute sql
  MsgBox "修改完毕！"
  Text1.Enabled = False
End If
rs9.Close
Set rs9 = Nothing
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
rs9.CursorLocation = adUseClient
rs9.Open "select  revid  from review where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
If rs9.RecordCount = 0 Then
    MsgBox ("您还未发表过影评！")
Else
    rs9.MoveFirst
    For i = 0 To rs9.RecordCount - 1
      e = rs9.Fields("revid")
      Combo1.AddItem (e)
      rs9.MoveNext
    Next i
End If
rs9.Close
Set rs9 = Nothing
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

Private Sub Command2_Click() '提交影评
tijiaoyingping.Show
If Text2.Text = "" Then
  MsgBox "影评为空，不能提交"
Else
  rs9.CursorLocation = adUseClient
  rs9.Open "select * from movie where movname='" & Text3.Text & "'or movnum='" & Text3.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
  If rs9.BOF And rs9.EOF Then MsgBox "没有此电影，很抱歉哦~感谢您的分享"
  If rs9.RecordCount > 1 Then MsgBox "有多部此名字的电影哦，请在搜索页面关注电影编号,并只输入电影编号~"
  g = rs9.Fields("movnum")
  rs9.Close
  rs9.Open "select * from review where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
  a = uid
  b = Text2.Text
  c = "insert into review(review,userid,movnum) values('" & b & "','" & a & "','" & g & "')"
  cnmovie.Execute c
  rs9.Close
  rs9.Open "select * from review", cnmovie, adOpenDynamic, adLockOptimistic
  rs9.MoveLast
  d = rs9.Fields("revid")
  MsgBox "提交成功，该影评编号为" + d
  rs9.Close
  Set rs9 = Nothing
End If
End Sub

Private Sub Command3_Click()
yingping.Hide
movie.Show
End Sub

Private Sub Command1_Click() '删除影评
e = Val(Combo1.Text)
rs9.CursorLocation = adUseClient
rs9.Open "select  *  from review where revid=" & e & "", cnmovie, adOpenDynamic, adLockOptimistic
  h = MsgBox("是否要删除该记录？", vbYesNo)
 If h = vbYes Then
  k = "delete from review where revid=" & e & ""
  cnmovie.Execute k
  MsgBox "删除成功"
  rs9.Close
  Set rs9 = Nothing
 End If
End Sub

Private Sub Command4_Click() '查看影评
e = Val(Combo1.Text)
rs9.CursorLocation = adUseClient
rs9.Open "select  *  from review where revid=" & e & "", cnmovie, adOpenDynamic, adLockOptimistic
Text1.Text = rs9.Fields("review")
rs9.Close
Set rs9 = Nothing
End Sub


VERSION 5.00
Begin VB.Form yingpingchakan 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   15240
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "反对"
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "点赞"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "yingpingchakan.frx":0000
      Top             =   600
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "返回"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "最后一条"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一条"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一条"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一条"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "反对"
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "赞同"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "评论"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "yingpingchakan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs10 As New ADODB.Recordset
Dim rs11 As New ADODB.Recordset
Dim a As Integer, b As Integer, d As Long

Private Sub Command1_Click()
rs10.MoveFirst
Call viewdata
End Sub

Private Sub Command2_Click()
 rs10.MovePrevious
 If rs10.BOF Then rs10.MoveFirst
 Call viewdata
End Sub

Private Sub Command3_Click()
 rs10.MoveNext
 If rs10.EOF Then rs10.MoveLast
 Call viewdata
End Sub

Private Sub Command4_Click()
rs10.MoveFirst
Call viewdata
End Sub

Private Sub Command5_Click()
movie.Show
End Sub

Private Sub Command6_Click()
a = a + 1
Label2.Caption = a
rs11.CursorLocation = adUseClient
rs11.Open "select * from review where review='" & Text3.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
c = "update review set revlove =" & a & " where review ='" & Text3.Text & "' "
cnmovie.Execute c
End Sub

Private Sub Command7_Click()
b = b + 1
Label4.Caption = b
rs10.CursorLocation = adUseClient
rs11.Open "select * from review where review='" & Text3.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
c = "update review set revhate =" & b & " where review ='" & Text3.Text & "' "
cnmovie.Execute c
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
rs10.CursorLocation = adUseClient
rs10.Open "select distinct* from review,movie where movie.movnum=review.movnum", cnmovie, adOpenDynamic, adLockOptimistic
  If rs10.BOF And rs10.EOF Then
    MsgBox "表中无记录！"
  Else
   rs10.MoveFirst
   Call viewdata
  End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End Sub
Private Sub viewdata()
Text1.Text = rs10.Fields("userid")
Text2.Text = rs10.Fields("movname")
Text3.Text = rs10.Fields("review")
Label2.Caption = rs10.Fields("revlove")
Label4.Caption = rs10.Fields("revhate")
a = rs10.Fields("revlove")
b = rs10.Fields("revhate")
End Sub

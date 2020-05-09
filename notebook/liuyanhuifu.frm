VERSION 5.00
Begin VB.Form liuyanhuifu 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   11595
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "删除留言"
      Height          =   615
      Left            =   6000
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   615
      Left            =   8160
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查看回复"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   3360
      TabIndex        =   0
      Top             =   2880
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "我的留言"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "我收到的回复"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "liuyanhuifu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs8 As New ADODB.Recordset
Dim e As Long, f As Long
Private Sub Command1_Click()
e = Val(Combo1.Text)
rs8.CursorLocation = adUseClient
rs8.Open "select  *  from message where mesid=" & e & "", cnmovie, adOpenDynamic, adLockOptimistic
Text2.Text = ""
If rs8.Fields("mesyes") = True Then
         Text1.Text = rs8("message")
         f = rs8.Fields("huifuid")
         rs8.Close
         Set rs8 = Nothing
         rs8.CursorLocation = adUseClient
         rs8.Open "select  *  from huifu where huifuid= " & f & " ", cnmovie, adOpenDynamic, adLockOptimistic
         Text2.Text = rs8.Fields("huifu")
Else
         Text1.Text = rs8.Fields("message")
         Text2.Text = "您的问题还没有被回复，请耐心等待~"
End If
rs8.Close
Set rs8 = Nothing
End Sub

Private Sub Command2_Click()
movie.Show
End Sub

Private Sub Command3_Click()
rs8.CursorLocation = adUseClient
rs8.Open "select  *  from message where mesid=" & e & "", cnmovie, adOpenDynamic, adLockOptimistic
e = Val(Combo1.Text)
  b = MsgBox("是否要删除该记录？", vbYesNo)
 If b = vbYes Then
  a = "delete from message where mesid=" & e & ""
  cnmovie.Execute a
  rs8.Close
  rs8.Open "select  *  from message where mesid=" & e & "", cnmovie, adOpenDynamic, adLockOptimistic
     If rs8.RecordCount = 0 Then
    MsgBox ("您还未进行过提问！")
Else
    rs8.MoveFirst
    For i = 0 To rs8.RecordCount - 1
      e = rs8.Fields("mesid")
      Combo1.AddItem (e)
      rs8.MoveNext
    Next i
End If
rs8.Close
Set rs8 = Nothing
 End If
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
rs8.CursorLocation = adUseClient
rs8.Open "select  mesid  from message where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
If rs8.RecordCount = 0 Then
    MsgBox ("您还未进行过提问！")
Else
    rs8.MoveFirst
    For i = 0 To rs8.RecordCount - 1
      e = rs8.Fields("mesid")
      Combo1.AddItem (e)
      rs8.MoveNext
    Next i
End If
rs8.Close
Set rs8 = Nothing
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

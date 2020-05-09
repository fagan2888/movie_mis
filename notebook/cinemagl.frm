VERSION 5.00
Begin VB.Form cinemagl 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   15045
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "取消"
      Height          =   375
      Left            =   8880
      TabIndex        =   21
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "返回"
      Height          =   375
      Left            =   6960
      TabIndex        =   20
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "保存"
      Height          =   375
      Left            =   8880
      TabIndex        =   19
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "删除"
      Height          =   375
      Left            =   8880
      TabIndex        =   18
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "修改"
      Height          =   375
      Left            =   8880
      TabIndex        =   17
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "添加"
      Height          =   375
      Left            =   8880
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "末一条"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下一条"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "前一条"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一条"
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   5895
      Begin VB.TextBox Text7 
         BackColor       =   &H80000004&
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
         Left            =   4080
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
         Left            =   4080
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000004&
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
         Left            =   720
         TabIndex        =   4
         Top             =   4560
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000004&
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
         TabIndex        =   3
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000004&
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
         Left            =   4080
         TabIndex        =   1
         Top             =   120
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
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   720
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
         Height          =   495
         Index           =   0
         Left            =   3000
         TabIndex        =   11
         Top             =   2160
         Width           =   975
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
         TabIndex        =   10
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
         Index           =   0
         Left            =   0
         TabIndex        =   9
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
         Index           =   0
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   3255
         Left            =   240
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2775
      End
   End
End
Attribute VB_Name = "cinemagl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs4 As New ADODB.Recordset
Dim edtif As Integer
Dim a(6) As String

Private Sub Command10_Click()
rs4.CancelUpdate
  rs4.MoveFirst
  Call viewdata
  Image1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = False
Command10.Enabled = False
End Sub

Private Sub Command5_Click()
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Image1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = True
Command10.Enabled = True
edtif = 0
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = False
Command10.Enabled = False
End Sub

Private Sub Command6_Click()
Image1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = True
Command10.Enabled = True
edtif = 1
End Sub

Private Sub Command7_Click()
 b = MsgBox("是否要删除该记录？", vbYesNo)
 If b = vbYes Then
  a = "delete from cinema where cinid='"
  a = a + Text1.Text + "'"
  cnmovie.Execute a
  rs4.Close
  sql = "select * from cinema"
  rs4.Open sql, cnmovie, adOpenDynamic, adLockOptimistic
     If rs.BOF And rs.EOF Then
       MsgBox "表中无记录！"
     Else
       rs4.MoveFirst
     Call viewdata
   End If
 End If
End Sub

Private Sub Command8_Click()
Select Case editf
 Case 0
 a(0) = Text1.Text
 a(1) = Text2.Text
 a(2) = Text3.Text
 a(3) = Text4.Text
 a(4) = Text5.Text
 a(5) = Text6.Text
 a(6) = Text7.Text
 a = "insert into cinema(cinid,cincity,cingrade,cinnam,cinloc,cinphone,price)  values('" & a(0) & "','" & a(1) & "','" & a(2) & "','" & a(3) & "','" & a(4) & "','" & a(5) & "','" & a(6) & "')"
  cnmovie.Execute a
  MsgBox "保存成功！"
Case 1
 a(0) = Text1.Text
 a(1) = Text2.Text
 a(2) = Text3.Text
 a(3) = Text4.Text
 a(4) = Text5.Text
 a(5) = Text6.Text
 a(6) = Text7.Text
 a = "update cinema set cinid='" & a(0) & "',cincity='" & a(1) & "',cingrade='" & a(2) & "',cinnam='" & a(3) & "',cinloc='" & a(4) & "',cinphone='" & a(5) & "',price='" & a(6) & "' where cinid='" & a(0) & "'"
cnmovie.Execute a
  MsgBox "修改完毕！"
  Text1(0).Locked = False
  editf = 0
  End Select
  Image1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
End Sub

Private Sub Form_Load() '显示窗体的时候显示影院信息
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
Me.AutoRedraw = True
n = 1
rs4.CursorLocation = adUseClient
rs4.Open "select * from cinema", cnmovie, adOpenDynamic, adLockOptimistic
Image1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Call viewdata

End Sub

Private Sub Form_Resize() '背景图片大小随窗体变动
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub
Private Sub viewdata()
Image1.Picture = LoadPicture(App.Path + "\..\photo\电影院\" + rs4.Fields("cinphoto"))
Text1 = rs4.Fields("cinid")
Text2 = rs4.Fields("cincity")
Text3 = rs4.Fields("cingrade")
Text4 = rs4.Fields("cinid") + rs4.Fields("cinnam")
Text5 = rs4.Fields("cinloc")
Text6 = rs4.Fields("cinphone")
Text7 = rs4.Fields("cinprice")
End Sub
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

Private Sub Command9_Click()
guanliyuan.Show
End Sub


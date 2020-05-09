VERSION 5.00
Begin VB.Form messagegl 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   14835
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3960
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   600
      Width           =   5055
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   3840
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   3240
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÉÏ´«"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "É¾³ý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "·µ»Ø"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ë¢ÐÂ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   5055
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   1
      Top             =   2640
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   0
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "»Ø¸´:"
      BeginProperty Font 
         Name            =   "ºÚÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ÒÑ»Ø¸´ÁôÑÔ:"
      BeginProperty Font 
         Name            =   "ºÚÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "»Ø¸´:"
      BeginProperty Font 
         Name            =   "ºÚÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Î´»Ø¸´ÁôÑÔ:"
      BeginProperty Font 
         Name            =   "ºÚÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "messagegl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs15 As New ADODB.Recordset
Dim a As Long

Private Sub Combo1_Click()
Dim c As Long
c = Val(Combo1.Text)
rs15.CursorLocation = adUseClient
rs15.Open "select * from message where mesid=" & c & "", cnmovie, adOpenDynamic, adLockOptimistic
Text4.Text = rs15.Fields("message")
rs15.Close
End Sub

Private Sub Combo2_Click()
Dim b As Long, c As Long
c = Val(Combo2.Text)
rs15.CursorLocation = adUseClient
rs15.Open "select * from message where mesid=" & c & "", cnmovie, adOpenDynamic, adLockOptimistic
Text3.Text = rs15.Fields("message")
b = rs15.Fields("huifuid")
rs15.Close
rs15.Open "select * from huifu where huifuid=" & b & " ", cnmovie, adOpenDynamic, adLockOptimistic
Text2.Text = rs15.Fields("huifu")
rs15.Close
End Sub

Private Sub Command1_Click()
Dim t As String, c As Long
c = Val(Combo1.Text)
rs15.Open "select * from huifu ", cnmovie, adOpenDynamic, adLockOptimistic
If Combo1.Text <> "" Then
If Text1.Text <> "" Then
t = "insert into huifu(huifu,userid) values('" & Text1.Text & "','" & uid1 & "')"
cnmovie.Execute t
rs15.Close
rs15.Open "select * from huifu where huifu='" & Text1.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
a = rs15.Fields("huifuid")
rs15.Close
rs15.Open "select * from message where mesid=" & c & "", cnmovie, adOpenDynamic, adLockOptimistic
t = "update message set huifuid=" & a & " where mesid=" & c & ""
cnmovie.Execute t
rs15.Close
MsgBox ("»Ø¸´³É¹¦£¡")
Combo1.RemoveItem Combo1.ListIndex
Text1.Text = ""
Else
MsgBox ("ÇëÊäÈë»Ø¸´ÄÚÈÝ£¡")
End If
Else
MsgBox ("ÇëÑ¡ÔñÐèÒª»Ø¸´µÄÁôÑÔ£¡")
End If
End Sub

Private Sub Command2_Click()
Dim t As String
If rs15.State = adStateOpen Then rs15.Close
rs15.Open "select * from message", cnmovie, adOpenDynamic, adLockOptimistic
If Combo2.Text <> "" Then
t = "delete from message where message ='" & Text3.Text & "'"
cnmovie.Execute t
Combo2.RemoveItem Combo2.ListIndex
Text2.Text = ""
Text3.Text = ""
MsgBox ("É¾³ý³É¹¦£¡")
Else
MsgBox ("ÇëÑ¡ÔñÐèÒªÉ¾³ýµÄÁôÑÔ£¡")
End If
End Sub

Private Sub Command3_Click()
Unload messagegl
guanliyuan.Show
End Sub

Private Sub Command4_Click()
If rs15.State = adStateOpen Then rs15.Close
rs15.CursorLocation = adUseClient
rs15.Open "select * from message where mesyes=true", cnmovie, adOpenDynamic, adLockOptimistic
Combo2.Clear
Text2.Text = ""
If rs15.EOF And rs15.BOF Then
MsgBox ("±íÖÐÎÞ¼ÇÂ¼!")
End If
For i = 0 To rs15.RecordCount - 1
Combo2.AddItem rs15.Fields("mesid")
rs15.MoveNext
Next i
rs15.Close
End Sub

Private Sub Form_Load()
rs15.CursorLocation = adUseClient
rs15.Open "select * from message where mesyes= true", cnmovie, adOpenDynamic, adLockOptimistic
If rs15.EOF And rs15.BOF Then
MsgBox ("±íÖÐÎÞ¼ÇÂ¼!")
End If
For i = 0 To rs15.RecordCount - 1
Combo2.AddItem rs15.Fields("mesid")
rs15.MoveNext
Next i
If rs15.State = adStateOpen Then rs15.Close
rs15.CursorLocation = adUseClient
rs15.Open "select * from message where mesyes= false", cnmovie, adOpenDynamic, adLockOptimistic
For i = 0 To rs15.RecordCount - 1
Combo1.AddItem rs15.Fields("mesid")
rs15.MoveNext
Next i
rs15.Close
End Sub

VERSION 5.00
Begin VB.Form liuyan 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   12615
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ύ����"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "��Ҫ����"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "liuyan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs8 As New ADODB.Recordset
Dim d As String, e As String, f As String

Private Sub Command1_Click()
If Text1.Text = "" Then
  MsgBox "����Ϊ�գ������ύ"
Else
  rs8.CursorLocation = adUseClient
  rs8.Open "select  *  from message where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
  a = uid
  b = Text1.Text
  c = "insert into message(message,userid) values('" & b & "','" & a & "')"
  rs8.MoveLast
  d = rs8.Fields("mesid")
  cnmovie.Execute c
  MsgBox "�ύ�ɹ��������Ա��Ϊ" + d
  rs8.Close
End If
End Sub

Private Sub Command3_Click()
movie.Show
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\����.jpg")
Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

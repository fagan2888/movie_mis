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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command6 
      Caption         =   "�����޸�"
      Height          =   615
      Left            =   9000
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�޸�Ӱ��"
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�鿴Ӱ��"
      Height          =   615
      Left            =   5280
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   615
      Left            =   9480
      TabIndex        =   10
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ύӰ��"
      Height          =   615
      Left            =   6240
      TabIndex        =   9
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ɾ����Ӱ��"
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
      Text            =   "��ֻ����ӰƬ����ֻ����ӰƬ��ѡҳ��ӰƬ��ţ�������������010101"
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
      Caption         =   "Ӱ������"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ӰƬ��"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "дӰ��"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "�ҵ�Ӱ��"
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

Private Sub Command5_Click() '�޸�Ӱ��
Text1.Enabled = True
End Sub

Private Sub Command6_Click()
If Text1.Text = "" Then
  MsgBox "Ӱ��Ϊ�գ������ύ"
Else
  sql = "update review set review='" & Text1.Text & "' where revid=" & e & ""
  cnmovie.Execute sql
  MsgBox "�޸���ϣ�"
  Text1.Enabled = False
End If
rs9.Close
Set rs9 = Nothing
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\����.jpg")
Me.AutoRedraw = True
rs9.CursorLocation = adUseClient
rs9.Open "select  revid  from review where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
If rs9.RecordCount = 0 Then
    MsgBox ("����δ�����Ӱ����")
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

Private Sub Command2_Click() '�ύӰ��
tijiaoyingping.Show
If Text2.Text = "" Then
  MsgBox "Ӱ��Ϊ�գ������ύ"
Else
  rs9.CursorLocation = adUseClient
  rs9.Open "select * from movie where movname='" & Text3.Text & "'or movnum='" & Text3.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
  If rs9.BOF And rs9.EOF Then MsgBox "û�д˵�Ӱ���ܱ�ǸŶ~��л���ķ���"
  If rs9.RecordCount > 1 Then MsgBox "�жಿ�����ֵĵ�ӰŶ����������ҳ���ע��Ӱ���,��ֻ�����Ӱ���~"
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
  MsgBox "�ύ�ɹ�����Ӱ�����Ϊ" + d
  rs9.Close
  Set rs9 = Nothing
End If
End Sub

Private Sub Command3_Click()
yingping.Hide
movie.Show
End Sub

Private Sub Command1_Click() 'ɾ��Ӱ��
e = Val(Combo1.Text)
rs9.CursorLocation = adUseClient
rs9.Open "select  *  from review where revid=" & e & "", cnmovie, adOpenDynamic, adLockOptimistic
  h = MsgBox("�Ƿ�Ҫɾ���ü�¼��", vbYesNo)
 If h = vbYes Then
  k = "delete from review where revid=" & e & ""
  cnmovie.Execute k
  MsgBox "ɾ���ɹ�"
  rs9.Close
  Set rs9 = Nothing
 End If
End Sub

Private Sub Command4_Click() '�鿴Ӱ��
e = Val(Combo1.Text)
rs9.CursorLocation = adUseClient
rs9.Open "select  *  from review where revid=" & e & "", cnmovie, adOpenDynamic, adLockOptimistic
Text1.Text = rs9.Fields("review")
rs9.Close
Set rs9 = Nothing
End Sub


VERSION 5.00
Begin VB.Form find 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   15060
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "find.frx":0000
      Left            =   3360
      List            =   "find.frx":0010
      TabIndex        =   9
      Text            =   "��Ĺ�����ʦ�ǣ�"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���������ˣ��˳�"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���æ���һ�����"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "���룺"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "�𰸣�"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "�ܱ����⣺"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�û�����"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs3 As New ADODB.Recordset

Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "�û�������Ϊ��,���������룡"
    Text1.SetFocus
    Exit Sub
End If
If Combo1.Text = "" Then
    MsgBox "��ѡ���ܱ����⣡"
    Exit Sub
End If
If Text3.Text = "" Then
    MsgBox "�ܱ��𰸲���Ϊ�գ����������룡"
    Text2.SetFocus
    Exit Sub
End If
    sql = "select * from user where userid='" & Text1.Text & "'"
    rs3.Open sql, cnmovie, adOpenDynamic, adLockOptimistic
    If rs3.RecordCount = 0 Then
    MsgBox "�û�������������������룡"
    rs3.Close
    Else
    If Trim(Combo1.Text) <> rs3.Fields("userquestion") Then
    MsgBox "��ѡ����ܱ���������������ѡ��"
    rs3.Close
    Else
    If Trim(Text3.Text) <> rs3.Fields("useranswer") Then
    MsgBox "�ܱ��𰸲���ȷ�����������룡"
    rs3.Close
    Else
    MsgBox "�����������һأ���ȷ�ϣ���������޸�"
    Text4.Text = rs3.Fields("userpassword")
    rs3.Close
    Set rs11 = Nothing
    End If
    End If
    End If
End Sub

Private Sub Command2_Click()
MsgBox "��ϲ�㣡�һ�����"
End
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\����.jpg")
Me.AutoRedraw = True
End Sub
Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form zhuce 
   Caption         =   "ע��"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   15015
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "�ϴ�ͷ��"
      Height          =   495
      Left            =   13200
      TabIndex        =   41
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   0
      ItemData        =   "zhuce.frx":0000
      Left            =   4320
      List            =   "zhuce.frx":000A
      TabIndex        =   35
      Text            =   "��ͨ"
      Top             =   3600
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13320
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   6600
      TabIndex        =   25
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   615
      Left            =   4440
      TabIndex        =   24
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   23
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   5520
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   3
      ItemData        =   "zhuce.frx":0019
      Left            =   4320
      List            =   "zhuce.frx":0029
      TabIndex        =   21
      Text            =   "��Ĺ�����ʦ�ǣ�"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   2
      ItemData        =   "zhuce.frx":0071
      Left            =   4320
      List            =   "zhuce.frx":0087
      TabIndex        =   20
      Text            =   "����"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   1
      ItemData        =   "zhuce.frx":00AF
      Left            =   4320
      List            =   "zhuce.frx":00B9
      TabIndex        =   19
      Text            =   "��"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   4
      Left            =   4320
      TabIndex        =   18
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   3
      Left            =   4320
      TabIndex        =   17
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   2
      Left            =   4320
      TabIndex        =   16
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "zhuce.frx":00C5
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "VIP�㲥��Ӱ��ѣ����ײ�5Ԫ/��"
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "����ʦ��˵��������Ա�˺�Ϊϵͳ��̨������Ա����ģ�����ע�ᣬ��Ϊʵ�������й���Ա�˺�����ע����������"
      Height          =   1215
      Left            =   600
      TabIndex        =   39
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "������Ҫ�������������ַ"
      Height          =   255
      Left            =   6960
      TabIndex        =   38
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "ע��ע�⣬���Ǻ�Ϊ������Ŷ"
      Height          =   255
      Left            =   4080
      TabIndex        =   37
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "����ʱ�侫�������Ͳ����ر�ȫ��������½�"
      Height          =   375
      Left            =   6840
      TabIndex        =   36
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "���ǵĸ�ʽ��2011/1/1"
      Height          =   255
      Left            =   6960
      TabIndex        =   34
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "��������д11λ�ֻ���"
      Height          =   255
      Left            =   6960
      TabIndex        =   33
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "������չʾ��ķ�ɰɣ�Ҳ����ʹ��Ĭ��ͷ��,"
      Height          =   1095
      Left            =   13080
      TabIndex        =   32
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   4
      Left            =   6480
      TabIndex        =   31
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   3
      Left            =   6480
      TabIndex        =   30
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   2
      Left            =   6480
      TabIndex        =   29
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   1
      Left            =   6480
      TabIndex        =   28
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "*"
      Height          =   135
      Index           =   0
      Left            =   6480
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "�û�ע�᣺"
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�û���"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��ϵ�绰"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�Ա�"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�û�ƫ��"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�ܱ�����"
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�ܱ���"
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   6
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ͷ��"
      Height          =   255
      Index           =   10
      Left            =   8880
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�û�����"
      Height          =   255
      Index           =   11
      Left            =   8760
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ȷ������"
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "zhuce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim str As String
Dim i As Integer, c As String, d As String
Dim a(12) As String
Private Sub Command1_Click()
If Combo1(0).Text = "VIP" Then
    b = MsgBox("ȷ��֧��5Ԫ/�£�", vbYesNo)
 If b = vbYes Then MsgBox "֧���ɹ�"
 End If
For i = 0 To 3
  If Text1(i).Text = "" Then
   MsgBox "����Ϊ������Ŷ��"
   Text1(i).SetFocus
   Exit Sub
  End If
  Next i
  If Len(Text1(2).Text) <> 11 Then
  MsgBox "������11λ�ֻ���"
  Exit Sub
  Text1(2).SetFocus
  End If
  If Text1(1).Text <> Text4.Text Then MsgBox "�����������벻��ͬ�����飡"
i = 1
For i = 1 To 11
 d = Right(Text1(2).Text, i)
 If Asc(d) < 48 Or Asc(d) > 57 Then MsgBox ("�����з������ַ�����")
Next i
str = "select * from user where userid='" & Trim(Text1(0).Text) & "';"
rs2.Open str, cnmovie, adOpenDynamic, adLockPessimistic
   If Not rs2.EOF Then
     MsgBox "�û����Ѵ��ڣ�����������"
     Text1(0).Text = ""
     Text1(0).SetFocus
     rs2.Close
   Exit Sub
   End If
i = 0
  For i = 0 To 4
    a(i) = Text1(i).Text
    Next i
i = 0
For i = 0 To 3
    a(i + 5) = Combo1(i).Text
    Next i
    a(9) = Text3.Text
    a(11) = Text2.Text
    If Command3.Enabled = False Then
    a(10) = Text1(0).Text & ".jpg"
    FileCopy CommonDialog1.FileName, App.Path & "\..\photo\�û�ͷ��\" & a(10)
    Else
     a(10) = "moren.jpg"
    End If
 c = "insert into user(userid,userpassword,userphone,usermail,userbir,usertype,usersex,userprefer,userquestion,useranswer,userphoto,userresume)  values('" & a(0) & "','" & a(1) & "','" & a(2) & "','" & a(3) & "','" & a(4) & "','" & a(5) & "','" & a(6) & "','" & a(7) & "','" & a(8) & "','" & a(9) & "','" & a(10) & "','" & a(11) & "')"
  cnmovie.Execute c
  MsgBox "ע��ɹ���"
End Sub

Private Sub Command2_Click()
b = MsgBox("���ߣ���û���ע���أ����Ҫ�뿪ô��", vbYesNo)
 If b = vbYes Then shouye.Show
End Sub


Private Sub Command3_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Command3.Enabled = False
Image1.Picture = LoadPicture(CommonDialog1.FileName)
End If
'CommonDialog1.ShowSave
'FileCopy CommonDialog1.FileName
'App.Path
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\����.jpg")
Me.AutoRedraw = True
Image1.Picture = LoadPicture(App.Path & "\..\photo\�û�ͷ��\moren.jpg")
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

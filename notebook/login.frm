VERSION 5.00
Begin VB.Form login 
   Caption         =   "login"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   14715
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command6 
      Caption         =   "ע��"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�������"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   9
      Text            =   "admin"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Text            =   "admin"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��������"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����Ա"
      Height          =   615
      Left            =   7680
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�û���¼"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "��û���˺ţ���ȥע�����԰�~"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "���룺"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "�û�����"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "������˵������ͨ�û���admin������admin����û��ʹ��������Ϊ�������á���������Ա�û���guanli001,����guanli001,ϣ����ʦʹ�����~"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   6240
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   $"login.frx":0000
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim n As Integer
Private Sub Command3_Click()
 b = MsgBox("��Ҫ���Ҫ���ٿ�����", vbYesNo)
 If b = vbNo Then shouye.Show
End Sub

Private Sub Command5_Click()
find.Show
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\����.jpg")
Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.PasswordChar = ""
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.PasswordChar = "*"
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
 MsgBox "�û�������Ϊ�գ����ż�"
 Text1.SetFocus
 Exit Sub
End If
If Text2.Text = "" Then
 MsgBox "���벻��Ϊ�գ����ż�"
 Text2.Text = ""
 Exit Sub
End If
rs1.CursorLocation = adUseClient
rs1.Open "select  *  from user where userid='" & Text1.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
If rs1.RecordCount = 0 Then
   MsgBox "����̫�ԣ���ע������������û���"
   n = n + 1
         If n = 3 Then
              MsgBox "������Ĵ�������Ѿ��ﵽ3�Σ������Сд�����ٻ���һ�£��Ժ��ڽ��е�½"
              End
         End If
    Text1.Text = ""
    Text1.SetFocus
    rs1.Close
    Set rs1 = Nothing
    Exit Sub
Else
    If Trim(Text2.Text) <> rs1.Fields("userpassword") Then
         MsgBox "��������ʺŻ������������������"
         Text2.Text = ""
         Text2.SetFocus
         rsclothes.Close
         Set rsclothes = Nothing
         Exit Sub
         n = n + 1
         If n = 3 Then
             MsgBox "��������ʺŻ������������Ѿ��ﵽ3�Σ�Ҫ��Ҫ���һ�����룿������������һ�Ŷ�����������ɣ�"
             rs1.Close
             Set rs1 = Nothing
             End
         End If
     End If
End If

rs1.Close
Set rs1 = Nothing
MsgBox "��½�ɹ�����������ɣ�"
uid = Text1.Text
movie.Show
End Sub

Private Sub Command6_Click()
zhuce.Show
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then
 MsgBox "�û�������Ϊ�գ����ż�"
 Text1.SetFocus
 Exit Sub
End If
If Text2.Text = "" Then
 MsgBox "���벻��Ϊ�գ����ż�"
 Text2.Text = ""
 Exit Sub
End If
rs1.CursorLocation = adUseClient
rs1.Open "select  *  from management where userid='" & Text1.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
If rs1.RecordCount = 0 Then
   MsgBox "����̫�ԣ���ע������������û���"
   n = n + 1
         If n = 3 Then
              MsgBox "������Ĵ�������Ѿ��ﵽ3�Σ������Сд�����ٻ���һ�£��Ժ��ڽ��е�½"
              End
         End If
    Text1.Text = ""
    Text1.SetFocus
    rs1.Close
    Set rs1 = Nothing
    Exit Sub
Else
    If Trim(Text2.Text) <> rs1.Fields("userpassword") Then
         MsgBox "��������ʺŻ������������������"
         Text2.Text = ""
         Text2.SetFocus
         rs1.Close
         Set rsclothes = Nothing
         Exit Sub
         n = n + 1
         If n = 3 Then
             MsgBox "��������ʺŻ������������Ѿ��ﵽ3�Σ�Ҫ��Ҫ���һ�����룿������������һ�Ŷ�����������ɣ�"
             rs1.Close
             Set rs1 = Nothing
             End
         End If
     End If
End If

rs1.Close
Set rs1 = Nothing
MsgBox "��½�ɹ�����������ɣ�"
uid1 = Text1.Text
guanliyuan.Show
End Sub


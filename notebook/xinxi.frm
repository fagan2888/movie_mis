VERSION 5.00
Begin VB.Form xinxi 
   Caption         =   "������Ϣ"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   15090
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   6600
      TabIndex        =   18
      Top             =   3840
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   3720
      TabIndex        =   16
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   15
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   3720
      TabIndex        =   14
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   3
      Left            =   3720
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Index           =   4
      Left            =   3720
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "���˽���"
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "ͷ��"
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�û�ƫ��"
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   11
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�Ա�"
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��ϵ�绰"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�û���"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "xinxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs6 As New ADODB.Recordset
Dim i As Integer
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\����.jpg")
Me.AutoRedraw = True
For i = 0 To 7
  Text1(i).Enabled = False
Next i
Text2.Enabled = False
rs6.CursorLocation = adUseClient
rs6.Open "select  *  from user where userid='" & uid & "'", cnmovie, adOpenDynamic, adLockOptimistic
For i = 0 To 7
  Text1(i).Text = rs6.Fields(i)
Next i
Text2.Text = rs6.Fields("userresume")
Image1.Picture = LoadPicture(App.Path + "\..\photo\�û�ͷ��\" + rs6.Fields("userphoto"))
rs6.Close
Set rs6 = Nothing
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

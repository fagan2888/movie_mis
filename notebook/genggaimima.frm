VERSION 5.00
Begin VB.Form genggaimima 
   Caption         =   "��������"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11685
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ���޸�"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   4560
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   4560
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "ȷ������"
      Height          =   255
      Index           =   12
      Left            =   3480
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "������"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�û���"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "genggaimima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs3 As New ADODB.Recordset
Private Sub Command1_Click()
    If Text3.Text = "" Then
    MsgBox "�޸����벻��Ϊ�գ�����������"
    Text2.SetFocus
    Exit Sub
    End If
    If Text2.Text = Text3.Text Then
    MsgBox "�޸ĺ�����벻����ԭ������ͬ"
    End If
    If Text3.Text <> Text4.Text Then
    MsgBox "������ȷ�����벻���������������"
    Else
    rs3.Fields("userid") = Text1.Text
    rs3.Fields("userpassword") = Text2.Text
    rs3.Update
    MsgBox ("�û���¼��Ϣ�޸ĳɹ�")
    Unload genggaimima
    rs3.Close
    Set rs3 = Nothing
    login.Show
    End If
End Sub

Private Sub Command2_Click()
    MsgBox "�û���Ϣδ�޸�"
    Unload genggaimima
    movie.Show
End Sub

Private Sub Form_Load()
    Text1.Text = login.Text1.Text
    Text2.Text = login.Text2.Text
    Text1.Enabled = False
    rs3.CursorLocation = adUseClient
    rs3.Open "select  *  from user where userid='" & Text1.Text & "'", cnmovie, adOpenDynamic, adLockOptimistic
    Me.Picture = LoadPicture(App.Path & "\����.jpg")
    Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

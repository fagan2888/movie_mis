VERSION 5.00
Begin VB.Form movie 
   Caption         =   "�û�ҳ��"
   ClientHeight    =   7815
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   15135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "�鿴Ӱ��"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ӰԺ�Ƽ�"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��Ӱ��ѡ"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "�˵���������Ӵ~"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "һ��һ���磬һҶһ���ᡣ�ڵ�Ӱ�и��ܲ�ͬ������"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "��ϵ����"
      Height          =   255
      Left            =   12000
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Menu GRXX 
      Caption         =   "������Ϣ"
      Begin VB.Menu XGMM 
         Caption         =   "�޸�����"
      End
      Begin VB.Menu XGXX 
         Caption         =   "�޸���Ϣ"
      End
      Begin VB.Menu CKXX 
         Caption         =   "�鿴��Ϣ"
      End
      Begin VB.Menu TC 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu CKDY 
      Caption         =   "�鿴��Ӱ"
      Begin VB.Menu WDLB 
         Caption         =   "�ҵ��б�"
      End
   End
   Begin VB.Menu LYFK 
      Caption         =   "���Է���"
      Begin VB.Menu CKHF 
         Caption         =   "�鿴�ظ�"
      End
      Begin VB.Menu SCLY 
         Caption         =   "ɾ������"
      End
      Begin VB.Menu WDLY 
         Caption         =   "�ҵ�����"
      End
   End
   Begin VB.Menu CKYP 
      Caption         =   "�鿴Ӱ��"
      Begin VB.Menu WDYP 
         Caption         =   "�ҵ�Ӱ��"
      End
      Begin VB.Menu SCYP 
         Caption         =   "ɾ���ҵ�Ӱ��"
      End
   End
End
Attribute VB_Name = "movie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKHF_Click()
liuyanhuifu.Show
End Sub

Private Sub CKXX_Click()
xinxi.Show
End Sub

Private Sub Command1_Click()
movieshow.Show
End Sub

Private Sub Command2_Click()
cinema.Show
End Sub

Private Sub Command3_Click()
yingpingchakan.Show
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\����.jpg")
Me.AutoRedraw = True
End Sub

Private Sub Form_Resize()
Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, Me.Picture.Width / 26.45836 * 15, Me.Picture.Height / 26.45836 * 15
End Sub

Private Sub Label1_Click()
lianxi.Show
End Sub

Private Sub SCLY_Click()
liuyanhuifu.Show
End Sub

Private Sub SCYP_Click()
yingping.Show
End Sub

Private Sub TC_Click()
 b = MsgBox("��Ҫ���Ҫ���ٿ�����", vbYesNo)
 If b = vbNo Then End
End Sub

Private Sub WDLB_Click()
list.Show
End Sub

Private Sub WDLY_Click()
liuyan.Show
End Sub

Private Sub WDYP_Click()
yingping.Show
End Sub

Private Sub XGMM_Click()
genggaimima.Show
End Sub

Private Sub XGXX_Click()
xinxixiugai.Show
End Sub

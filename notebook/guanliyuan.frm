VERSION 5.00
Begin VB.Form guanliyuan 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   15180
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label Label1 
      Caption         =   "����Լ����еĲ�������"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4560
      TabIndex        =   0
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Menu YHGL 
      Caption         =   "�û�����"
      Begin VB.Menu CKXX 
         Caption         =   "�鿴��Ϣ"
      End
      Begin VB.Menu XGQX 
         Caption         =   "�޸�Ȩ��"
      End
      Begin VB.Menu GLYH 
         Caption         =   "�����û�"
      End
   End
   Begin VB.Menu DYGL 
      Caption         =   "��Ӱ����"
      Begin VB.Menu CKDY 
         Caption         =   "�鿴��Ϣ"
      End
   End
   Begin VB.Menu YYGL 
      Caption         =   "ӰԺ����"
      Begin VB.Menu CKYY 
         Caption         =   "�鿴��Ϣ"
      End
      Begin VB.Menu GLYY 
         Caption         =   "ӰԺ����"
      End
   End
   Begin VB.Menu LYGL 
      Caption         =   "���Թ���"
      Begin VB.Menu GLLY 
         Caption         =   "��������"
      End
      Begin VB.Menu HFLY 
         Caption         =   "�ظ�����"
      End
   End
   Begin VB.Menu YPGL 
      Caption         =   "Ӱ������"
      Begin VB.Menu SCYP 
         Caption         =   "ɾ��Ӱ��"
      End
   End
End
Attribute VB_Name = "guanliyuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKDY_Click()
moviechakan.Show
End Sub

Private Sub CKXX_Click()
chakanxinxi.Show
End Sub

Private Sub CKYY_Click()
cinemagl.Show
End Sub

Private Sub GLDY_Click()
movieguanli.Show
End Sub

Private Sub GLLY_Click()
messagegl.Show
End Sub

Private Sub GLYH_Click()
guanliyonghu.Show
End Sub



Private Sub GLYY_Click()
cinemagl.Show
End Sub

Private Sub HFLY_Click()
messagegl.Show
End Sub

Private Sub SCYP_Click()
yingpinggl.Show
End Sub

Private Sub XGQX_Click()
guanliyonghu.Show
End Sub

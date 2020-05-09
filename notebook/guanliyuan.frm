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
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      Caption         =   "请对自己进行的操作负责"
      BeginProperty Font 
         Name            =   "幼圆"
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
      Caption         =   "用户管理"
      Begin VB.Menu CKXX 
         Caption         =   "查看信息"
      End
      Begin VB.Menu XGQX 
         Caption         =   "修改权限"
      End
      Begin VB.Menu GLYH 
         Caption         =   "管理用户"
      End
   End
   Begin VB.Menu DYGL 
      Caption         =   "电影管理"
      Begin VB.Menu CKDY 
         Caption         =   "查看信息"
      End
   End
   Begin VB.Menu YYGL 
      Caption         =   "影院管理"
      Begin VB.Menu CKYY 
         Caption         =   "查看信息"
      End
      Begin VB.Menu GLYY 
         Caption         =   "影院管理"
      End
   End
   Begin VB.Menu LYGL 
      Caption         =   "留言管理"
      Begin VB.Menu GLLY 
         Caption         =   "管理留言"
      End
      Begin VB.Menu HFLY 
         Caption         =   "回复留言"
      End
   End
   Begin VB.Menu YPGL 
      Caption         =   "影评管理"
      Begin VB.Menu SCYP 
         Caption         =   "删除影评"
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

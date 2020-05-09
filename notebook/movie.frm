VERSION 5.00
Begin VB.Form movie 
   Caption         =   "用户页面"
   ClientHeight    =   7815
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   15135
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "查看影评"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "影院推荐"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "电影精选"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "菜单栏在这里哟~"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "一花一世界，一叶一菩提。在电影中感受不同的人生"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "联系我们"
      Height          =   255
      Left            =   12000
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Menu GRXX 
      Caption         =   "个人信息"
      Begin VB.Menu XGMM 
         Caption         =   "修改密码"
      End
      Begin VB.Menu XGXX 
         Caption         =   "修改信息"
      End
      Begin VB.Menu CKXX 
         Caption         =   "查看信息"
      End
      Begin VB.Menu TC 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu CKDY 
      Caption         =   "查看电影"
      Begin VB.Menu WDLB 
         Caption         =   "我的列表"
      End
   End
   Begin VB.Menu LYFK 
      Caption         =   "留言反馈"
      Begin VB.Menu CKHF 
         Caption         =   "查看回复"
      End
      Begin VB.Menu SCLY 
         Caption         =   "删除留言"
      End
      Begin VB.Menu WDLY 
         Caption         =   "我的留言"
      End
   End
   Begin VB.Menu CKYP 
      Caption         =   "查看影评"
      Begin VB.Menu WDYP 
         Caption         =   "我的影评"
      End
      Begin VB.Menu SCYP 
         Caption         =   "删除我的影评"
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
Me.Picture = LoadPicture(App.Path & "\背景.jpg")
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
 b = MsgBox("不要走嘛，要不再看看？", vbYesNo)
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

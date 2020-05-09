VERSION 5.00
Begin VB.Form tijiaoyingping 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6000
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "tijiaoyingping.frx":0000
      Top             =   360
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "tijiaoyingping.frx":004B
      Top             =   1200
      Width           =   4815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "我已阅读并承诺遵循上述规定进行相关操作"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "我不同意上述规定"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   3855
   End
End
Attribute VB_Name = "tijiaoyingping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
tijiaoyingping.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub Option2_Click()
Command1.Enabled = True
End Sub

Private Sub Option1_Click()
Command1.Enabled = False
End Sub

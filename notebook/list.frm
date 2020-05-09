VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form list 
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   14745
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   14280
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Left            =   14280
      Top             =   4320
   End
   Begin VB.CommandButton Command4 
      Caption         =   "循环播放"
      Height          =   375
      Left            =   13320
      TabIndex        =   6
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "顺序播放"
      Height          =   375
      Left            =   12240
      TabIndex        =   5
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一个"
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上一个"
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   6240
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   5280
      ItemData        =   "list.frx":0000
      Left            =   10200
      List            =   "list.frx":0002
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
   Begin VB.PictureBox WindowsMediaPlayer1 
      Height          =   7695
      Left            =   240
      ScaleHeight     =   7635
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
         Height          =   7695
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   9495
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   16748
         _cy             =   13573
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"list.frx":0004
      Height          =   855
      Left            =   10080
      TabIndex        =   7
      Top             =   6840
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "查看我的列表"
      Height          =   255
      Left            =   10200
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs5 As New ADODB.Recordset
Private Sub Form_Load()
rs5.CursorLocation = adUseClient
rs5.Open "select dianbo.userid,movie.movnum,dianbo.movnum,movie.movname  from dianbo ,movie  where userid='" & uid & "'And dianbo.movnum = movie.movnum", cnmovie, adOpenDynamic, adLockOptimistic
If rs5.RecordCount = 0 Then MsgBox "您还没有点播呢，快去选择电影吧"
   rs5.MoveFirst
For i = 0 To rs5.RecordCount - 1
  b = rs5.Fields("movname")
  List1.AddItem b
  rs5.MoveNext
Next i
End Sub
Private Sub Command1_Click()
WindowsMediaPlayer1.URL = List1.list(List1.ListIndex - 1)
List1.ListIndex = List1.ListIndex - 1
End Sub
Private Sub Command2_Click()
WindowsMediaPlayer1.URL = List1.list(List1.ListIndex + 1)
List1.ListIndex = List1.ListIndex + 1
End Sub
Private Sub Command3_Click()
Timer1.Enabled = True
End Sub
Private Sub Command4_Click()
Timer2.Enabled = True
End Sub
Private Sub Timer1_Timer()
If WindowsMediaPlayer1.PlayState = wmppsStopped Then
  WindowsMediaPlayer1.URL = List1.list(List1.ListIndex + 1)
  List1.ListIndex = List1.ListIndex + 1
End If
End Sub


Private Sub Timer2_Timer()
If WindowsMediaPlayer1.PlayState = wmppsStopped Then
 If List1.Index = List1.ListCount - 1 Then
    List1.ListIndex = 0
    WindowsMediaPlayer1.URL = List1.list(List1.ListIndex)
 Else
  WindowsMediaPlayer1.URL = List1.list(List1.ListIndex + 1)
  List1.ListIndex = List1.ListIndex + 1
End If
End If
End Sub


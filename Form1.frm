VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Begin VB.Form Form1 
   Caption         =   "Agent Simple"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   63
      Text            =   "Form1.frx":0000
      Top             =   7440
      Width           =   5415
   End
   Begin VB.CommandButton Command58 
      Caption         =   "SHOW"
      Height          =   255
      Left            =   7680
      TabIndex        =   61
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command57 
      Caption         =   "HIDE"
      Height          =   255
      Left            =   7680
      TabIndex        =   60
      Top             =   6600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speak"
      Height          =   735
      Left            =   120
      TabIndex        =   57
      Top             =   6120
      Width           =   7455
      Begin VB.CommandButton Command45 
         Caption         =   "SAY"
         Height          =   375
         Left            =   6240
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Text            =   "Text to Speech"
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CommandButton Command56 
      Caption         =   "Thinking"
      Height          =   375
      Left            =   7440
      TabIndex        =   56
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command55 
      Caption         =   "Reading"
      Height          =   375
      Left            =   7440
      TabIndex        =   55
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command54 
      Caption         =   "Processing"
      Height          =   375
      Left            =   7440
      TabIndex        =   54
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command53 
      Caption         =   "WriteReturn"
      Height          =   375
      Left            =   7440
      TabIndex        =   53
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command52 
      Caption         =   "WriteContinued"
      Height          =   375
      Left            =   7440
      TabIndex        =   52
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command51 
      Caption         =   "Write"
      Height          =   375
      Left            =   7440
      TabIndex        =   51
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Wave"
      Height          =   375
      Left            =   7440
      TabIndex        =   50
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Uncertain"
      Height          =   375
      Left            =   7440
      TabIndex        =   49
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Think"
      Height          =   375
      Left            =   7440
      TabIndex        =   48
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command47 
      Caption         =   "Surprised"
      Height          =   375
      Left            =   7440
      TabIndex        =   47
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command46 
      Caption         =   "Suggest"
      Height          =   375
      Left            =   7440
      TabIndex        =   46
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command44 
      Caption         =   "StartListening"
      Height          =   375
      Left            =   5640
      TabIndex        =   45
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command43 
      Caption         =   "Search"
      Height          =   375
      Left            =   5640
      TabIndex        =   44
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Sad"
      Height          =   375
      Left            =   5640
      TabIndex        =   43
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Readheturn"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   42
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command40 
      Caption         =   "ReadContinued"
      Height          =   375
      Left            =   5640
      TabIndex        =   41
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Read"
      Height          =   375
      Left            =   5640
      TabIndex        =   40
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Process"
      Height          =   375
      Left            =   5640
      TabIndex        =   39
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Pleased"
      Height          =   375
      Left            =   5640
      TabIndex        =   38
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command36 
      Caption         =   "MoveRight"
      Height          =   375
      Left            =   5640
      TabIndex        =   37
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command35 
      Caption         =   "MoveLeft"
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command34 
      Caption         =   "StopListening"
      Height          =   375
      Left            =   5640
      TabIndex        =   35
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command33 
      Caption         =   "MoveDown"
      Height          =   375
      Left            =   3840
      TabIndex        =   34
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command32 
      Caption         =   "LookUpReturn"
      Height          =   375
      Left            =   3840
      TabIndex        =   33
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command31 
      Caption         =   "LookUpBlink"
      Height          =   375
      Left            =   3840
      TabIndex        =   32
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command30 
      Caption         =   "LookUp"
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command29 
      Caption         =   "LookRightReturn"
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command28 
      Caption         =   "LookRightBlink"
      Height          =   375
      Left            =   3840
      TabIndex        =   29
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command27 
      Caption         =   "LookRight"
      Height          =   375
      Left            =   3840
      TabIndex        =   28
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command26 
      Caption         =   "LookLeftReturn"
      Height          =   375
      Left            =   3840
      TabIndex        =   27
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command25 
      Caption         =   "LookLeftBlink"
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command24 
      Caption         =   "LookLeft"
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command23 
      Caption         =   "MoveUp"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command22 
      Caption         =   "LookDownReturn"
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Explain"
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command20 
      Caption         =   "GestureDown"
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command19 
      Caption         =   "GestureLeft"
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command18 
      Caption         =   "GestureRight"
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command17 
      Caption         =   "GestureUp"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      Caption         =   "GetAttention"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "GetAttentionReturn"
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Greet"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "LookDown"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "LookDownBlink"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Dont Recognize"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Magic2"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Magic1"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Decline"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Congratulate_2"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Congratulate"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Confused"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Blink"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Announce"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Alert"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "if you don't have Microsoft Agent download it from:"
      Height          =   375
      Left            =   0
      TabIndex        =   62
      Top             =   7440
      Width           =   9855
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   1800
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "MADE BY Wizard Connect me by wizard2004@coolmail.co.il "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   10575
   End
   Begin VB.Label Label2 
      Caption         =   "What To do?"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Agent1.Characters.Character("Merlin").Play "Acknowledge"
End Sub

Private Sub Command10_Click()
    Agent1.Characters.Character("Merlin").Play "DoMagic2"
End Sub

Private Sub Command11_Click()
    Agent1.Characters.Character("Merlin").Play "DontRecognize"
End Sub

Private Sub Command12_Click()
Agent1.Characters.Character("Merlin").Play "LookDownBlink"
End Sub

Private Sub Command13_Click()
Agent1.Characters.Character("Merlin").Play "LookDown"
End Sub

Private Sub Command14_Click()
Agent1.Characters.Character("Merlin").Play "Greet"
End Sub

Private Sub Command15_Click()
    Agent1.Characters.Character("Merlin").Play "GetAttentionReturn"
End Sub

Private Sub Command16_Click()
    Agent1.Characters.Character("Merlin").Play "GetAttentionContinuedCongratulate_2"
End Sub

Private Sub Command17_Click()
    Agent1.Characters.Character("Merlin").Play "GestureUp"
End Sub

Private Sub Command18_Click()
    Agent1.Characters.Character("Merlin").Play "GestureRight"
End Sub

Private Sub Command19_Click()
    Agent1.Characters.Character("Merlin").Play "GestureLeft"
End Sub

Private Sub Command2_Click()
    Agent1.Characters.Character("Merlin").Play "Alert"
End Sub

Private Sub Command20_Click()
    Agent1.Characters.Character("Merlin").Play "GestureDown"
End Sub

Private Sub Command21_Click()
    Agent1.Characters.Character("Merlin").Play "Explain"
End Sub

Private Sub Command22_Click()
    Agent1.Characters.Character("Merlin").Play "LookDownReturn"
End Sub

Private Sub Command23_Click()
    Agent1.Characters.Character("Merlin").Play "MoveUp"
End Sub

Private Sub Command24_Click()
    Agent1.Characters.Character("Merlin").Play "LookLeft"
End Sub

Private Sub Command25_Click()
    Agent1.Characters.Character("Merlin").Play "LookLeftBlink"
End Sub

Private Sub Command26_Click()
    Agent1.Characters.Character("Merlin").Play "LookLeftReturn"
End Sub

Private Sub Command27_Click()
    Agent1.Characters.Character("Merlin").Play "LookRight"
End Sub

Private Sub Command28_Click()
    Agent1.Characters.Character("Merlin").Play "LookRightBlink"
End Sub

Private Sub Command29_Click()
    Agent1.Characters.Character("Merlin").Play "LookRightReturn"
End Sub

Private Sub Command3_Click()
    Agent1.Characters.Character("Merlin").Play "Announce"
End Sub

Private Sub Command30_Click()
    Agent1.Characters.Character("Merlin").Play "LookUp"
End Sub

Private Sub Command31_Click()
    Agent1.Characters.Character("Merlin").Play "LookUpBlink"
End Sub

Private Sub Command32_Click()
    Agent1.Characters.Character("Merlin").Play "LookUpReturn"
End Sub

Private Sub Command33_Click()
    Agent1.Characters.Character("Merlin").Play "MoveDown"
End Sub

Private Sub Command34_Click()
    Agent1.Characters.Character("Merlin").Play "StopListening"
End Sub

Private Sub Command35_Click()
    Agent1.Characters.Character("Merlin").Play "MoveLeft"
End Sub

Private Sub Command36_Click()
    Agent1.Characters.Character("Merlin").Play "MoveRight"
End Sub

Private Sub Command37_Click()
    Agent1.Characters.Character("Merlin").Play "Pleased"
End Sub

Private Sub Command38_Click()
    Agent1.Characters.Character("Merlin").Play "Process"
End Sub

Private Sub Command39_Click()
    Agent1.Characters.Character("Merlin").Play "Read"
End Sub

Private Sub Command4_Click()
    Agent1.Characters.Character("Merlin").Play "Blink"
End Sub

Private Sub Command40_Click()
    Agent1.Characters.Character("Merlin").Play "ReadContinued"
End Sub

Private Sub Command41_Click()
    Agent1.Characters.Character("Merlin").Play "Readheturn"
End Sub

Private Sub Command42_Click()
    Agent1.Characters.Character("Merlin").Play "Sad"
End Sub

Private Sub Command43_Click()
    Agent1.Characters.Character("Merlin").Play "Search"
End Sub

Private Sub Command44_Click()
    Agent1.Characters.Character("Merlin").Play "StartListening"
End Sub

Private Sub Command45_Click()
    Agent1.Characters.Character("Merlin").Stop
    Agent1.Characters.Character("Merlin").Speak Text1.Text
End Sub

Private Sub Command46_Click()
    Agent1.Characters.Character("Merlin").Play "Suggest"
End Sub

Private Sub Command47_Click()
    Agent1.Characters.Character("Merlin").Play "Surprised"
End Sub

Private Sub Command48_Click()
    Agent1.Characters.Character("Merlin").Play "Think"
End Sub

Private Sub Command49_Click()
    Agent1.Characters.Character("Merlin").Play "Uncertain"
End Sub

Private Sub Command5_Click()
    Agent1.Characters.Character("Merlin").Play "Confused"
End Sub

Private Sub Command50_Click()
    Agent1.Characters.Character("Merlin").Play "Wave"
End Sub

Private Sub Command51_Click()
    Agent1.Characters.Character("Merlin").Play "Write"
End Sub

Private Sub Command52_Click()
    Agent1.Characters.Character("Merlin").Play "WriteContinued"
End Sub

Private Sub Command53_Click()
    Agent1.Characters.Character("Merlin").Play "WriteReturn"
End Sub

Private Sub Command54_Click()
    Agent1.Characters.Character("Merlin").Play "Processing"
End Sub

Private Sub Command55_Click()
    Agent1.Characters.Character("Merlin").Play "Reading"
End Sub

Private Sub Command56_Click()
    Agent1.Characters.Character("Merlin").Play "Thinking"
End Sub

Private Sub Command57_Click()
On Error Resume Next
    Agent1.Characters.Character("Merlin").Hide
End Sub

Private Sub Command58_Click()
On Error Resume Next
    Agent1.Characters.Character("Merlin").Show
End Sub

Private Sub Command6_Click()
    Agent1.Characters.Character("Merlin").Play "Congratulate"
End Sub

Private Sub Command7_Click()
    Agent1.Characters.Character("Merlin").Play "Congratulate_2"
End Sub

Private Sub Command8_Click()
    Agent1.Characters.Character("Merlin").Play "Decline"
End Sub

Private Sub Command9_Click()
    Agent1.Characters.Character("Merlin").Play "DoMagic1"
End Sub

Private Sub Form_Load()
    Agent1.Characters.Load ("Merlin"), "Merlin.acs"
    Agent1.Characters.Character("Merlin").Show
End Sub

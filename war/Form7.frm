VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Communicate"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option6 
      Caption         =   "Option6"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option5"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "Player5"
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      ToolTipText     =   "Player4"
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      ToolTipText     =   "player3"
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "player2"
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "player1"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "You"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Used in multiplayer only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Put Text in the box"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Player"
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Call TEXT1
Unload Me

End Sub

Private Sub Option5_Click()

If Option5.Value = True Then


Call CHECKED
End If
End Sub
'jonathan valentin 1999

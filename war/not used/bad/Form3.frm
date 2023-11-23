VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Main Menu"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form3"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Level Editor"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load Saved Game"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Scernio"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Multiplayer"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Campain"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -840
      Picture         =   "Form3.frx":0000
      Top             =   -2040
      Width           =   12000
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'Dim index As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
race.Show vbModal, Me
Unload Me

End Sub

Private Sub Command2_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
End Sub

Private Sub Command3_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
End Sub

Private Sub Command4_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
CommonDialog1.ShowOpen
End Sub

Private Sub Command5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
End
End Sub

Private Sub Command6_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
Form1.Show vbModal, Me

End Sub

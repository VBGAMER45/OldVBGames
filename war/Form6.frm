VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   Caption         =   "Level"
   ClientHeight    =   3750
   ClientLeft      =   3975
   ClientTop       =   2400
   ClientWidth     =   2475
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   2475
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Status"
      Top             =   3480
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/28/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:20 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Ok"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton Option3 
         Caption         =   "Hard"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         ToolTipText     =   "Hard Is Hard"
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mediuim"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         ToolTipText     =   "Meduium"
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Easy"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         ToolTipText     =   "Easy"
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Unload Me
race.Show

End Sub

Private Sub Option1_Click()
If vbChecked Then
Command1.Enabled = True
End If
End Sub

Private Sub Option2_Click()
If vbChecked Then
Command1.Enabled = True
End If
End Sub

Private Sub Option3_Click()
If vbChecked Then
Command1.Enabled = True
End If
End Sub

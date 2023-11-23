VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form race 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Race"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "race.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   3840
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5520
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Choose a race"
      Top             =   3990
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "2/5/00"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "2:52 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Choose your race"
            TextSave        =   "Choose your race"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image7 
      Height          =   1080
      Left            =   3600
      Picture         =   "race.frx":030A
      Top             =   2040
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image4 
      Height          =   1290
      Left            =   4200
      Picture         =   "race.frx":250C
      Top             =   960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "United States"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GERMANY"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forgotton"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alien"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Soviet"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   6720
      X2              =   0
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image Image6 
      Height          =   765
      Left            =   1800
      Picture         =   "race.frx":8A16
      ToolTipText     =   "Be us because of are space travel and all the cool tech stuff help us defeat Earth!"
      Top             =   2160
      Width           =   915
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4680
      Picture         =   "race.frx":9B18
      ToolTipText     =   "We are part machine and part human made by expermints of the USA help was defeat them and you will be rewarded"
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   960
      Picture         =   "race.frx":A3E2
      ToolTipText     =   "Be the Soviets and defeat hitler and expand the Soviet Union"
      Top             =   360
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3240
      Picture         =   "race.frx":BF0C
      ToolTipText     =   "Be us so you can see what would happen!!"
      Top             =   480
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5400
      Picture         =   "race.frx":C216
      ToolTipText     =   "Be us because are wepons"
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "race"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'Dim index As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Private Sub Command1_Click()
'
'Unload Me
Form3.Show
End Sub

Private Sub Form_DblClick()
i = sndPlaySound(ByVal CStr(App.Path & "\comandr.wav"), SND_ASYNC)
End Sub



Private Sub Image1_Click()
Form9.Label3.Caption = InputBox("Enter Your Name:", "Name")
Form9.Label4.Visible = True
Form9.Show
Unload Me
End Sub

Private Sub Image2_Click()
Form9.Label3.Caption = InputBox("Enter Your Name:", "Name")
Form9.Label8.Visible = True

Form9.Show
Unload Me
End Sub
Private Sub Image3_Click()
 Form9.Label3.Caption = InputBox("Enter Your Name:", "Name")
 Form9.Label5.Visible = True
 Form9.Image6.Visible = True
 Form9.Image5.Visible = True
 Form9.Show
 Unload Me
 
End Sub
Private Sub Image5_Click()

Form9.Label3.Caption = InputBox("Enter Your Name:", "Name")
Form9.Label6.Visible = True
Form9.Image4.Visible = True
Form9.Show
Unload Me

End Sub
Private Sub Image6_Click()
Form9.Label3.Caption = InputBox("Enter Your Name:", "Name")
Form9.Label7.Visible = True
Form9.Image3.Visible = True
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
Form9.Show
Unload Me

End Sub

'jonathan valentin 1999
Private Sub Label4_Click()

End Sub

Private Sub Timer1_Timer()
Image4.Visible = True
Image7.Visible = True
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Image4.Visible = False
Image7.Visible = False
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

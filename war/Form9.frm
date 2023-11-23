VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mission"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   720
      Picture         =   "Form9.frx":030A
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4080
      Picture         =   "Form9.frx":0614
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   1320
      Left            =   1560
      Picture         =   "Form9.frx":091E
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   4320
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2280
      Picture         =   "Form9.frx":52C0
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your mission is to destroy all the sam sites and kill the humans."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your mission is to destroy three gap generators"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your mission is to kill the queen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your mission is to destroy all enemy troops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your mission is to destroy the enemys communication center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MISSION 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   1200
      Picture         =   "Form9.frx":55CA
      ToolTipText     =   "Enter"
      Top             =   3000
      Width           =   2775
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'Dim index As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long




Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Image1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\CLASIFD.wav"), SND_ASYNC)
'MsgBox "if your foggoton close first sceen"
Call pic
Call hi
Call cool
Call sov
Call mis
'Call alien
'Dim form2 As New form2
'i = sndPlaySound(ByVal CStr(App.Path & "\CLASIFD.wav"), SND_ASYNC)
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
'form2.Show vbModal, Me
'If Label7.Visible = True Then
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")                         '("c:\windows\desktop\vb projects\war\com.bmp")
'End If
End Sub
Public Sub pic()
If Label7.Visible = True Then
'Call jv
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
Dim form2 As New form2

form2.Image33.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
form2.Image24.Visible = False
form2.Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viechels\alein\commabnfd.bmp")
form2.Picture2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viechels\alein\o.bmp")
form2.Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\w1.bmp")
form2.Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\q6.bmp")
form2.Picture4.Picture = LoadPicture("c:\windows\desktop\vb projects\war\\alien\q1.bmp")
form2.Picture7.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\q.bmp")
form2.Picture18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\q.bmp")
form2.Picture10.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viechels\alein\odd.bmp")
form2.Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\w3.bmp")
form2.Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\q4.bmp")
form2.Image155.Visible = True
'form2.Picture1.DragMode = 1
'form2.Picture1.OLEDropMode = 2
'form2.Picture1.OLEDragMode = 1
'form2.Picture2.DragMode = 1
'form2.Picture2.OLEDropMode = 2
'form2.Picture2.OLEDragMode = 1
'form2.Image3.DragMode = 1
form2.Image3.OLEDropMode = 2
'form2.Image3.OLEDragMode = 1
form2.Timer16.Enabled = True

form2.Show
'Else
'form2.Show vbModal, Me

End If
'If Label6.Visible = True Then
'Call jv
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
'Dim form2 As New form2

'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back2.bmp")
'form2.Show vbModal, Me
'Else
'form2.Show vbModal, Me

'End If
End Sub
Public Sub sov()
If Label5.Visible = True Then
form2.Label5.Caption = "1a"
form2.Show
End If

End Sub
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
Public Sub hi()
If Label6.Visible = True Then
'Call jv
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
'Dim form2 As New form2
form2.Image24.Visible = False
form2.Image33.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back2.bmp")
form2.Image2.Visible = False
form2.Image14.Visible = False
form2.Image29.Visible = True
'form2.Image2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viechels\comm.bmp")
'form2.Image2.Visible = False
form2.Show
'Else
'form2.Show vbModal, Me

End If
End Sub
Public Sub alien()
'If Label7.Visible = True Then
'form2.Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viechels\alein\commabnfd.bmp")
'form2.Picture2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viecheks\alien\o.bmp")
'form2.Picture1.DragMode = 1
'form2.Picture1.OLEDropMode = 2
'form2.Picture1.OLEDragMode = 2
'form2.Picture2.DragMode = 1
'form2.Picture2.OLEDropMode = 2
'form2.Picture2.OLEDragMode = 2
'form2.Image9.DragMode = 1
'form2.Image9.OLEDropMode = 2
'form2.Image9.OLEDragMode = 2
'form2.Show vbModal, Me

'End If
End Sub
Public Sub cool()
If Label4.Visible = True Then
form2.Image142.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t1.bmp")
form2.Image96.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t3.bmp")
form2.Image96.OLEDropMode = 0
form2.Image142.OLEDropMode = 0
form2.Label5.Caption = "1c"
form2.Image33.Picture = LoadPicture("c:\windows\desktop\vb projects\war\jcvxs.bmp")
form2.Show
End If
End Sub
'jonathan valentin 1999
Public Sub jv()
form2.Image33.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
End Sub
Public Sub mis()
If Label1.Caption = "Mission 2" Then
Dim form2 As New form2
form2.Image33.Picture = LoadPicture("c:\windows\desktop\vb projects\war\desertback.bmp")
form2.Image120.Picture = LoadPicture("c:\windows\desktop\vb projects\war\bush.bmp")
form2.Image91.Picture = LoadPicture("c:\windows\desktop\vb projects\war\bush1.bmp")
form2.Image54.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cactus.bmp")
form2.Image120.OLEDropMode = 0
form2.Label5.Caption = "2a"
form2.Image91.OLEDropMode = 0
form2.Image54.OLEDropMode = 0
form2.Show

End If

End Sub
'Private Sub Marquee1_OnStartOfImage()
'Marquee1.DrawImmediately = load("c:\windows\desktop\vb projects\war\anti.bmp")

'End Sub

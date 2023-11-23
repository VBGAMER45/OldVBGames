VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Dimploacy"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   ForeColor       =   &H00404040&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1550
      Left            =   2040
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   3000
      Picture         =   "Form4.frx":030A
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   39
      ToolTipText     =   "Next age"
      Top             =   3960
      Width           =   735
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   2400
      TabIndex        =   24
      Top             =   480
      Width           =   3015
      Begin VB.OptionButton Option7 
         Caption         =   "Option7"
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   2400
      TabIndex        =   23
      Top             =   1080
      Width           =   3015
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   195
         Left            =   2400
         TabIndex        =   27
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   2400
      TabIndex        =   22
      Top             =   2520
      Width           =   3015
      Begin VB.OptionButton Option11 
         Caption         =   "Option11"
         Height          =   195
         Left            =   2400
         TabIndex        =   33
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Option9"
         Height          =   195
         Left            =   1440
         TabIndex        =   32
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option8"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2400
      TabIndex        =   21
      Top             =   3240
      Width           =   3015
      Begin VB.OptionButton Option14 
         Caption         =   "Option14"
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Option13"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Option12"
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2400
      TabIndex        =   17
      Top             =   1800
      Width           =   3015
      Begin VB.OptionButton Option10 
         Caption         =   "Option10"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Option16"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2040
      TabIndex        =   12
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Text            =   "0"
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Advance to next age"
      Height          =   495
      Left            =   1920
      TabIndex        =   38
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label11 
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3240
      TabIndex        =   37
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "Offer allies or enemy money and click check box"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Ally"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Neteual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Enemy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Check5_Click()
If vbChecked Then

MsgBox "" = Text1.Text
End If

End Sub

Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
' form2.money.Caption = -Text1.Text   '-""
'form2.money.Caption = -Text1.Text
form2.money.Caption = form2.money.Caption - Text1.Text
'If Option7 = vbChecked Then
form2.Label1.Visible = True
If Text1.Text >= 4000 Then
form2.Label1.Caption = "You are my ally"
End If
If Text1.Text <= 4000 Then
form2.Label1.Caption = "You give pitty gold ,enemy"
End If
form2.Label6.Caption = Label11.Caption

Unload Me


'vbUnchecked '0   Unchecked
'vbChecked
End Sub

Private Sub Form_Load()
Label11.Caption = Int(Rnd * 10)
'Label12.Caption = Int(Rnd * 10)
'Label13.Caption = Int(Rnd * 10)
'Randomize
If Label11.Caption = 1 Then

Label5.Caption = "greek"
Label6.Caption = "alliance"
Label7.Caption = "france"
Label8.Caption = "federation"
Label9.Caption = "japen"
End If
If Label11.Caption = 2 Then

Label5.Caption = "gremany"
Label6.Caption = "china"
Label7.Caption = "france"
Label8.Caption = "federation"
Label9.Caption = "zerg"
End If
If Label11.Caption = 3 Then

Label5.Caption = "Mars Attacks"
Label6.Caption = "alliance"
Label7.Caption = "Cuba"
Label8.Caption = "Hittie"
Label9.Caption = "Yamato"
End If
If Label11.Caption = 4 Then

Label5.Caption = "England"
Label6.Caption = "Rome"
Label7.Caption = "Egypt"
Label8.Caption = "India"
Label9.Caption = "japen"
End If
If Label11.Caption = 5 Then

Label5.Caption = "Alien"
Label6.Caption = "Alantis"
Label7.Caption = "Russia"
Label8.Caption = "Unknown"
Label9.Caption = "Rebels"
End If
If Label11.Caption = 6 Then

Label5.Caption = "greek"
Label6.Caption = "Maya"
Label7.Caption = "france"
Label8.Caption = "Cybrogs"
Label9.Caption = "Inca"
End If
If Label11.Caption = 7 Then

Label5.Caption = "Persain"
Label6.Caption = "?"
Label7.Caption = "Choson"
Label8.Caption = "Forgotton"
Label9.Caption = "Spain"
End If
If Label11.Caption = 8 Then

Label5.Caption = "Irag"
Label6.Caption = "enemy1"
Label7.Caption = "Enemy3"
Label8.Caption = "Rebels"
Label9.Caption = "Mutants"
End If
If Label11.Caption = 9 Then

Label5.Caption = "greek"
Label6.Caption = "alliance"
Label7.Caption = "france"
Label8.Caption = "federation"
Label9.Caption = "japen"
End If
If Label11.Caption = 0 Then

Label5.Caption = "Sumarians"
Label6.Caption = "Shang"
Label7.Caption = "Mongels"
Label8.Caption = "Inuit"
Label9.Caption = "Master"
End If
Randomize
End Sub

'jonathan valentin 1999
Private Sub Option7_Click()
'If vbChecked Then
form2.Label1.Visible = True
'End If
End Sub

Private Sub Picture1_Click()
form2.money.Caption = form2.money.Caption - 1000

'Timer1.Enabled = True
form2.Timer5.Enabled = True
End Sub

Private Sub Timer1_Timer()
'Label13.Caption = Label13.Caption + 1
'form2.Label7.Caption = form2.Label7.Caption + 1
End Sub

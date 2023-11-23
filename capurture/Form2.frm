VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Capture The Flag!!"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11850
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6780
   ScaleWidth      =   11850
   Begin VB.Timer Timer12 
      Interval        =   1
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6360
      Top             =   840
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Interval        =   20
      Left            =   2400
      Top             =   360
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   10200
      Top             =   1080
   End
   Begin VB.Timer Timer7 
      Interval        =   20000
      Left            =   10680
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9120
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   840
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   1800
      Top             =   720
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1320
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1680
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   720
   End
   Begin VB.Image Image11 
      Height          =   60
      Left            =   1800
      Picture         =   "Form2.frx":030A
      Top             =   1320
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label11 
      Height          =   135
      Left            =   9120
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "gun"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   720
      TabIndex        =   8
      ToolTipText     =   "gun"
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   270
      Left            =   7080
      Picture         =   "Form2.frx":039C
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image9 
      Height          =   195
      Left            =   10920
      Picture         =   "Form2.frx":0816
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label7 
      Caption         =   "7680"
      Height          =   255
      Left            =   9960
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   5
      ToolTipText     =   "Team Two Score!"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10080
      TabIndex        =   4
      ToolTipText     =   "Team 1 Score"
      Top             =   480
      Width           =   855
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   480
      Picture         =   "Form2.frx":0B98
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   7200
      Picture         =   "Form2.frx":0EA2
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   6720
      Picture         =   "Form2.frx":128C
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      ToolTipText     =   "Player 4 Time Left before your out!"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "Player 3 Time Left before your out!"
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Player 2 Time Left before your out!"
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Player 1 Time Left before your out!"
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3360
      Picture         =   "Form2.frx":1676
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   3360
      Picture         =   "Form2.frx":1A3C
      Top             =   2280
      Width           =   180
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   11040
      Picture         =   "Form2.frx":1E02
      Top             =   6120
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   6240
      X2              =   6240
      Y1              =   0
      Y2              =   8280
   End
   Begin VB.Image Image8 
      Height          =   7200
      Left            =   -480
      Picture         =   "Form2.frx":210C
      Top             =   0
      Width           =   9600
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   3960
      Picture         =   "Form2.frx":E314E
      Top             =   0
      Width           =   9600
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnunew 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnukeys 
         Caption         =   "Keys"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response As String

Private Sub Command1_Click()
Label7.Caption = Label7.Caption + 480

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then
If Image2.Left < 11520 Then
Image2.Left = Image2.Left + 480
End If
Image2.Picture = LoadPicture(App.Path & "\man2.bmp")
If Image2.Left = Image7.Left And Image2.Top = Image7.Top Then
Image7.Left = Image7.Left - 480

Image7.Visible = True
End If
End If


Image2.ToolTipText = Image2.Left

If KeyCode = vbKeyLeft Then
'If Image2.Left >= 0 Then
If Image2.Left > 0 Then
Image2.Left = Image2.Left - 480
End If

Image2.Picture = LoadPicture(App.Path & "\man1.bmp")
If Image2.Left = Image7.Left And Image2.Top = Image7.Top Then
Image7.Left = Image7.Left - 480

Image7.Visible = True
End If


End If



If KeyCode = vbKeySpace Then

If Label8.Caption = "r" Then
'Image11.Left = Image11.Left - 480
Timer11.Enabled = True
End If
If Label9.Caption = "r" Then
Timer11.Enabled = True
'Image11.Left = Image11.Left - 480
End If
If Label10.Caption = "r" Then
'Image11.Left = Image11.Left - 480
Timer11.Enabled = True
End If
If Label1.Caption = "r" Then
'Image11.Left = Image11.Left - 480
Timer11.Enabled = True
End If


End If



If KeyCode = vbKeyUp Then
Image2.ToolTipText = Image2.Top
If Image2.Top > 840 Then
Image2.Top = Image2.Top - 480
End If
If Image2.Left = Image7.Left And Image2.Top = Image7.Top Then
Image7.Top = Image7.Top - 480

Image7.Visible = True
End If
End If
If KeyCode = vbKeyDown Then
If Image2.Top < 8040 Then
Image2.Top = Image2.Top + 480
End If
If Image2.Left = Image7.Left And Image2.Top = Image7.Top Then
Image7.Top = Image7.Top + 480

Image7.Visible = True
End If
End If
If KeyCode = vbKeyNumpad6 Then
If Image5.Left < 11520 Then
Image5.Left = Image5.Left + 480
End If
Image5.Picture = LoadPicture(App.Path & "\man8.bmp")
If Image5.Left = Image7.Left And Image5.Top = Image7.Top Then
Image7.Left = Image7.Left - 480
Image7.Visible = True
End If
End If
If KeyCode = vbKeyNumpad4 Then

If Image5.Left > 0 Then
Image5.Left = Image5.Left - 480
End If
Image5.Picture = LoadPicture(App.Path & "\man7.bmp")
If Image5.Left = Image7.Left And Image5.Top = Image7.Top Then
Image7.Left = Image7.Left - 480
Image7.Visible = True
End If
End If


If KeyCode = vbKeyNumpad8 Then
If Image5.Top > 840 Then
Image5.Top = Image5.Top - 480
End If
If Image5.Left = Image7.Left And Image5.Top = Image7.Top Then
Image7.Top = Image7.Top - 480
Image7.Visible = True
End If
End If
If KeyCode = vbKeyNumpad2 Then
If Image5.Top < 8040 Then
Image5.Top = Image5.Top + 480
End If
If Image5.Left = Image7.Left And Image5.Top = Image7.Top Then
Image7.Top = Image7.Top + 480
Image7.Visible = True
End If
End If
If KeyCode = vbKeyW Then
If Image3.Top > 840 Then
Image3.Top = Image3.Top - 480
End If
If Image3.Left = Image6.Left And Image3.Top = Image6.Top Then
Image6.Visible = True
Image6.Top = Image6.Top - 480
End If
End If
If KeyCode = vbKeyA Then
If Image3.Left > 0 Then
Image3.Left = Image3.Left - 480
End If
Image3.Picture = LoadPicture(App.Path & "\man3.bmp")
If Image3.Left = Image6.Left And Image3.Top = Image6.Top Then
Image6.Visible = True
Image6.Left = Image6.Left + 480
End If
End If

If KeyCode = vbKeyS Then
If Image3.Top < 8040 Then
Image3.Top = Image3.Top + 480
End If
If Image3.Left = Image6.Left And Image3.Top = Image6.Top Then
Image6.Visible = True
Image6.Top = Image6.Top + 480
End If
End If
If KeyCode = vbKeyD Then
If Image3.Left < 11520 Then
Image3.Left = Image3.Left + 480
End If
Image3.Picture = LoadPicture(App.Path & "\man4.bmp")
If Image3.Left = Image6.Left And Image3.Top = Image6.Top Then
Image6.Visible = True

Image6.Left = Image6.Left + 480
End If

End If
If KeyCode = vbKeyI Then
If Image4.Top > 840 Then
Image4.Top = Image4.Top - 480
End If
If Image4.Left = Image6.Left And Image4.Top = Image6.Top Then
Image6.Visible = True

Image6.Left = Image6.Left + 480
End If
End If
If KeyCode = vbKeyJ Then
If Image4.Left > 0 Then
Image4.Left = Image4.Left - 480
End If
Image4.Picture = LoadPicture(App.Path & "\man5.bmp")
If Image4.Left = Image6.Left And Image4.Top = Image6.Top Then
Image6.Visible = True

Image6.Left = Image6.Left + 480

End If

End If
If KeyCode = vbKeyL Then
'If Image4.Left >= 0 Then
If Image4.Left < 11520 Then
Image4.Left = Image4.Left + 480
End If
Image4.Picture = LoadPicture(App.Path & "\man6.bmp")
'End If
If Image4.Left = Image6.Left And Image4.Top = Image6.Top Then
Image6.Visible = True

Image6.Left = Image6.Left + 480
End If
End If
If KeyCode = vbKeyK Then
If Image4.Top < 8040 Then
Image4.Top = Image4.Top + 480
End If
If Image4.Left = Image6.Left And Image4.Top = Image6.Top Then
Image6.Visible = True

Image6.Left = Image6.Left + 480
End If
End If
End Sub

Private Sub Form_Load()
Form2.KeyPreview = True
Image8.Width = 7200
  Midi = "uyworld.mid"
    CmdStr = "play " + App.Path + "\" & Midi
    ReturnVal = mciSendString(CmdStr, 0&, 0, 0&)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CmdStr = "stop " + App.Path + "\" & Midi
    ReturnVal = mciSendString(CmdStr, 0&, 0, 0&)
End Sub

Private Sub mnuabout_Click()
MsgBox "Jonathan Valentin Jan 3,2000 Fixed March 9,2001", vbExclamation

End Sub

Private Sub mnuexit_Click()
response = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit!!!")

If response = vbYes Then
    CmdStr = "stop " + App.Path + "\" & Midi
    ReturnVal = mciSendString(CmdStr, 0&, 0, 0&)

End
End If
End Sub

Private Sub mnukeys_Click()
MsgBox "The keys are player1 w,a,s,d player2, j,k,i,l player3 up, down, right,left, player4, numpad arrows", vbInformation

End Sub

Private Sub mnunew_Click()
    CmdStr = "stop " + App.Path + "\" & Midi
    ReturnVal = mciSendString(CmdStr, 0&, 0, 0&)
Unload Me

Form1.Show
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Label2.Caption + 1
Image3.Top = 1800
Image3.Left = 6240
Image3.Enabled = False
If Label2.Caption > 20 Then
Image3.Visible = True
Image3.Left = 6240
Image3.Enabled = True
Label2.Caption = 0
Timer1.Enabled = False
End If
End Sub

Private Sub Timer10_Timer()
If Image11.Left <= 0 Then
If Label8.Caption = "r" Then
Image11.Left = Image2.Left
Image11.Top = Image2.Top
Image11.Visible = True
End If
If Label9.Caption = "r" Then
Image11.Left = Image3.Left
Image11.Top = Image3.Top
Image11.Visible = True
End If
If Label10.Caption = "r" Then
Image11.Left = Image5.Left
Image11.Top = Image5.Top
Image11.Visible = True
End If
If Label1.Caption = "r" Then
Image11.Left = Image9.Left
Image11.Top = Image9.Top
Image11.Visible = True

End If
End If

End Sub

Private Sub Timer11_Timer()
Image11.Left = Image11.Left - 480

End Sub

Private Sub Timer12_Timer()
If Image6.Left >= 6240 Then
MsgBox "Team2 got the flag!", vbExclamation, "Team2 Wins"
Image6.Left = 480
Image6.Top = 3720
Label6.Caption = Label6.Caption + 1
Timer6.Enabled = True
End If
If Image7.Left <= 6240 Then
MsgBox "Team1 got the flag", vbExclamation, "Team1 Wins"
Image7.Left = 11040
Image7.Top = 6120
Label5.Caption = Label5.Caption + 1
Image2.Left = 3360
Timer6.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Label3.Caption = Label3.Caption + 1
Image4.Top = 1800
Image4.Left = 6240
Image4.Enabled = False
If Label3.Caption > 20 Then
Image4.Visible = True
Image4.Left = 6240
Image4.Enabled = True
Label3.Caption = 0
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
Label1.Caption = Label1.Caption + 1
Image2.Top = 1800
Image2.Left = 3360
Image2.Enabled = False
If Label1.Caption > 20 Then
Image2.Visible = True
Image2.Left = 3360

Label1.Caption = 0
Image2.Enabled = True
Timer3.Enabled = False
End If

End Sub

Private Sub Timer4_Timer()
Label4.Caption = Label4.Caption + 1
Image4.Top = 1800
Image5.Left = 3360
Image5.Enabled = False
If Label4.Caption > 20 Then
Image5.Visible = True
Image5.Enabled = True
Label4.Caption = 0
Timer4.Enabled = False
End If

End Sub

Private Sub Timer5_Timer()
If Image2.Left = Image3.Left And Image2.Top = Image3.Top And Image3.Left <= 6220 Then
Timer1.Enabled = True
Image3.Visible = False
End If
If Image2.Left = Image4.Left And Image2.Top = Image4.Top And Image4.Left <= 6220 Then
Timer2.Enabled = True
Image4.Visible = False
End If
If Image5.Left = Image3.Left And Image5.Top = Image3.Top And Image3.Left <= 6220 Then
Timer1.Enabled = True
Image3.Visible = False
End If
If Image5.Left = Image4.Left And Image5.Top = Image4.Top And Image4.Left <= 6220 Then
Timer2.Enabled = True
Image4.Visible = False
End If
If Image3.Left = Image5.Left And Image3.Top = Image5.Top And Image5.Left >= 6220 Then
Timer4.Enabled = True
Image5.Visible = False
End If
If Image3.Left = Image2.Left And Image3.Top = Image2.Top And Image2.Left >= 6220 Then
Timer3.Enabled = True
Image2.Visible = False
End If
If Image4Left = Image5.Left And Image4.Top = Image5.Top And Image5.Left >= 6220 Then
Timer4.Enabled = True
Image5.Visible = False
End If
If Image4.Left = Image2.Left And Image4.Top = Image2.Top And Image2.Left >= 6220 Then
Timer3.Enabled = True
Image2.Visible = False
End If
End Sub

Private Sub Timer6_Timer()
'If Image3.Left = Image6.Left And Image3.Top = Image6.Top Then

'End If
'If Image4.Left = Image6.Left And Image4.Top = Image6.Top Then

'End If
'If Image2.Left = Image7.Left And Image2.Top = Image7.Top Then

'End If
'If Image5.Left = Image7.Left And Image5.Top = Image7.Top Then

'End If
'If Image6.Left >= 6240 Then
'MsgBox "team2"
'Image6.Left = 480
'Image6.Top = 3720
'Label6.Caption = Label6.Caption + 1

'End If
'If Image7.Left <= 6240 Then
'MsgBox "team1"
'Image7.Left = 11040
'Image7.Top = 6120
'Label5.Caption = Label5.Caption + 1
'image2.Left = 3360

'End If
If Label6.Caption = 2 Then
Image6.Top = 6120

End If
If Label5.Caption = 2 Then
Image7.Top = 3720

End If
If Label6.Caption = 4 Then
Image6.Top = 2280


End If
If Label5.Caption = 4 Then
Image7.Top = 5160

End If
If Label6.Caption = 10 Then
Image6.Top = 2280
Image6.Left = 3360


End If
If Label5.Caption = 10 Then
Image7.Top = 5160
Image7.Left = 7200
End If
If Label6.Caption = 15 Then
Image6.Top = 2280
Image6.Left = 3360


End If
If Label5.Caption = 15 Then
Image7.Top = 5160

Image7.Left = 7200
End If
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
Image9.Top = 5640
Image9.Left = 6720
Image9.Visible = True
Image9.Enabled = True
Timer8.Enabled = True
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
Image9.Top = 6120
Image9.Left = 3360
Image9.Visible = True
Image9.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = False
End Sub

Private Sub Timer9_Timer()
If Image2.Left = Image9.Left And Image2.Top = Image9.Top Then
Label8.Caption = "r"
Image9.Visible = False
Image11.Top = Image2.Top
Image11.Left = Image2.Left
End If
If Image5.Left = Image9.Left And Image5.Top = Image9.Top Then
Label11.Caption = "r"
Image9.Visible = False
Image11.Top = Image5.Top
Image11.Left = Image5.Left
End If
If Image3.Left = Image9.Left And Image3.Top = Image9.Top Then
Label9.Caption = "r"
Image9.Visible = False
Image11.Top = Image3.Top
Image11.Left = Image3.Left
End If
If Image4.Left = Image9.Left And Image4.Top = Image9.Top Then
Label10.Caption = "r"
Image9.Visible = False
Image11.Top = Image4.Top
Image11.Left = Image4.Left
End If
End Sub

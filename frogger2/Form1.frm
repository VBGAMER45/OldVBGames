VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Frogger 2!  The Frog Returns!"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   11160
      Top             =   5400
   End
   Begin VB.Timer GotFlag 
      Interval        =   10
      Left            =   1560
      Top             =   3240
   End
   Begin VB.Timer Water 
      Interval        =   1
      Left            =   7560
      Top             =   3240
   End
   Begin VB.Timer Wait 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   6840
   End
   Begin VB.Timer Colllide 
      Interval        =   1
      Left            =   1920
      Top             =   7560
   End
   Begin VB.Timer Traffic 
      Interval        =   200
      Left            =   1680
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   2400
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      Height          =   1575
      Left            =   6480
      ScaleHeight     =   1515
      ScaleWidth      =   5340
      TabIndex        =   5
      Top             =   0
      Width           =   5400
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Level 
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   4200
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.Label Lives 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1080
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lives:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Image Life 
      Height          =   330
      Left            =   1080
      Picture         =   "Form1.frx":030A
      Top             =   7800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image B6 
      Height          =   300
      Left            =   10680
      Picture         =   "Form1.frx":097C
      Top             =   6120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image B5 
      Height          =   300
      Left            =   10920
      Picture         =   "Form1.frx":0EE9
      Top             =   7080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image B4 
      Height          =   300
      Left            =   10920
      Picture         =   "Form1.frx":1340
      Top             =   7560
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image B3 
      Height          =   300
      Left            =   10800
      Picture         =   "Form1.frx":17D0
      Top             =   6600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image B2 
      Height          =   300
      Left            =   9600
      Picture         =   "Form1.frx":1C4F
      Top             =   6720
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image B1 
      Height          =   225
      Left            =   9960
      Picture         =   "Form1.frx":2064
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "200"
      Height          =   135
      Left            =   360
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   5520
      Picture         =   "Form1.frx":2480
      Top             =   6960
      Width           =   660
   End
   Begin VB.Image Frog 
      Height          =   645
      Left            =   120
      Picture         =   "Form1.frx":28F1
      Top             =   6600
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image f5 
      Height          =   735
      Left            =   8880
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image f4 
      Height          =   735
      Left            =   7080
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image f3 
      Height          =   735
      Left            =   5040
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image f2 
      Height          =   615
      Left            =   3360
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image f1 
      Height          =   615
      Left            =   1320
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image s5 
      Height          =   375
      Left            =   13500
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Image s4 
      Height          =   855
      Left            =   8520
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Image s3 
      Height          =   735
      Left            =   5880
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image s2 
      Height          =   855
      Left            =   3600
      Top             =   2640
      Width           =   975
   End
   Begin VB.Image s1 
      Height          =   735
      Left            =   2160
      Top             =   2640
      Width           =   735
   End
   Begin VB.Image c5 
      Height          =   735
      Left            =   13500
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image c4 
      Height          =   855
      Left            =   8040
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image c3 
      Height          =   855
      Left            =   5640
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Image c2 
      Height          =   855
      Left            =   3360
      Top             =   4080
      Width           =   975
   End
   Begin VB.Image c1 
      Height          =   735
      Left            =   1440
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image dead 
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":2D62
      Top             =   7440
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image L5 
      Height          =   495
      Left            =   14760
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image L4 
      Height          =   495
      Left            =   9240
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image L3 
      Height          =   495
      Left            =   6480
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Image L2 
      Height          =   495
      Left            =   4560
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Image L1 
      Height          =   495
      Left            =   2400
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image Car6 
      Height          =   630
      Left            =   240
      Picture         =   "Form1.frx":337D
      Top             =   5640
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image Car5 
      Height          =   630
      Left            =   120
      Picture         =   "Form1.frx":38D1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image Car4 
      Height          =   585
      Left            =   120
      Picture         =   "Form1.frx":3FCB
      Top             =   4080
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Car3 
      Height          =   600
      Left            =   120
      Picture         =   "Form1.frx":44DC
      Top             =   3360
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Car2 
      Height          =   600
      Left            =   0
      Picture         =   "Form1.frx":4A9C
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Car1 
      Height          =   615
      Left            =   240
      Picture         =   "Form1.frx":4FC0
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Flag 
      Height          =   645
      Left            =   5520
      Picture         =   "Form1.frx":5517
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image E1 
      Height          =   270
      Left            =   2280
      Picture         =   "Form1.frx":5821
      Top             =   6960
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image E2 
      Height          =   300
      Left            =   4320
      Picture         =   "Form1.frx":5EDB
      Top             =   6960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image E3 
      Height          =   270
      Left            =   6840
      Picture         =   "Form1.frx":65AD
      Top             =   6960
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image E4 
      Height          =   300
      Left            =   8880
      Picture         =   "Form1.frx":6C67
      Top             =   6960
      Visible         =   0   'False
      Width           =   420
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
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnueasy 
         Caption         =   "Easy"
      End
      Begin VB.Menu mnumedium 
         Caption         =   "Medium"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuhard 
         Caption         =   "Hard"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String
Const MB_ICONASTERISK = &H40 'SystemAsterisk sound.
Const MB_ICONEXCLAMATION = &H30 'SystemExclamation sound.
Const MB_ICONHAND = &H10 'SystemHand sound.
Const MB_ICONQUESTION = &H20 'SystemQuestion sound.
Const MB_OK = &H0 'SystemDefault sound.
Private Declare Function MessageBeep Lib "user32.dll" (ByVal wType As Long) As Long
Private Sub Colllide_Timer()
If Image1.Top = L1.Top And Image1.Left >= L1.Left And Image1.Left <= L1.Left + L1.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = L2.Top And Image1.Left >= L2.Left And Image1.Left <= L2.Left + L2.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = L3.Top And Image1.Left >= L3.Left And Image1.Left <= L3.Left + L3.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = L4.Top And Image1.Left >= L4.Left And Image1.Left <= L4.Left + L4.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = L5.Top And Image1.Left >= L5.Left And Image1.Left <= L5.Left + L5.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = c1.Top And Image1.Left >= c1.Left And Image1.Left <= c1.Left + c1.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = c2.Top And Image1.Left >= c2.Left And Image1.Left <= c2.Left + c2.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = c3.Top And Image1.Left >= c3.Left And Image1.Left <= c3.Left + c3.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = c4.Top And Image1.Left >= c4.Left And Image1.Left <= c4.Left + c4.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = c5.Top And Image1.Left >= c5.Left And Image1.Left <= c5.Left + c5.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If



If Image1.Top = s1.Top And Image1.Left >= s1.Left And Image1.Left <= s1.Left + s1.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = s2.Top And Image1.Left >= s2.Left And Image1.Left <= s2.Left + s2.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = s3.Top And Image1.Left >= s3.Left And Image1.Left <= s3.Left + s3.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = s4.Top And Image1.Left >= s4.Left And Image1.Left <= s4.Left + s4.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = s5.Top And Image1.Left >= s5.Left And Image1.Left <= s5.Left + s5.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If

If Image1.Top = f5.Top And Image1.Left >= f5.Left And Image1.Left <= f5.Left + f5.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = f1.Top And Image1.Left >= f1.Left And Image1.Left <= f1.Left + f1.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = f2.Top And Image1.Left >= f2.Left And Image1.Left <= f2.Left + f2.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = f3.Top And Image1.Left >= f3.Left And Image1.Left <= f3.Left + f3.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If
If Image1.Top = f4.Top And Image1.Left >= f4.Left And Image1.Left <= f4.Left + f4.Width Then
Lives.Caption = Lives.Caption - 1
Image1.Picture = dead.Picture
Wait.Enabled = True
Colllide.Enabled = False

End If

Call GameOver

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
If Image1.Top > 16 Then
Image1.Top = Image1.Top - 32
End If

End If

If KeyCode = vbKeyDown Then
If Image1.Top < 496 Then
Image1.Top = Image1.Top + 32
End If

End If
If KeyCode = vbKeyLeft Then
If Image1.Left > 16 Then
Image1.Left = Image1.Left - 32
End If

End If

If KeyCode = vbKeyRight Then
If Image1.Left < 516 Then
Image1.Left = Image1.Left + 32
End If

End If


End Sub

Private Sub Form_Load()
L1.Picture = Car2.Picture
L2.Picture = Car3.Picture
L3.Picture = Car3.Picture
L4.Picture = Car2.Picture
L5.Picture = Car2.Picture
c1.Picture = Car1.Picture
c2.Picture = Car4.Picture
c3.Picture = Car5.Picture
c4.Picture = Car6.Picture
c5.Picture = Car4.Picture
s1.Picture = Car1.Picture
s2.Picture = Car5.Picture
s3.Picture = Car5.Picture
s4.Picture = Car6.Picture
s5.Picture = Car5.Picture

f1.Picture = Car3.Picture
f2.Picture = Car2.Picture
f3.Picture = Car2.Picture
f4.Picture = Car3.Picture
f5.Picture = Car3.Picture
End Sub

Private Sub GotFlag_Timer()
Dim speed As Integer
'If Level.Caption = "1" Then
'speed = 200
'End If
If Image1.Top = "16" And Image1.Left = "368" Then

MsgBox "Level " & Level.Caption & " Completed!", vbExclamation
Level.Caption = Level.Caption + 1
If Label5.Caption > 0 Then
Label5.Caption = Label5.Caption - 10
End If
'speed = speed - 10
Traffic.Interval = Label5.Caption
Lives.Caption = 3
Image1.Top = 464
Image1.Left = 368
End If
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Image3_Click()

End Sub

Private Sub mnuabout_Click()
Form2.Show

End Sub

Private Sub mnueasy_Click()
mnumedium.Checked = False
mnueasy.Checked = True
mnuhard.Checked = False
End Sub

Private Sub mnuexit_Click()
Response = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit?")
If Response = vbYes Then
End

End If


End Sub

Private Sub mnuhard_Click()
mnumedium.Checked = False
mnueasy.Checked = False
mnuhard.Checked = True
End Sub

Private Sub mnukeys_Click()
MsgBox "Up, Down, Left, Right arrow keys control the frogs direction!", vbInformation, "Keys"

End Sub

Private Sub mnumedium_Click()
mnumedium.Checked = True
mnueasy.Checked = False
mnuhard.Checked = False

End Sub

Private Sub mnunew_Click()
E1.Visible = False
E2.Visible = False
E3.Visible = False
E4.Visible = False

Lives.Caption = "3"
Image1.Top = 464
Image1.Left = 368
Flag.Left = 368
Flag.Top = 16
Level.Caption = "1"
Label5.Caption = "200"
Traffic.Interval = 200
Timer2.Enabled = True
L1.Picture = Car2.Picture
L2.Picture = Car3.Picture
L3.Picture = Car3.Picture
L4.Picture = Car2.Picture
L5.Picture = Car2.Picture

c1.Picture = Car1.Picture
c2.Picture = Car4.Picture
c3.Picture = Car5.Picture
c4.Picture = Car6.Picture
c5.Picture = Car4.Picture

s1.Picture = Car1.Picture
s2.Picture = Car6.Picture
s3.Picture = Car6.Picture
s4.Picture = Car1.Picture
s5.Picture = Car5.Picture


f1.Picture = Car3.Picture
f2.Picture = Car2.Picture
f3.Picture = Car2.Picture
f4.Picture = Car3.Picture
f5.Picture = Car3.Picture
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Image1.Left
Label3.Caption = Image1.Top

End Sub

Private Sub Timer2_Timer()
If Level.Caption = "20" Then
Traffic.Interval = "120"
Label5.Caption = "120"
E1.Visible = True
E2.Visible = True
E3.Visible = True
E4.Visible = True
L1.Picture = B2.Picture
L2.Picture = B1.Picture
L3.Picture = B1.Picture
L4.Picture = B2.Picture
L5.Picture = B2.Picture

c1.Picture = B3.Picture
c2.Picture = B4.Picture
c3.Picture = B5.Picture
c4.Picture = B6.Picture
c5.Picture = B6.Picture

s1.Picture = B3.Picture
s2.Picture = B6.Picture
s3.Picture = B6.Picture
s4.Picture = B3.Picture
s5.Picture = B5.Picture


f1.Picture = B1.Picture
f2.Picture = B2.Picture
f3.Picture = B2.Picture
f4.Picture = B1.Picture
f5.Picture = B1.Picture
Timer2.Enabled = False
End If

End Sub

Private Sub Traffic_Timer()
'If Level.Caption = "1" Then
If mnumedium.Checked = True Then
L4.Left = L4.Left - 16
L3.Left = L3.Left - 16
L2.Left = L2.Left - 16
L1.Left = L1.Left - 16
L5.Left = L5.Left - 16
'L2.Left = L1.Left + 128
c4.Left = c4.Left + 16
c3.Left = c3.Left + 16
c2.Left = c2.Left + 16
c1.Left = c1.Left + 16
c5.Left = c5.Left + 16
s4.Left = s4.Left + 16
s3.Left = s3.Left + 16
s2.Left = s2.Left + 16
s1.Left = s1.Left + 16
s5.Left = s5.Left + 16
f4.Left = f4.Left - 16
f3.Left = f3.Left - 16
f2.Left = f2.Left - 16
f1.Left = f1.Left - 16
f5.Left = f5.Left - 16
End If
If mnueasy.Checked = True Then
L4.Left = L4.Left - 8
L3.Left = L3.Left - 8
L2.Left = L2.Left - 8
L1.Left = L1.Left - 8
L5.Left = L5.Left - 8
'L2.Left = L1.Left + 128
c4.Left = c4.Left + 8
c3.Left = c3.Left + 8
c2.Left = c2.Left + 8
c1.Left = c1.Left + 8
c5.Left = c5.Left + 8
s4.Left = s4.Left + 8
s3.Left = s3.Left + 8
s2.Left = s2.Left + 8
s1.Left = s1.Left + 8
s5.Left = s5.Left + 8
f4.Left = f4.Left - 8
f3.Left = f3.Left - 8
f2.Left = f2.Left - 8
f1.Left = f1.Left - 8
f5.Left = f5.Left - 8
End If
If mnuhard.Checked = True Then
L4.Left = L4.Left - 32
L3.Left = L3.Left - 32
L2.Left = L2.Left - 32
L1.Left = L1.Left - 32
L5.Left = L5.Left - 32
'L2.Left = L1.Left + 128
c4.Left = c4.Left + 32
c3.Left = c3.Left + 32
c2.Left = c2.Left + 32
c1.Left = c1.Left + 32
c5.Left = c5.Left + 32
s4.Left = s4.Left + 32
s3.Left = s3.Left + 32
s2.Left = s2.Left + 32
s1.Left = s1.Left + 32
s5.Left = s5.Left + 32
f4.Left = f4.Left - 32
f3.Left = f3.Left - 32
f2.Left = f2.Left - 32
f1.Left = f1.Left - 32
f5.Left = f5.Left - 32
End If
If L1.Left < 16 Then
L1.Left = 616

End If
If L2.Left < 16 Then
L2.Left = 616

End If
If L3.Left < 16 Then
L3.Left = 616

End If
If L4.Left < 16 Then
L4.Left = 616

End If
If L5.Left < 16 Then
L5.Left = 616

End If
If c1.Left > 616 Then
c1.Left = 16

End If
If c2.Left > 616 Then
c2.Left = 16

End If
If c3.Left > 616 Then
c3.Left = 16

End If
If c4.Left > 616 Then
c4.Left = 16

End If
If c5.Left > 616 Then
c5.Left = 16

End If



If s1.Left > 616 Then
s1.Left = 16

End If
If s2.Left > 616 Then
s2.Left = 16

End If
If s3.Left > 616 Then
s3.Left = 16

End If
If s4.Left > 616 Then
s4.Left = 16

End If
If s5.Left > 616 Then
s5.Left = 16

End If
If f1.Left < 16 Then
f1.Left = 616

End If
If f2.Left < 16 Then
f2.Left = 616

End If
If f3.Left < 16 Then
f3.Left = 616

End If
If f4.Left < 16 Then
f4.Left = 616

End If
If f5.Left < 16 Then
f5.Left = 616

End If


'End If

End Sub

Private Sub Wait_Timer()

Image1.Picture = Frog.Picture
Image1.Top = 464
Image1.Left = 368
Colllide.Enabled = True
Water.Enabled = True
Wait.Enabled = False

End Sub

Private Sub Water_Timer()
If Image1.Top <= Picture2.Top + Picture2.Height And Image1.Left >= Picture2.Left Then
Lives.Caption = Lives.Caption - 1
Wait.Enabled = True
Water.Enabled = False
End If

If Image1.Top <= Picture1.Top + Picture1.Height And Image1.Left <= Picture1.Left + Picture1.Width Then
Lives.Caption = Lives.Caption - 1
Wait.Enabled = True
Water.Enabled = False
End If
Call GameOver

End Sub
Public Sub GameOver()

If Lives.Caption <= "0" Then
Response = MsgBox("Game Over! You made it to level " & Level.Caption & " Would you like to play again?", vbYesNo, "Play again?")
If Response = vbYes Then
Lives.Caption = "3"
Image1.Top = 464
Image1.Left = 368
Flag.Left = 368
Flag.Top = 16
Level.Caption = "1"
Label5.Caption = "200"
Traffic.Interval = 200
Timer2.Enabled = True
E1.Visible = False
E2.Visible = False
E3.Visible = False
E4.Visible = False
L1.Picture = Car2.Picture
L2.Picture = Car3.Picture
L3.Picture = Car3.Picture
L4.Picture = Car2.Picture
L5.Picture = Car2.Picture

c1.Picture = Car1.Picture
c2.Picture = Car4.Picture
c3.Picture = Car5.Picture
c4.Picture = Car6.Picture
c5.Picture = Car4.Picture

s1.Picture = Car1.Picture
s2.Picture = Car6.Picture
s3.Picture = Car6.Picture
s4.Picture = Car1.Picture
s5.Picture = Car5.Picture


f1.Picture = Car3.Picture
f2.Picture = Car2.Picture
f3.Picture = Car2.Picture
f4.Picture = Car3.Picture
f5.Picture = Car3.Picture
Else
End

End If

End If
End Sub

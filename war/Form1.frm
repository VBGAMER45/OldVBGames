VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Phone\Cd player"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2400
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1320
      Top             =   2520
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Skip"
      Height          =   255
      Left            =   840
      TabIndex        =   37
      Top             =   3360
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Un mute"
      Height          =   375
      Left            =   2640
      TabIndex        =   28
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mute Game music"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox timeWindow 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   735
      Left            =   2280
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton play 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton pause 
      Caption         =   "||"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton stopbtn 
      Caption         =   "x"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btrack 
      Caption         =   "|<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton ftrack 
      Caption         =   ">>|"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton rew 
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton ff 
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton eject 
      Caption         =   "ejc"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "*"
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
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      Caption         =   "#"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Reset"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   36
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   35
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   34
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   33
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   255
      Left            =   4680
      TabIndex        =   29
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Cd player"
      Height          =   495
      Left            =   2400
      TabIndex        =   26
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label totalplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   25
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label tracktime 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Option Explicit

Dim fastForwardSpeed As Long    ' seconds to seek for ff/rew
Dim fPlaying As Boolean         ' true if CD is currently playing
Dim fCDLoaded As Boolean        ' true if CD is the the player
Dim numTracks As Integer        ' number of tracks on audio CD
Dim trackLength() As String     ' array containing length of each track
Dim track As Integer            ' current track
Dim min As Integer              ' current minute on track
Dim sec As Integer              ' current second on track
Dim cmd As String               ' string to hold mci command strings

' Send a MCI command string
' If fShowError is true, display a message box on error
Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hwnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function

Private Sub Check1_Click()
form2.MediaPlayer1.Mute = True
End Sub

Private Sub Check2_Click()
form2.MediaPlayer1.Mute = False
End Sub

Private Sub Command13_Click()
If Label2.Caption = 1 Then
Label3.Caption = "#"
End If
If Label2.Caption = 2 Then
Label4.Caption = "#"
End If
If Label2.Caption = 3 Then
Label5.Caption = "#"
End If
If Label2.Caption = 4 Then
Label6.Caption = "#"
End If
If Label2.Caption = 5 Then
Label7.Caption = "#"
End If
If Label2.Caption = 6 Then
Label8.Caption = "#"
End If
If Label2.Caption = 7 Then
Label9.Caption = "#"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command14_Click()
If Label2.Caption = 1 Then
Label3.Caption = "*"
End If
If Label2.Caption = 2 Then
Label4.Caption = "*"
End If
If Label2.Caption = 3 Then
Label5.Caption = "*"
End If
If Label2.Caption = 4 Then
Label6.Caption = "*"
End If
If Label2.Caption = 5 Then
Label7.Caption = "*"
End If
If Label2.Caption = 6 Then
Label8.Caption = "*"
End If
If Label2.Caption = 7 Then
Label9.Caption = "*"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command3_Click()
If Label2.Caption = 1 Then
Label3.Caption = "3"
End If
If Label2.Caption = 2 Then
Label4.Caption = "3"
End If
If Label2.Caption = 3 Then
Label5.Caption = "3"
End If
If Label2.Caption = 4 Then
Label6.Caption = "3"
End If
If Label2.Caption = 5 Then
Label7.Caption = "3"
End If
If Label2.Caption = 6 Then
Label8.Caption = "3"
End If
If Label2.Caption = 7 Then
Label9.Caption = "3"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command4_Click()
If Label2.Caption = 1 Then
Label3.Caption = "4"
End If
If Label2.Caption = 2 Then
Label4.Caption = "4"
End If
If Label2.Caption = 3 Then
Label5.Caption = "4"
End If
If Label2.Caption = 4 Then
Label6.Caption = "4"
End If
If Label2.Caption = 5 Then
Label7.Caption = "4"
End If
If Label2.Caption = 6 Then
Label8.Caption = "4"
End If
If Label2.Caption = 7 Then
Label9.Caption = "4"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command5_Click()
If Label2.Caption = 1 Then
Label3.Caption = "5"
End If
If Label2.Caption = 2 Then
Label4.Caption = "5"
End If
If Label2.Caption = 3 Then
Label5.Caption = "5"
End If
If Label2.Caption = 4 Then
Label6.Caption = "5"
End If
If Label2.Caption = 5 Then
Label7.Caption = "5"
End If
If Label2.Caption = 6 Then
Label8.Caption = "5"
End If
If Label2.Caption = 7 Then
Label9.Caption = "5"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command6_Click()
If Label2.Caption = 1 Then
Label3.Caption = "6"
End If
If Label2.Caption = 2 Then
Label4.Caption = "6"
End If
If Label2.Caption = 3 Then
Label5.Caption = "6"
End If
If Label2.Caption = 4 Then
Label6.Caption = "6"
End If
If Label2.Caption = 5 Then
Label7.Caption = "6"
End If
If Label2.Caption = 6 Then
Label8.Caption = "6"
End If
If Label2.Caption = 7 Then
Label9.Caption = "6"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command7_Click()
If Label2.Caption = 1 Then
Label3.Caption = "7"
End If
If Label2.Caption = 2 Then
Label4.Caption = "7"
End If
If Label2.Caption = 3 Then
Label5.Caption = "7"
End If
If Label2.Caption = 4 Then
Label6.Caption = "7"
End If
If Label2.Caption = 5 Then
Label7.Caption = "7"
End If
If Label2.Caption = 6 Then
Label8.Caption = "7"
End If
If Label2.Caption = 7 Then
Label9.Caption = "7"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command8_Click()
If Label2.Caption = 1 Then
Label3.Caption = "8"
End If
If Label2.Caption = 2 Then
Label4.Caption = "8"
End If
If Label2.Caption = 3 Then
Label5.Caption = "8"
End If
If Label2.Caption = 4 Then
Label6.Caption = "8"
End If
If Label2.Caption = 5 Then
Label7.Caption = "8"
End If
If Label2.Caption = 6 Then
Label8.Caption = "8"
End If
If Label2.Caption = 7 Then
Label9.Caption = "8"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command9_Click()
If Label2.Caption = 1 Then
Label3.Caption = "9"
End If
If Label2.Caption = 2 Then
Label4.Caption = "9"
End If
If Label2.Caption = 3 Then
Label5.Caption = "9"
End If
If Label2.Caption = 4 Then
Label6.Caption = "9"
End If
If Label2.Caption = 5 Then
Label7.Caption = "9"
End If
If Label2.Caption = 6 Then
Label8.Caption = "9"
End If
If Label2.Caption = 7 Then
Label9.Caption = "9"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Form_Load()

' If we're already running, then quit
If (App.PrevInstance = True) Then
    End
End If

' Initialize variables
Timer1.Enabled = False
fastForwardSpeed = 5
fCDLoaded = False

' If the cd is being used, then quit
If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    End
End If

SendMCIString "set cd time format tmsf wait", True
Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close all MCI devices opened by this program
SendMCIString "close all", False
End Sub

' Play the CD
Private Sub play_Click()
SendMCIString "play cd", True
fPlaying = True
End Sub
' Stop the CD play
Private Sub stopbtn_Click()
SendMCIString "stop cd wait", True
cmd = "seek cd to " & track
SendMCIString cmd, True
fPlaying = False
Update
End Sub
' Pause the CD
Private Sub pause_Click()
SendMCIString "pause cd", True
fPlaying = False
Update
End Sub
' Eject the CD
Private Sub eject_Click()
SendMCIString "set cd door open", True
Update
End Sub
' Fast forward
Private Sub ff_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) + fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) + fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
' Rewind the CD
Private Sub rew_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) - fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) - fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
' Forward track
Private Sub ftrack_Click()
If (track < numTracks) Then
    If (fPlaying) Then
        cmd = "play cd from " & track + 1
        SendMCIString cmd, True
    Else
        cmd = "seek cd to " & track + 1
        SendMCIString cmd, True
    End If
Else
    SendMCIString "seek cd to 1", True
End If
Update
End Sub
' Go to previous track
Private Sub btrack_Click()
Dim from As String
If (min = 0 And sec = 0) Then
    If (track > 1) Then
        from = CStr(track - 1)
    Else
        from = CStr(numTracks)
    End If
Else
    from = CStr(track)
End If
If (fPlaying) Then
    cmd = "play cd from " & from
    SendMCIString cmd, True
Else
    cmd = "seek cd to " & from
    SendMCIString cmd, True
End If
Update
End Sub
' Update the display and state variables
Private Sub Update()
Static s As String * 30

' Check if CD is in the player
mciSendString "status cd media present", s, Len(s), 0
If (CBool(s)) Then
    ' Enable all the controls, get CD information
    If (fCDLoaded = False) Then
        mciSendString "status cd number of tracks wait", s, Len(s), 0
        numTracks = CInt(Mid$(s, 1, 2))
        eject.Enabled = True
        
        ' If CD only has 1 track, then it's probably a data CD
        If (numTracks = 1) Then
            Exit Sub
        End If
        
        mciSendString "status cd length wait", s, Len(s), 0
        totalplay.Caption = "Tracks: " & numTracks & "  Total time: " & s
        ReDim trackLength(1 To numTracks)
        Dim i As Integer
        For i = 1 To numTracks
            cmd = "status cd length track " & i
            mciSendString cmd, s, Len(s), 0
            trackLength(i) = s
        Next
        play.Enabled = True
        pause.Enabled = True
        ff.Enabled = True
        rew.Enabled = True
        ftrack.Enabled = True
        btrack.Enabled = True
        stopbtn.Enabled = True
        fCDLoaded = True
        SendMCIString "seek cd to 1", True
    End If

    ' Update the track time display
    mciSendString "status cd position", s, Len(s), 0
    track = CInt(Mid$(s, 1, 2))
    min = CInt(Mid$(s, 4, 2))
    sec = CInt(Mid$(s, 7, 2))
    timeWindow.Text = "[" & Format(track, "00") & "] " & Format(min, "00") _
            & ":" & Format(sec, "00")
    tracktime.Caption = "Track time: " & trackLength(track)
    
    ' Check if CD is playing
    mciSendString "status cd mode", s, Len(s), 0
    fPlaying = (Mid$(s, 1, 7) = "playing")
Else
    eject.Enabled = False
    ' Disable all the controls, clear the display
    If (fCDLoaded = True) Then
        play.Enabled = False
        pause.Enabled = False
        ff.Enabled = False
        rew.Enabled = False
        ftrack.Enabled = False
        btrack.Enabled = False
        stopbtn.Enabled = False
        fCDLoaded = False
        fPlaying = False
        totalplay.Caption = ""
        tracktime.Caption = ""
        timeWindow.Text = ""
    End If
End If
End Sub
' Set the fast-forward speed
Private Sub ffspeed_Click()
Dim s As String
s = InputBox("Enter the new speed in seconds", "Fast Forward Speed", CStr(fastForwardSpeed))
If IsNumeric(s) Then
    fastForwardSpeed = CLng(s)
End If
End Sub

Private Sub Timer1_Timer()
Update
If Text1.Text = "yeah" Then
form2.money.Caption = form2.money.Caption + 100
End If
If Text1.Text = "rocket" Then
form2.mnuspace.Enabled = True
End If
If Text1.Text = "cia" Then
Text2.Visible = True
Text2.Text = "           CIA           Help watch out form the underground they have sercert units . End Of transmission!       "
End If

End Sub
Private Sub Command1_Click()
If Label2.Caption = 1 Then
Label3.Caption = "1"
End If
If Label2.Caption = 2 Then
Label4.Caption = "1"
End If
If Label2.Caption = 3 Then
Label5.Caption = "1"
End If
If Label2.Caption = 4 Then
Label6.Caption = "1"
End If
If Label2.Caption = 5 Then
Label7.Caption = "1"
End If
If Label2.Caption = 6 Then
Label8.Caption = "1"
End If
If Label2.Caption = 7 Then
Label9.Caption = "1"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command10_Click()
'vbNormal.Enabled = False
If Label2.Caption = 1 Then
Label3.Caption = "0"
End If
If Label2.Caption = 2 Then
Label4.Caption = "0"
End If
If Label2.Caption = 3 Then
Label5.Caption = "0"
End If
If Label2.Caption = 4 Then
Label6.Caption = "0"
End If
If Label2.Caption = 5 Then
Label7.Caption = "0"
End If
If Label2.Caption = 6 Then
Label8.Caption = "0"
End If
If Label2.Caption = 7 Then
Label9.Caption = "0"
End If
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command11_Click()
'i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'If Text1.Text = 3070727 Then
'MsgBox "cheater"
'form2.money = money + 1000

If Text1.Text = "yeah" Then
form2.money.Caption = form2.money.Caption + 100
End If
If Text1.Text = "rocket" Then
form2.mnuspace.Enabled = True
End If
Unload Me

'
'Unload Me
'End If
End Sub

Private Sub Command12_Click()
'i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Text1.Text = 1
'Label1.Caption = 0
Label2.Caption = 1
Label3.Caption = 0
Label4.Caption = 0
Label5.Caption = 0
Label6.Caption = 0
Label7.Caption = 0
Label8.Caption = 0
Label9.Caption = 0

End Sub

Private Sub Command15_Click()
Label2.Caption = Label2.Caption + 1

End Sub

Private Sub Command16_Click()
If Label2.Caption <= 0 Then
MsgBox "Not a Sation"
End If
If Label2.Caption >= 0 Then
Label2.Caption = Label2.Caption - 1
End If



End Sub

Private Sub Command2_Click()
If Label2.Caption = 1 Then
Label3.Caption = "2"
End If
If Label2.Caption = 2 Then
Label4.Caption = "2"
End If
If Label2.Caption = 3 Then
Label5.Caption = "2"
End If
If Label2.Caption = 4 Then
Label6.Caption = "2"
End If
If Label2.Caption = 5 Then
Label7.Caption = "2"
End If
If Label2.Caption = 6 Then
Label8.Caption = "2"
End If
If Label2.Caption = 7 Then
Label9.Caption = "2"
End If
Label2.Caption = Label2.Caption + 1

End Sub

'jonathan valentin 1999


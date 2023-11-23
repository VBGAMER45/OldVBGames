VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{D194DF7A-B852-11D3-8FE2-0040D00AB433}#1.0#0"; "PROJECT1.OCX"
Begin VB.Form Form3 
   Caption         =   "Main Menu"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "Form3.frx":030A
   ScaleHeight     =   6525
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin Project1.TrayControl TrayControl1 
      Left            =   120
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "War"
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   30
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   30
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   30
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "c:\windows\desktop\vb projects\war\open.mid"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -440
      WindowlessVideo =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   720
      Top             =   0
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   7695
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6855
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   30
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "C:\windows\desktop\vb projects\war\cntdwn.avi"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   -1  'True
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   -1  'True
   End
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
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load Saved Game"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Scenario"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Multiplayer"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Campaign"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Unit Creater"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Universe Pack"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   7680
      Left            =   -2280
      Picture         =   "Form3.frx":15FC4E
      Top             =   0
      Width           =   11280
   End
   Begin VB.Label year 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -1080
      Picture         =   "Form3.frx":279C90
      Top             =   -120
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
'Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim mbCloseApp As Boolean 'Used to over-ride the 'minimize to tray' bahaviour
Dim mlWindowState As Long 'Remember the previous window state when form is minimized




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Only close the app if the user has selected exit. This provides
    'the 'minimize to tray' behaviour
    If Not mbCloseApp Then
        Cancel = True
        Me.Hide
        TrayControl1.Enabled = True 'minimize to tray
    End If

End Sub

Private Sub Form_Resize()

    'Hide this form if minimized. This provides
    'the 'minimize to tray' behaviour
    
    If Me.WindowState = vbMinimized Then
        Me.Hide
        TrayControl1.Enabled = True
    Else
        'remember the new window state
        mlWindowState = Me.WindowState
    End If
    
End Sub

'   ' mbCloseApp = True
   ' Unload Me
'End Sub

Private Sub TrayControl1_DblClick()
    
    TrayControl1.Enabled = False 'remove tray icon
    If Me.WindowState = vbMinimized Then
        Me.WindowState = mlWindowState
    End If
    Me.Show

End Sub

Private Sub TrayControl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu form2.mnufile
    End If
End Sub

'Private Sub formload()
'Private Sub Form_Load()
 '  avifile1 "C:\windows\desktop\vb projects\war\cntdwn.avi")
   
   
   
'End Sub
Private Sub avifile1_GotFocus()
'Private Sub cmdPlay_Click()
    'dlgOpen.Filter = "avi files (*.avi)|*.avi"
   ' dlgOpen.ShowOpen
  'dlgOpen.filename
    ' Play the file indefinitel
    'avifile1. (App.Path & "\cntdown.avi") '("C:\windows\desktop\vb projects\war\cntdwn.avi")
  ' , avifile1.Play 1, 0, 20
End Sub

'Private Sub cmdStop_Click()
   ' anmAVI.Stop
'End Sub

'The following code plays the file ten times, from the sixth to sixteenth frames (the first frame is frame 0):
'anmAVI.Play 10, 5, 15

'End Sub

Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
'race.Show vbModal, Me
Form6.Show
Unload Me

End Sub

Private Sub Command2_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
'form2.Command3.Visible = True


Form11.Label9.Caption = InputBox("Enter Your Name:", "Name")
Form61.Show vbModal, Me
End Sub

Private Sub Command3_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)

MsgBox "buy the add on pack"
End Sub

Private Sub Command4_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
CommonDialog1.ShowOpen
End Sub

Private Sub Command5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\opterm1.wav"), SND_ASYNC)
End
End Sub

Private Sub Command6_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
MsgBox "ORDER THE ADD ON PACK "
End Sub

Private Sub MediaPlayer1_stop() 'DVDNotify(ByVal EventCode As Long, ByVal EventParam1 As Long, ByVal EventParam2 As Long)
'Private Sub Form_Unload(Cancel As Integer)
    'Form1.MMControl1.Command = "Close"
'End Sub



MediaPlayer1.SendMouseClickEvents
'End If
MediaPlayer1.Visible = False
End Sub

Private Sub Command7_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
MsgBox "ORDER THE ADD ON PACK "
End Sub

Private Sub Command8_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\satlaser.wav"), SND_ASYNC)
MsgBox "Order the add on pack with 10 more solar systems"

End Sub

Private Sub Image2_Click()
Image2.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True

End Sub

Private Sub MediaPlayer1_Click(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
MediaPlayer1.Visible = False
End Sub

Private Sub Timer1_Timer()
year = year + 10
If year.Caption >= 310 Then
MediaPlayer1.Visible = False
year.Visible = False
End If
'If MediaPlayer1.Visible = False Then

End Sub
'jonathan valentin 1999
'App.Path & "\open.mid"
'c:\windows\desktop\vb projects\war\open.mid

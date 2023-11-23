VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "War is War"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":030A
   Picture         =   "Form2.frx":0BD4
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   953
      ButtonWidth     =   794
      ButtonHeight    =   847
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "a"
            Object.ToolTipText     =   "move out"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "b"
            Object.ToolTipText     =   "Medic"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "c"
            Object.ToolTipText     =   "don't shoot"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "d"
            Object.ToolTipText     =   "take cover"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "e"
            Object.ToolTipText     =   "take area"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "f"
            Object.ToolTipText     =   "Help"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "q1"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "q2"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "q3"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "q4"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   550
      Left            =   8400
      Top             =   2640
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   8760
      TabIndex        =   62
      Top             =   1440
      Width           =   135
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
      Filename        =   "C:\WINDOWS\Desktop\vb projects\war\jpi.mid"
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
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.CommandButton Command16 
      Height          =   255
      Left            =   8640
      TabIndex        =   60
      Top             =   3480
      Width           =   135
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6840
      Top             =   3960
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   9480
      TabIndex        =   58
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer18 
      Interval        =   400
      Left            =   6480
      Top             =   1440
   End
   Begin VB.Timer Timer17 
      Interval        =   50
      Left            =   7680
      Top             =   3960
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   2880
   End
   Begin VB.Timer Timer15 
      Interval        =   50
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer Timer14 
      Interval        =   300
      Left            =   6840
      Top             =   2520
   End
   Begin VB.Timer Timer13 
      Interval        =   75
      Left            =   6840
      Top             =   2520
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   6720
      Top             =   2760
   End
   Begin VB.Timer Timer11 
      Interval        =   20
      Left            =   7080
      Top             =   1800
   End
   Begin VB.Timer Timer10 
      Interval        =   30
      Left            =   3000
      Top             =   3480
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   2400
   End
   Begin VB.Timer Timer8 
      Interval        =   10000
      Left            =   2400
      Top             =   3120
   End
   Begin VB.HScrollBar hscmap 
      Height          =   255
      Left            =   0
      Max             =   6000
      TabIndex        =   54
      Top             =   7680
      Width           =   8895
   End
   Begin VB.VScrollBar vscmap 
      Height          =   7935
      Left            =   8880
      Max             =   6000
      TabIndex        =   53
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UnderGround"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   52
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   5040
      Top             =   0
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Help"
      Height          =   375
      Left            =   0
      TabIndex        =   51
      ToolTipText     =   "Help"
      Top             =   7920
      Width           =   735
   End
   Begin VB.Timer Timer6 
      Interval        =   3100
      Left            =   7560
      Top             =   2640
   End
   Begin VB.Timer Timer4 
      Interval        =   75
      Left            =   6240
      Top             =   720
   End
   Begin VB.Timer Timer3 
      Interval        =   300
      Left            =   8520
      Top             =   0
   End
   Begin VB.CommandButton Command15 
      Caption         =   "M"
      Height          =   375
      Left            =   11040
      TabIndex        =   43
      ToolTipText     =   "More"
      Top             =   7920
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      Caption         =   "M"
      Height          =   375
      Left            =   9720
      TabIndex        =   42
      ToolTipText     =   "More"
      Top             =   7920
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   8160
      Top             =   0
   End
   Begin VB.PictureBox Picture4 
      Height          =   1095
      Left            =   9360
      Picture         =   "Form2.frx":4B16
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture30 
         Height          =   1095
         Left            =   -120
         Picture         =   "Form2.frx":7650
         ScaleHeight     =   1035
         ScaleWidth      =   1755
         TabIndex        =   45
         ToolTipText     =   "CHAIN LINK"
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Phone\Cd"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   28
      ToolTipText     =   "Phone"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.PictureBox Picture29 
      Height          =   975
      Left            =   10680
      Picture         =   "Form2.frx":BECE
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   27
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture32 
         Height          =   975
         Left            =   0
         Picture         =   "Form2.frx":10DD4
         ScaleHeight     =   915
         ScaleWidth      =   1155
         TabIndex        =   47
         ToolTipText     =   "Moblie"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture8 
         Height          =   975
         Left            =   0
         Picture         =   "Form2.frx":15546
         ScaleHeight     =   915
         ScaleWidth      =   1155
         TabIndex        =   30
         ToolTipText     =   "MEDIUM TANK"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture28 
      Height          =   975
      Left            =   9360
      Picture         =   "Form2.frx":19ED8
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   26
      ToolTipText     =   "advanced power plant"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture36 
         Height          =   975
         Left            =   0
         Picture         =   "Form2.frx":1D87A
         ScaleHeight     =   915
         ScaleWidth      =   1155
         TabIndex        =   65
         ToolTipText     =   "Airport"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture17 
         Height          =   1575
         Left            =   0
         Picture         =   "Form2.frx":1FDE8
         ScaleHeight     =   1515
         ScaleWidth      =   1155
         TabIndex        =   33
         ToolTipText     =   "HELIPAD"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Sell"
      Height          =   375
      Left            =   8040
      MouseIcon       =   "Form2.frx":21BB2
      MousePointer    =   99  'Custom
      TabIndex        =   25
      ToolTipText     =   "Sell"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.PictureBox Picture26 
      Height          =   1095
      Left            =   9360
      Picture         =   "Form2.frx":21EBC
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   24
      ToolTipText     =   "Barrecks"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture35 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":26962
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   50
         ToolTipText     =   "Sam site"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture10 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":2737C
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   32
         ToolTipText     =   "radar"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Repair"
      Height          =   375
      Left            =   7080
      MouseIcon       =   "Form2.frx":284F6
      MousePointer    =   99  'Custom
      TabIndex        =   23
      ToolTipText     =   "Repair"
      Top             =   7920
      Width           =   975
   End
   Begin VB.PictureBox Picture7 
      Height          =   1215
      Left            =   9360
      Picture         =   "Form2.frx":28800
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   22
      ToolTipText     =   "Refinary"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture31 
         Height          =   1215
         Left            =   0
         Picture         =   "Form2.frx":2D9A2
         ScaleHeight     =   1155
         ScaleWidth      =   1515
         TabIndex        =   46
         ToolTipText     =   "CONCRETE WALL"
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture9 
         Height          =   1095
         Left            =   -240
         Picture         =   "Form2.frx":32004
         ScaleHeight     =   1035
         ScaleWidth      =   1395
         TabIndex        =   31
         ToolTipText     =   "Guard Tower"
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   9360
      Picture         =   "Form2.frx":36876
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   21
      ToolTipText     =   "Construction yard"
      Top             =   960
      Width           =   1215
      Begin VB.PictureBox Picture19 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":394B8
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   35
         ToolTipText     =   "REPAIR CENTER"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Down"
      Height          =   375
      Left            =   11280
      TabIndex        =   20
      ToolTipText     =   "Down"
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Up"
      Height          =   375
      Left            =   10560
      TabIndex        =   19
      ToolTipText     =   "Up"
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Down"
      Height          =   375
      Left            =   9960
      TabIndex        =   18
      ToolTipText     =   "Down"
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Up"
      Height          =   375
      Left            =   9120
      TabIndex        =   17
      ToolTipText     =   "Up"
      Top             =   7920
      Width           =   615
   End
   Begin VB.PictureBox Picture16 
      Height          =   975
      Left            =   10680
      Picture         =   "Form2.frx":3D906
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture25 
         Height          =   975
         Left            =   0
         Picture         =   "Form2.frx":426E8
         ScaleHeight     =   915
         ScaleWidth      =   1155
         TabIndex        =   41
         ToolTipText     =   "Mcv"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture15 
      Height          =   1095
      Left            =   10680
      Picture         =   "Form2.frx":472A2
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture24 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":4BC2C
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   40
         ToolTipText     =   "SUPPORT AIRCRAFT"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture14 
      Height          =   1095
      Left            =   10680
      Picture         =   "Form2.frx":507DE
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture23 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":55390
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   39
         ToolTipText     =   "TRANSPORT"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture13 
      Height          =   1095
      Left            =   10680
      Picture         =   "Form2.frx":5A4AE
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture22 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":5F17C
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   38
         ToolTipText     =   "OCRA"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture12 
      Height          =   1095
      Left            =   10680
      Picture         =   "Form2.frx":6407E
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture34 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":68C30
         ScaleHeight     =   1035
         ScaleWidth      =   1155
         TabIndex        =   49
         ToolTipText     =   "Gun Boat"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture21 
         Height          =   1095
         Left            =   -120
         Picture         =   "Form2.frx":6D7D2
         ScaleHeight     =   1035
         ScaleWidth      =   1275
         TabIndex        =   37
         ToolTipText     =   "MAMMOTH TANK"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture11 
      Height          =   975
      Left            =   10680
      Picture         =   "Form2.frx":72494
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture33 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":7739A
         ScaleHeight     =   1035
         ScaleWidth      =   1515
         TabIndex        =   48
         ToolTipText     =   "Hover"
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture20 
         Height          =   975
         Left            =   0
         Picture         =   "Form2.frx":7C3BC
         ScaleHeight     =   915
         ScaleWidth      =   1155
         TabIndex        =   36
         ToolTipText     =   "APC"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   855
      Left            =   9360
      Picture         =   "Form2.frx":811A6
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   8
      ToolTipText     =   "Tech Center $1000"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture27 
         Height          =   855
         Left            =   -240
         Picture         =   "Form2.frx":83E28
         ScaleHeight     =   795
         ScaleWidth      =   1515
         TabIndex        =   44
         ToolTipText     =   "SAND BAGS"
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   1095
      Left            =   9360
      Picture         =   "Form2.frx":8859A
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   7
      ToolTipText     =   "silo"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   9360
      Picture         =   "Form2.frx":88F7C
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   6
      ToolTipText     =   "Communications"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   9360
      Picture         =   "Form2.frx":8B10E
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   5
      ToolTipText     =   "Power Plant"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox Picture18 
         Height          =   1095
         Left            =   0
         Picture         =   "Form2.frx":8D244
         ScaleHeight     =   1035
         ScaleWidth      =   1755
         TabIndex        =   34
         ToolTipText     =   "ADVANCED COMMUNICATIONS"
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   8880
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dipomacly"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Communicate"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      ToolTipText     =   "Communicate"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Wheather\Map"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      ToolTipText     =   "News"
      Top             =   7920
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "war"
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   9840
      TabIndex        =   64
      ToolTipText     =   "research"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image39 
      Height          =   480
      Left            =   8400
      Picture         =   "Form2.frx":918A6
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   8640
      TabIndex        =   63
      Top             =   4320
      Width           =   255
   End
   Begin VB.Image Image40 
      Height          =   480
      Left            =   2280
      Picture         =   "Form2.frx":91BB0
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image80 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   10080
      TabIndex        =   59
      Top             =   840
      Width           =   255
   End
   Begin VB.Image Image63 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image62 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image61 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image60 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   960
      TabIndex        =   56
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image53 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image52 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image50 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image49 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image48 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image47 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image46 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image44 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "You have been changed to enemy."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   55
      Top             =   840
      Width           =   1695
   End
   Begin VB.Image Image43 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image42 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image32 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image29 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image35 
      Height          =   480
      Left            =   7680
      Picture         =   "Form2.frx":91EBA
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image28 
      Height          =   495
      Left            =   6720
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image26 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image24 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image23 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image22 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image21 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image20 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image19 
      Height          =   495
      Left            =   7680
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image18 
      Height          =   495
      Left            =   7680
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image16 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image15 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image14 
      Height          =   495
      Left            =   6720
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image13 
      Height          =   495
      Left            =   6720
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image12 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   7680
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image36 
      Height          =   480
      Left            =   8160
      Picture         =   "Form2.frx":921C4
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   7680
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6720
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image38 
      Height          =   480
      Left            =   1560
      Picture         =   "Form2.frx":924CE
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image37 
      Height          =   480
      Left            =   2640
      Picture         =   "Form2.frx":927D8
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Image27 
      Height          =   135
      Left            =   9480
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image155 
      Height          =   135
      Left            =   9600
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image154 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image153 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image152 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image151 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image150 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image148 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image147 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image146 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image145 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image144 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image143 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image142 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image141 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image140 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image139 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image136 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image135 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image134 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image133 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image88 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image132 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image131 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image130 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image129 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image128 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image127 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image126 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image125 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image124 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image123 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image121 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image120 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image119 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image118 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image117 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image116 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image115 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image114 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image113 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image112 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image111 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image110 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image109 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image108 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image107 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image106 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image105 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image104 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image103 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image102 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image101 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image100 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image99 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image97 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image96 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image95 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image94 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image93 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image92 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image91 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image89 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image86 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image85 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image84 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image83 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image79 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image78 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image77 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image76 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image75 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image74 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image73 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image72 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image71 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image70 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image69 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image68 
      Height          =   495
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image67 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   3840
      Width           =   975
   End
   Begin VB.Image Image66 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image65 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image64 
      Height          =   495
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image59 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image Image58 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   5760
      Width           =   975
   End
   Begin VB.Image Image57 
      Height          =   495
      Left            =   5760
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image56 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image55 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image Image54 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image51 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image45 
      Height          =   495
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   6720
      Width           =   975
   End
   Begin VB.Image Image41 
      Height          =   495
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image34 
      Height          =   585
      Left            =   7680
      Picture         =   "Form2.frx":92AE2
      ToolTipText     =   "German Figther"
      Top             =   600
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "You are an ally"
      Height          =   255
      Left            =   7560
      TabIndex        =   57
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image31 
      Height          =   495
      Left            =   3840
      OLEDropMode     =   2  'Automatic
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Image30 
      Height          =   495
      Left            =   4800
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   26
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":94A38
            Key             =   "a"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":950FA
            Key             =   "b"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":958CC
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":9615E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":969D0
            Key             =   "e"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":9727E
            Key             =   "f"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":97598
            Key             =   "for"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":980A2
            Key             =   "for2"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":98B0C
            Key             =   "for4"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":99492
            Key             =   "for5"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":99F24
            Key             =   "q"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":9A862
            Key             =   "q2"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":9B20C
            Key             =   "q3"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":9BB92
            Key             =   "q4"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image25 
      Height          =   135
      Left            =   9720
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Power 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   9600
      TabIndex        =   16
      ToolTipText     =   "Power"
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Year1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9360
      TabIndex        =   15
      Top             =   600
      Width           =   735
   End
   Begin VB.Label year 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10080
      TabIndex        =   4
      ToolTipText     =   "Year"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label money 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5000"
      Height          =   255
      Left            =   10680
      TabIndex        =   1
      ToolTipText     =   "Money"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image33 
      Height          =   8340
      Left            =   0
      Picture         =   "Form2.frx":9C518
      Top             =   0
      Width           =   12525
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuload 
         Caption         =   "Load"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnmain 
         Caption         =   "Main Menu"
      End
      Begin VB.Menu Mnurestart 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Options"
      Begin VB.Menu mnuset 
         Caption         =   "Setting"
      End
   End
   Begin VB.Menu mnuspace 
      Caption         =   "Space"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnumusic 
      Caption         =   "Music"
      Begin VB.Menu mnut1 
         Caption         =   "Track1"
      End
      Begin VB.Menu mnut2 
         Caption         =   "Track2"
      End
      Begin VB.Menu mnutrack3 
         Caption         =   "Track3"
      End
      Begin VB.Menu mnutrack4 
         Caption         =   "Track4"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnutip 
         Caption         =   "Tip"
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
'jonathan valentin 1999

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim grid As String
Dim red As String
Dim done As Boolean
Dim response As String
'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


    Public Sub DIE()
    Call Name
    
    End Sub


'Private Sub Command1_Click()
'Dim Form9 As New Form9
'Form9.Label2.Visible = False
'Form9.Show vbModal, Me

Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Form12.Show

End Sub

'End Sub
Private Sub Command10_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'MouseIcon =
'form2.MouseIcon = form2.
End Sub
Private Sub Command11_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Private Sub Form_Click()
   ' Dim I   ' Declare variable.
    ' Change mouse pointer to hourglass.
  'Screen.MousePointer = vbCustom
   ' Set random color and draw circles on form.
  '  For I = 0 To ScaleWidth Step 50
        'ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        'Circle (I, ScaleHeight * Rnd), 400
  'Screen.MousePointer = vbCustom
  ' Next Screen
    ' Return mouse pointer to normal.
   ' Screen.MousePointer = vbDefault
End Sub

'Private Sub Command11_KeyPress(KeyAscii As Integer)

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'f vbKeyS Then

'End If
'If vbKeyA Then
'Screen.MousePointer = vbDefault
'End If
'E'nd Sub

'Call sell
Private Sub Command12_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Dim Form1 As New Form1
Form1.Show

End Sub

Private Sub Command12_KeyPress(KeyAscii As Integer)
If KeyCode = vbKeyP Then
mnupassword.Enabled = True

End If
If KeyCode = vbKeyMenu Then
Unload Me
Dim Form3 As New Form3
Form3.Show vbModal, Me
End If
If KeyCode = vbKeyEscape Then
End
End If
If KeyCode = vbKeyPause Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Timer5.Interval = 0
Timer6.Interval = 0
Timer7.Interval = 0
Timer8.Interval = 0
Timer9.Interval = 0
Timer10.Interval = 0
End If
If KeyCode = vbKeyP Then
Timer1.Interval = 300
Timer2.Interval = 300
Timer3.Interval = 300
Timer4.Interval = 300
Timer5.Interval = 300
Timer6.Interval = 300
Timer7.Interval = 300
Timer8.Interval = 300
Timer9.Interval = 300
Timer10.Interval = 300
End If

End Sub

Private Sub Command13_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
frmAbout.Show

End Sub

Private Sub Command13_KeyPress(KeyAscii As Integer)
If KeyCode = vbKeyP Then
mnupassword.Enabled = True

End If
If KeyCode = vbKeyMenu Then
Dim Form3 As New Form3
Form3.Show vbModal, Me
End If
If KeyCode = vbKeyEscape Then
End
End If
If KeyCode = vbKeyPause Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Timer5.Interval = 0
Timer6.Interval = 0
Timer7.Interval = 0
Timer8.Interval = 0
Timer9.Interval = 0
Timer10.Interval = 0
End If
If KeyCode = vbKeyP Then
Timer1.Interval = 300
Timer2.Interval = 300
Timer3.Interval = 300
Timer4.Interval = 300
Timer5.Interval = 300
Timer6.Interval = 300
Timer7.Interval = 300
Timer8.Interval = 300
Timer9.Interval = 300
Timer10.Interval = 300
End If
End Sub


Private Sub Command14_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Picture27.Visible = True
Picture30.Visible = True
Picture31.Visible = True
Picture35.Visible = True
Picture36.Visible = True
End Sub

Private Sub Command18_Click()
'Private Sub Command5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Form6.Show vbModal, Me
If Label5.Caption = "1a" Then
End If

If Label5.Caption = "1b" Then
End If


If Label5.Caption = "1c" Then
End If

If Label5.Caption = "1d" Then

End If
If Label5.Caption = "1e" Then
End If
End Sub
Private Sub Command5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Form6.Show vbModal, Me
If Label5.Caption = "1a" Then
Image136.Picture = Image44.Picture
Image136.OLEDropMode = 1
Image44.OLEDropMode = 0
'Image136.Picture = (none)
If Image44.OLEDropMode = 0 Then

Image44.Picture = Image119.Picture
'Image44.Picture = (none)
Image44.OLEDropMode = 1
Image119.OLEDropMode = 0
End If


End If

If Label5.Caption = "1b" Then
End If


If Label5.Caption = "1c" Then
End If

If Label5.Caption = "1d" Then

End If
If Label5.Caption = "1e" Then
End If
End Sub

Private Sub Command16_Click()
'Private Sub Command5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Form6.Show vbModal, Me
If Label5.Caption = "1a" Then
End If

If Label5.Caption = "1b" Then
End If


If Label5.Caption = "1c" Then
End If

If Label5.Caption = "1d" Then

End If
If Label5.Caption = "1e" Then
End If
'Private Sub Form_Click()
   ' Dim I   ' Declare variable.
    ' Change mouse pointer to hourglass.
   ' Screen.MousePointer = vbHourglass
   ' Set random color and draw circles on form.
   ' For I = 0 To ScaleWidth Step 50
       ' ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
       ' Circle (I, ScaleHeight * Rnd), 400
   ' Next
     Screen.MousePointer = vbDefault
'End Sub
End Sub ' Return mouse pointer to normal.

Private Sub Command15_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Picture32.Visible = True
Picture33.Visible = True
Picture34.Visible = True
End Sub

Private Sub Command17_Click()
'If Label5.Caption = "1a" Then




'End If
'Private Sub Command5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Form6.Show vbModal, Me
If Label5.Caption = "1a" Then
End If

If Label5.Caption = "1b" Then
End If


If Label5.Caption = "1c" Then
End If

If Label5.Caption = "1d" Then

End If
If Label5.Caption = "1e" Then
End If
End Sub

Private Sub Command2_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Form5.Show



End Sub

Private Sub Command3_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Form7.Show vbModal, Me

Label3.Caption = InputBox("Talk", "enter stuff")
'If Label3.Caption = "jv rules" Then
'End If
Timer9.Enabled = True
If Label3.Caption = "1" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt008.wav"), SND_ASYNC)
End If
If Label3.Caption = "2" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt009.wav"), SND_ASYNC)
End If
If Label3.Caption = "3" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt013.wav"), SND_ASYNC)
End If
If Label3.Caption = "4" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt015.wav"), SND_ASYNC)
End If
If Label3.Caption = "5" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt017.wav"), SND_ASYNC)
End If
If Label3.Caption = "6" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt018.wav"), SND_ASYNC)
End If
If Label3.Caption = "7" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt019.wav"), SND_ASYNC)
End If
If Label3.Caption = "8" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt020.wav"), SND_ASYNC)
End If
If Label3.Caption = "9" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt022.wav"), SND_ASYNC)
End If
If Label3.Caption = "10" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt025.wav"), SND_ASYNC)
End If
If Label3.Caption = "11" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt001.wav"), SND_ASYNC)
End If
If Label3.Caption = "12" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt002.wav"), SND_ASYNC)
End If
If Label3.Caption = "13" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt007.wav"), SND_ASYNC)
End If
If Label3.Caption = "14" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt011.wav"), SND_ASYNC)
End If
If Label3.Caption = "15" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt012.wav"), SND_ASYNC)
End If
If Label3.Caption = "16" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt014.wav"), SND_ASYNC)
End If
If Label3.Caption = "17" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt016.wav"), SND_ASYNC)
End If
If Label3.Caption = "18" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt021.wav"), SND_ASYNC)
End If
If Label3.Caption = "19" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt023.wav"), SND_ASYNC)
End If
If Label3.Caption = "20" Then
i = sndPlaySound(ByVal CStr(App.Path & "\taunt024.wav"), SND_ASYNC)
End If
Label3.Visible = True
End Sub

Private Sub Command4_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
If Label6.Caption = "1" Then
Form4.Label11.Caption = "1"

End If
If Label6.Caption = "2" Then
Form4.Label11.Caption = "2"

End If
If Label6.Caption = "3" Then
Form4.Label11.Caption = "3"

End If
If Label6.Caption = "4" Then
Form4.Label11.Caption = "4"

End If
If Label6.Caption = "5" Then
Form4.Label11.Caption = "5"

End If
If Label6.Caption = "6" Then
Form4.Label11.Caption = "6"

End If
If Label6.Caption = "7" Then
Form4.Label11.Caption = "7"

End If
If Label6.Caption = "8" Then
Form4.Label11.Caption = "8"

End If
If Label6.Caption = "9" Then
Form4.Label11.Caption = "9"

End If
Form4.Show

End Sub

'Private Sub Command5_Click()
'i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'Form6.Show vbModal, Me
'If Label5.Caption = "1a" Then
'End If

'If Label5.Caption = "1b" Then
'End If


'If Label5.Caption = "1c" Then
'End If

'If Label5.Caption = "1d" Then

'End If
'If Label5.Caption = "1e" Then
'End If

'End Sub

Private Sub Command6_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Picture4.Visible = False
Picture6.Visible = False
Picture9.Visible = False
Picture10.Visible = False
Picture17.Visible = False
Picture18.Visible = False
Picture19.Visible = False
Picture27.Visible = False
Picture30.Visible = False
Picture31.Visible = False
Picture35.Visible = False
Picture2.Visible = True
Picture36.Visible = False

End Sub

Private Sub Command7_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
'If Picture3.Visible = True Then
'Picture5.Visible = False
Picture6.Visible = True
Picture4.Visible = True
Picture9.Visible = True
Picture10.Visible = True
Picture17.Visible = True
Picture18.Visible = True
Picture19.Visible = True
Picture27.Visible = False
Picture30.Visible = False
Picture31.Visible = False
Picture35.Visible = False
Picture2.Visible = False
Picture36.Visible = False


'End If

End Sub

Private Sub Command8_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Picture8.Visible = False
Picture20.Visible = False
Picture21.Visible = False
Picture22.Visible = False
Picture23.Visible = False
Picture24.Visible = False
Picture25.Visible = False
Picture32.Visible = False
Picture33.Visible = False
Picture34.Visible = False
End Sub

Private Sub Command9_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Picture8.Visible = True
Picture20.Visible = True
Picture21.Visible = True
Picture22.Visible = True
Picture23.Visible = True
Picture24.Visible = True
Picture25.Visible = True
Picture32.Visible = False
Picture33.Visible = False
Picture34.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyMenu Then
'Dim Form3 As New Form3
Form3.Show
End If
If KeyCode = vbKeyEscape Then
Call pgain

'End
End If
If KeyCode = vbKeyPause Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Timer5.Interval = 0
Timer6.Interval = 0
Timer7.Interval = 0
Timer8.Interval = 0
Timer9.Interval = 0
Timer10.Interval = 0
Timer11.Interval = 0
Timer12.Interval = 0
Timer13.Interval = 0
Timer14.Interval = 0
Timer15.Interval = 0
Timer16.Interval = 0
Timer17.Interval = 0
Timer18.Interval = 0
Timer19.Interval = 0
End If
If KeyCode = vbKeyP Then
Timer1.Interval = 300
Timer2.Interval = 10000
Timer3.Interval = 300
Timer4.Interval = 75
Timer5.Interval = 550
Timer6.Interval = 3100
Timer7.Interval = 1000
Timer8.Interval = 10000
Timer9.Interval = 1000
Timer10.Interval = 30
Timer11.Interval = 20
Timer12.Interval = 70
Timer13.Interval = 75
Timer14.Interval = 300
Timer15.Interval = 50
Timer16.Interval = 40
Timer17.Interval = 50
Timer18.Interval = 400
Timer19.Interval = 2000
Timer1.Interval = 300

End If
End Sub


Private Sub Form_DblClick()
 Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Click()
'If Image12.Visible = True Or Image13.Visible = True Or Image5.Visible = True Or Image4.Visible Or Image7.Visible = True Then
Image12.Visible = False
Image13.Visible = False
Image4.Visible = False
Image5.Visible = False
Image7.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Toolbar1.Visible = False
'\\'If Then
'If Command1.Default Then
'Screen.MousePointer = vbCustom
'End If
'If Image4.Visible = True Then
'If Image5.Visible = True Then
'If year.Caption >= 1800 Then
'Command2.Visible = True
'Else
'If year.Caption <= 1800 Then
'Command2.Visible = False
'End If
'End If
'End If
'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyCode = vbKeyP Then
'mnupassword.Enabled = True

'End If
If KeyAscii = 18 Then
'Dim Form3 As New Form3
Form3.Show 'vbModal, Me
End If
If KeyAscii = 27 Then
End
End If

'I'f KeyAscii = 112 Then
'frmAbout.Show
'End If

If KeyAscii = 19 Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Timer5.Interval = 0
Timer6.Interval = 0
Timer7.Interval = 0
Timer8.Interval = 0
Timer9.Interval = 0
Timer10.Interval = 0
Timer11.Interval = 0
Timer12.Interval = 0
Timer13.Interval = 0
Timer14.Interval = 0
Timer15.Interval = 0
Timer16.Interval = 0
Timer17.Interval = 0
Timer18.Interval = 0
Timer19.Interval = 0
End If
If KeyAscii = 80 Then
Timer1.Interval = 300
Timer2.Interval = 10000
Timer3.Interval = 300
Timer4.Interval = 75
Timer5.Interval = 550
Timer6.Interval = 3100
Timer7.Interval = 1000
Timer8.Interval = 10000
Timer9.Interval = 1000
Timer10.Interval = 30
Timer11.Interval = 20
Timer12.Interval = 70
Timer13.Interval = 75
Timer14.Interval = 300
Timer15.Interval = 50
Timer16.Interval = 40
Timer17.Interval = 50
Timer18.Interval = 400
Timer19.Interval = 2000

End If
End Sub

Private Sub Form_Load()
form2.KeyPreview = True
'If Form9.Label7.Visible = True Then
'Call Form9.jv
'form2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\back1.bmp")
  ProgressBar1.Visible = False                          '("c:\windows\desktop\vb projects\war\com.bmp")
'End If
If Label5.Caption = "1a" Then
Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\townsnow.bmp")
End If
Call Map2
Call Module1.tree
'Sub Initialize_Program()
'Call GameEngine.Initialize_GameEngine
'Call GameEngine.Initialize_GameWindows
'CHARACTER_QOUTE = Chr$(34)
'GameEngine.GameDataFile$ = Directory_DefaultGameData
'Call GameEngine.Initialize_BattleView

End Sub
Public Sub Map2()
If Label5.Caption = "1c" Then
Image16.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t2.bmp")
Image150.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t1.bmp")
Image48.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t2.bmp")
Image91.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t1.bmp")
Image16.OLEDragMode = 0
Image150.OLEDragMode = 0
Image48.OLEDragMode = 0
Image91.OLEDragMode = 0

End If

End Sub
Private Sub HScroll1_Change()
'Picture37.Left = -HScroll1.Value
Image33.Left = -HScroll1.Value
End Sub

Private Sub hscmap_Change()
Module1.Map.CurrentX = Me.hscmap.Value
    
    ' Redraw the map
    Call DrawMap
End Sub

Private Sub Image1_Click()
Image22.Visible = True
End Sub
Private Sub image1_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
Image1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\com.bmp")
i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)
End Sub

Private Sub Image10_Click()
'i = sndPlaySound(ByVal CStr(App.Path & "\SKING1.wav"), SND_ASYNC)
'Toolbar1.Visible = True
'Image10.Picture = LoadPicture("c:\windows\desktop\vb projects\war\TAN.bmp")
'Image12.Visible = True
Image22.Visible = True
End Sub

Private Sub Image10_DblClick()
Image22.Visible = False

End Sub

Private Sub Image11_Click()
Image10.Visible = True
'
End Sub

'Private Sub Image14_Click()
'If Image12.Visible = True Then
'Image12.Visible = False
'End If
Private Sub image14_KeyDown(KeyCode As Integer, _
Shift As Integer)
  If KeyCode = vbKeyA Then
    Image12.Visible = False
    'MsgBox "You pressed _ the A key."
'i = sndPlaySound(ByVal CStr(App.Path & "\Taunt009.wav"), SND_ASYNC)
End If
'End Sub
End Sub

Private Sub Image11_DblClick()
Image10.Visible = False
End Sub

Private Sub Image14_Click()
Image5.Visible = True
End Sub

Private Sub Image15_Click()
Image13.Visible = True
End Sub

Private Sub Image16_Click()
Image13.Visible = True
End Sub

Private Sub Image17_Click()
Image13.Visible = True
End Sub

Private Sub Image18_Click()


Image42.Visible = True
End Sub

Private Sub Image18_DblClick()
Image42.Visible = False
End Sub

Private Sub Image19_Click()
'Image20.Visible = True
Image18.Visible = True
End Sub

Private Sub Image19_DblClick()
Image18.Visible = False
End Sub

Private Sub Image22_Click()
Image19.Visible = True
End Sub

Private Sub Image22_DblClick()
Image19.Visible = False
End Sub

Private Sub Image25_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
 
 'I'f Picture19.Drag Then

 'Call a
' Else
 
 

 

'Picture19 =
'
 Source.Drag vbEndDrag
Source.Visible = True

'Picture4.Visible = True
Image25.Picture = LoadPicture("C:\windows\desktop\vb projects\war\war.bmp")
i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)

'End If
'End If
End Sub
Public Sub a()
 
'Picture19.Drag vbEndDrag
'Picture19.Visible = True
'Image25.Picture = LoadPicture("C:\windows\desktop\vb projects\war\b3.bmp")
'i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)
End Sub

Private Sub Image26_Click()
Image5.Visible = True
End Sub

Private Sub Image26_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Image26.Visible = True
Image26.Enabled = True
End Sub

Private Sub Image29_Click()
Image5.Visible = True
End Sub

Private Sub Image33_Click()
Image12.Visible = False
Image13.Visible = False
Image4.Visible = False
Image5.Visible = False
Image7.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Toolbar1.Visible = False
End Sub

Private Sub Image34_Click()
' For i = 1 To 250000
    'Next i
    'score = score + 10
   ' Call choosepic(1)
    'Image34.Left = -315
    'done = False
   ' Image7.Visible = True
End Sub

Private Sub Image42_Click()
Image43.Visible = True
End Sub

Private Sub Image42_DblClick()
Image43.Visible = False
End Sub

Private Sub Image43_Click()
Image18.Visible = True

End Sub

Private Sub Image43_DblClick()
Image18.Visible = False
End Sub

'Private Sub Image6_Click()

Private Sub Image6_Click()
   ' enemy1.Picture = explode.Picture
    For i = 1 To 250000
    Next i
    'score = score + 10
   ' Call choosepic(1)
    Image6.Left = -315
    done = False
    Image7.Visible = True
End Sub

Private Sub Image8_Click()
Image21.Visible = True
End Sub

Private Sub image8_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Image8.Picture = LoadPicture("c:\windows\desktop\vb projects\war\BUILDINGS\sam.bmp")
i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)
End Sub

Private Sub Image9_Click()
Image13.Visible = True
End Sub
Private Sub image9_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Image9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\BUILDINGS\b1.bmp")
i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)

'End Sub
End Sub

Private Sub Label2_Click()
MsgBox "cheater"
money = money + 5000
Power = Power + 1000
year = year + 1000
End Sub

Private Sub mnmain_Click()
Dim Form3 As New Form3
Form3.Show vbModal, Me

End Sub

Private Sub MNUABOUT_Click()
Dim frmAbout As New frmAbout
frmAbout.Show vbModal, Me


End Sub

Private Sub mnucontents_Click()
    'Dim nRet As Integer
    CommonDialog1.HelpFile = "C:\windows\desktop\war.hlp"
    CommonDialog1.HelpCommand = cdlHelpContents
 CommonDialog1.ShowHelp

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    'If Len(App.HelpFile) = 0 Then
       ' MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    'Else
        'On Error Resume Next
       ' nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        'If Err Then
         '   MsgBox Err.Description
        'End If
    'End If
End Sub


Private Sub Mnurestart_Click()
'MsgBox "Are You sure you want to restart", vbYesNo
'If vbYes Then
'End
'End If
Call restart
End Sub

Private Sub mnuSearch_Click()


    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuexit_Click()
' response = MsgBox "Are you sure you want to quit", vbYesNo
'I'f vbYes Then
'End
'Else
Call pgain
'End If
End Sub

Private Sub mnuload_Click()
CommonDialog1.ShowOpen

End Sub


Private Sub mnusave_Click()
CommonDialog1.ShowSave
'cdlOFNFileMustExist 'the string are delimited by spaces.
'cdlOFNOFN
'CommonDialog1
'cdlOFNOverwritePrompt
'C'ommonDialog1. = createprompt
'CommonDialog1.Flags = cdlOFNOverwritePrompt And cdlOFNCreatePrompt
'OverwritePrompt
CommonDialog1.Flags = &H2& Or &H2000&
CommonDialog1.Filter = ".war"

End Sub

Private Sub mnuset_Click()
Form8.Show vbModal, Me

End Sub



Private Sub mnuspace_Click()
'If year.Caption <= 1950 Then
'mnuspace.Enabled = True
'E'nd If
If Form10.Label16.Caption = "0" Then

Form10.Show
Else
Form10.WindowState = 0

End If


End Sub

Private Sub mnut1_Click()
MediaPlayer1.FileName = "C:\WINDOWS\Desktop\vb projects\war\jpi.mid"


End Sub

Private Sub mnut2_Click()
MediaPlayer1.FileName = "C:\WINDOWS\Desktop\vb projects\war\war.mid"

End Sub

Private Sub mnutip_Click()
frmTip.Show vbModal, Me
End Sub

Private Sub mnutrack3_Click()
MediaPlayer1.FileName = "C:\WINDOWS\Desktop\vb projects\war\music7.mid"
End Sub

Private Sub mnutrack4_Click()
MediaPlayer1.FileName = "C:\WINDOWS\Desktop\vb projects\war\music9.mid"
End Sub

Private Sub Picture1_Click()
ProgressBar1.Top = 1800
Power = Power + 0
Picture2.Visible = True
money = money - 1000
'Picture1.DragMode = 1
'Picture1.OLEDragMode = 1
'Picture1.OLEDropMode = 2
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture1.OLEDragMode = 1
    Timer19.Enabled = True
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If

End Sub

'Private Sub Picture1_Click()
'Private Sub picture1_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

'    Picture1.Drag vbBeginDrag
'    Set Picture1.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico") '("c:\Program _
        files\ Microsoft Visual _
       Basic\Icons\Dragdrop\Dragfldr.ico")
'End Sub
'do
'Power = Power + 0
'Picture2.Visible = True
'money = money - 1000

'If money = 6000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 12000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'End If
'End If
'End If




'Else

'If grid = False Then
'jv sound
'End If
'End If
'End Sub
Private Sub picture1_dblclick()

i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub



Private Sub Picture10_Click()
ProgressBar1.Top = 4800
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture10.ToolTipText = "Radar1"
    
money = money + 1500
Power = Power + 20
Picture10.OLEDragMode = 1
Call Power1
End Sub

'Private Sub Picture10_Click()
'Private Sub picture10_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

   ' Picture10.Drag vbBeginDrag
    'Set Picture10.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'money = money - 1500
''Call Power1
'If money = 6000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'i 'f money = 12000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'End If
'End If
'End If
'money = money - 100
'End Sub
Private Sub picture10_dblclick()
ProgressBar1.Top = 4800
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Picture11_Click()
money = money - 100
End Sub

Private Sub Picture12_Click()
money = money - 100
End Sub

Private Sub Picture13_Click()
money = money - 100
End Sub

Private Sub Picture14_Click()
money = money - 100
End Sub

Private Sub Picture15_Click()
money = money - 100
'Picture37.PaintPicture Picture36, X1, Y1
'y1 '= 1.61396, Y = 2.460628)' = Picture36
', 2.460628,)
'X  1.61396, Y 2.460628 ', 'width1, height1, x2, y2, width2,object.PaintPicture picture, x1, y1, width1, height1, x2, y2, width2,
End Sub

Private Sub Picture15_Paint()
'X1095 , Y975

End Sub

Private Sub Picture16_Click()
money = money - 100
Image10.Picture = LoadPicture("c:\windows\desktop\vb projects\war\TAN.bmp")
Image10.Visible = True
End Sub

Private Sub Picture17_Click()
ProgressBar1.Top = 3840
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture17.OLEDragMode = 1
    Timer19.Enabled = True
money = money - 1500
Power = Power + 20
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

'Private Sub picture17_MouseDown(Button As Integer, _
'Shift As Integer, X As Single, Y As Single)
'1End Sub
        

'If grid = True Then

  '  Picture17.Drag vbBeginDrag
   'Set Picture17.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'Private Sub picture17_click() 'DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
 ' money = money - 1500
'Power = Power - 20


'End Sub
Private Sub picture17_dblclick()
'ProgressBar1.Top = 3840
'money = money + 1500
'Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub picture18_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

    Picture18.Drag vbBeginDrag
    Set Picture18.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'Private Sub picture18_click() 'DragOver(Source As Control,
money = money - 1500
Power = Power - 20
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub
Private Sub picture18_dblclick()
ProgressBar1.Top = 2760
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

'Private Sub picture19_click() 'DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)

Private Sub Picture19_Click()
ProgressBar1.Top = 1800
   money = money - 1500
Power = Power - 20
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture19.OLEDragMode = 1
    Timer19.Enabled = True
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub
  ' Private Sub picture19_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
'If grid = True Then
   ' Picture19.Drag vbBeginDrag
   ' Set Picture19.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'End If
   'Source.Drag vbEndDrag
    'Source.Visible = True
   ' Picture19.Picture = LoadPicture("c:\windows\desktop\vb projects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _
'End Sub

Private Sub picture19_dblclick()
Power = Power + 20
money = money + 1500
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Picture2_Click()
ProgressBar1.Top = 2760
Power = Power + 100
'Picture3.Visible = True
'Picture4.Visible = True
'Picture5.Visible = True
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
Picture7.Visible = True
Picture26.Visible = True
Picture28.Visible = True
Picture2.OLEDragMode = 1
Timer19.Enabled = True
'Picture2.OLEDragMode = 1
'Picture2.OLEDropMode = 2
money = money - 200
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If

End Sub

Private Sub picture2_dblclick()
money = money + 1500
Power = Power - 100
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub


Private Sub picture20_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture20.Picture = LoadPicture("c:\windows\desktop\vb projects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub


Private Sub picture21_click() 'DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
  money = money - 100
End Sub
Private Sub picture21_dblclick()
money = money + 100

i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub picture22_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture22.Picture = LoadPicture("c:\windows\desktop\vb projects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub

Private Sub Image3_Click()
Image23.Visible = True
End Sub
'Private Sub image3_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
  ' Source.Drag vbEndDrag
 '   Source.Visible = True
 '   Image3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\plant.bmp")
'i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)
'End Sub

Private Sub Picture24_Click()
money = money - 150
'i = sndPlaySound(ByVal CStr(App.Path & "\slcttgt1.wav"), SND_ASYNC)
'Power = Power - 20
'Call Power1

MsgBox "Support air craft on the way"
Timer12.Enabled = True
Image34.Visible = True
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
'Picture12.Visible = True
'Picture11.Visible = True
'Picture10.Visible = True
'Picture25.Visible = True
'Picture27.Visible = True
'Picture29.Visible = True

End Sub

Private Sub Picture25_Click()
money = money - 5000
End Sub

'Private Sub Picture26_Click()
''P'rivate Sub picture26_MouseDown(Button As Integer, _
 '       Shift As Integer, X As Single, Y As Single)
 '   Picture26.Drag vbBeginDrag
   'S'et Picture26.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
''money = money - 1500
''ower = Power - 20
'
''Call Power1
'''If money = 6000 Then
''If money = 12000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
''Picture16.Visible = True
''Picture15.Visible = True
'Picture14.Visible = True
'Picture13.Visible = True
''Picture12.Visible = True
'Picture11.Visible = True

''End If
'End If
'End If



Private Sub Picture26_Click()
ProgressBar1.Top = 4800
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
Picture16.Visible = True
Picture15.Visible = True
Picture14.Visible = True
Picture13.Visible = True
Picture12.Visible = True
Picture26.OLEDragMode = 1
Timer19.Enabled = True
'Picture11.Visible = True
money = money + 1500
Power = Power + 20
End Sub

Private Sub picture26_dblclick()

i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Picture27_Click()
ProgressBar1.Top = 7800
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
money = money - 20
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture28_Click()
ProgressBar1.Top = 3840
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture28.OLEDragMode = 1
    Timer19.Enabled = True
money = money + 1500
Power = Power - 200
End Sub

Private Sub Picture3_Click()
ProgressBar1.Top = 7800
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture3.OLEDragMode = 1
    Timer19.Enabled = True
money = money - 1000
'If grid = True Then
'do
Power = Power - 20
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture30_Click()
ProgressBar1.Top = 6840
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
money = money - 35
'Power = Power - 20
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

'Private Sub Picture30_Click()
'Private Sub picture30_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
'If grid = True Then
'
   ' Picture30.Drag vbBeginDrag
   ' Set Picture30.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'End Sub
Private Sub picture30_dblclick()
Power = Power + 20
money = money + 1500
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

'Private Sub picture27_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
'If grid = True Then
   ' Picture27.Drag vbBeginDrag
    'Set Picture27.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'money = money - 25
'Power = Power - 20
'money = money - 100
'End Sub
Private Sub picture27_dblclick()
Power = Power + 20
money = money + 1500
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub
'Private Sub Picture28_Click()
'Private Sub picture28_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
    'Picture28.Drag vbBeginDrag
    'Set Picture28.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico") '("c:\Program _

'money = money - 1500
'Power = Power + 200
' Power1
'If money = 6000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 12000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'End If
'End If
'End If

'End Sub
Private Sub picture28_dblclick()

i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Picture29_Click()
money = money - 100
End Sub

'Private Sub Picture3_Click()
'Private Sub picture3_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

    'Picture1.Drag vbBeginDrag
    'Set Picture3.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico") '("c:\Program _



'End If
'Else

'If grid = False Then
'jv sound
'End If
'End Sub
Private Sub picture3_dblclick()
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Picture31_Click()
ProgressBar1.Top = 5760
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
Call Power1
money = money - 20
End Sub

'Private Sub Picture31_Click()
'Private Sub picture31_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

    'Picture31.Drag vbBeginDrag
    'Set Picture31.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'money = money - 50
'Power = Power - 20
'Call Power1
'If money = 6000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 12000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
''End If
'End If
'End If
'End Sub
Private Sub picture31_dblclick()
p
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Picture35_dblClick()
money = money + 500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub
'Private Sub picture35_MouseDown(Button As Integer, _
       ' Shift As Integer, X As Single, Y As Single)
    'Picture35.Drag vbBeginDrag
   '' Set Picture35.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
Private Sub picture35_click()
ProgressBar1.Top = 4800
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture35.OLEDragMode = 1
    Timer19.Enabled = True
money = money - 500
Power = Power - 20
'Picture11.Visible = True
'Picture29.Visible = True
'Picture35.OLEDragMode = 1
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
Image8.Visible = True



End Sub

Private Sub Picture36_Click()
ProgressBar1.Top = 3840
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture36.OLEDragMode = 1
    Timer19.Enabled = True
money = money - 500
Power = Power + 20
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture4_Click()
ProgressBar1.Top = 6840
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture4.OLEDragMode = 1
    Timer19.Enabled = True
money = money - 1000
Power = Power - 20
Picture11.Visible = True
Picture29.Visible = True



Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

'Private Sub Picture4_Click()
'Private Sub picture4_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
  '  Picture4.Drag vbBeginDrag
  '  Set Picture4.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")

'money = money - 500
'Power = Power - 20
'Picture11.Visible = True
'Picture29.Visible = True



'Call Power1
'If money = 6000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 12000 Then
'i' = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'End If
'End If
'End If


'End Sub
Private Sub picture4_dblclick()
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Picture5_Click()
'If grid = True Then
'do
ProgressBar1.Top = 6840
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture5.OLEDragMode = 1
    Timer19.Enabled = True
Power = Power - 20
money = money - 100
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
'End If
'Else
'If grid = False Then
'jv sound
'End If
End Sub
Private Sub picture5_dblclick()
money = money + 100
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

'Private Sub Picture6_Click()
'Private Sub picture6_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

   ' Picture6.Drag vbBeginDrag
   ' Set Picture6.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'money = money - 1500
'Power = Power - 20
'Call Power1
'If money = 6000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 12000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'End If
'End If
'End If
'Picture24.Visible = True
'Picture26.Visible = True
'Picture28.Visible = True


End Sub

Private Sub Picture6_Click()
ProgressBar1.Top = 7800
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture6.OLEDragMode = 1
    Timer19.Enabled = True
money = money - 1000
Power = Power - 20
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub picture6_dblclick()
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub

Private Sub Image2_Click()
Image5.Visible = True
End Sub
Private Sub image2_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Image2.Visible = True
   Source.Drag vbEndDrag
    Source.Visible = True
    Image2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")       '( "c:\Program _
PlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)
       'files\Microsoft Visual _
       ' Basic\Icons\Office\Files03a.ico")
'End Sub
'PlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)

i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC)

End Sub


Private Sub Picture7_Click()
ProgressBar1.Top = 5760
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture7.OLEDragMode = 1
    Timer19.Enabled = True
money = money - 1500
Power = Power - 20
Picture5.Visible = True
Call Power1
If money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

'Private Sub picture7_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)
   ' Picture7.Drag vbBeginDrag
   'Set Picture7.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")

''Picture3.Visible = True
'Picture4.Visible = True


'Picture6.Visible = True
'End Sub
Private Sub picture7_dblclick()
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub



Private Sub Picture9_Click()
ProgressBar1.Top = 5760
Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    Picture9.OLEDragMode = 1
    Timer19.Enabled = True
Call Power1
money = money - 500
Power = Power - 20
End Sub

'Private Sub picture9_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

   ' Picture9.Drag vbBeginDrag
    'Set Picture9.DragIcon = LoadPicture("c:\windows\desktop\vb projects\war\trash02a.ico")
'If gri
'money = money - 1500
'Power = Power - 20
'Call Power1
'If money = 6000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 12000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'If money = 24000 Then
'i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
'End If
'End If
'End If
'E'nd Sub
Private Sub picture9_dblclick()
money = money + 1500
Power = Power + 20
i = sndPlaySound(ByVal CStr(App.Path & "\cancld1.wav"), SND_ASYNC)
End Sub



Private Sub Picture8_Click()
money = money - 550

End Sub

'Private Sub Text1_Change()

'End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'If vbKeyS Then
'form2.MousePointer = vbCustom
'End If
'If vbKeyA Then
'form2.MousePointer = vbDefault
'End If

'End Sub


Private Sub Timer1_Timer()
If year.Caption = 300 Or year.Caption = 310 Or year.Caption = 320 Or year.Caption = 330 Or year.Caption = 340 Or year.Caption = 350 Or year.Caption = 360 Or year.Caption = 370 Then
Image35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\rain.ico")
Image36.Picture = LoadPicture("c:\windows\desktop\vb projects\war\rain.ico")
Image37.Picture = LoadPicture("c:\windows\desktop\vb projects\war\rain.ico")
Image38.Picture = LoadPicture("c:\windows\desktop\vb projects\war\rain.ico")
Image39.Picture = LoadPicture("c:\windows\desktop\vb projects\war\rain.ico")
Image40.Picture = LoadPicture("c:\windows\desktop\vb projects\war\rain.ico")
Command16.Caption = "r"

End If
If year.Caption = 380 Then
Image35.Visible = False
Image36.Visible = False
Image37.Visible = False
Image38.Visible = False
Image39.Visible = False
Image40.Visible = False
End If


If year.Caption = 500 Then
Image35.Visible = True
Image36.Visible = True
Image37.Visible = True
Image38.Visible = True
Image39.Visible = True
Image40.Visible = True
End If

If year.Caption = 570 Then
Image35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cloud.ico")
Image36.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cloud.ico")
Image37.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cloud.ico")
Image38.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cloud.ico")
Image39.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cloud.ico")
Image40.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cloud.ico")
Command16.Caption = "c"

End If
'If Interval = 125 Then
'year = year + 10
'End If
'If Timer1.Interval = 300 Then

'If year.Caption >= 1900 Then
'Picture24.Visible = True
'Timer12.Enabled = True
'Image33.Visible = True
'End If
'End If
'Call TIMER
year = year + 10
If year.Caption >= 1950 Then
mnuspace.Enabled = True
End If
If Picture15.Visible = True Then
'Picture18.Visible = False
End If
If year.Caption >= 9000 Then
mnupassword.Enabled = True
End If
'i 'f year.Caption >= 1909 Then
'Image18.Visible = False
'End If
If year.Caption >= 1900 Then
Command12.Enabled = True
End If
If KeyCode = vbKeyMenu Then
Dim Form3 As New Form3
Form3.Show vbModal, Me
End If
If KeyCode = vbKeyEscape Then
End
End If
If KeyCode = vbKeyPause Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Timer5.Interval = 0
Timer6.Interval = 0
Timer7.Interval = 0
Timer8.Interval = 0
Timer9.Interval = 0
Timer10.Interval = 0
End If
If KeyCode = vbKeyP Then
Timer1.Interval = 300
Timer2.Interval = 300
Timer3.Interval = 300
Timer4.Interval = 300
Timer5.Interval = 300
Timer6.Interval = 300
Timer7.Interval = 300
Timer8.Interval = 300
Timer9.Interval = 300
Timer10.Interval = 300
End If
'If year.Caption >= 2050 Then
'Image14.Visible = True
'Image2.Visible = False
'Image16.Visible = True
'Image15.Visible = False
'Image9.Visible = False
'Image17.Visible = False
'i ' 'mage18.Visible = False
'End If
'If Image9.Visible = True Then

'C 'all b
'If year.Caption >= 500 Then
'Image17.Visible = True
'Image9.Visible = False
'End If
''stuff with '' may be neede
'If year.Caption >= 1910 Then
'Image15.Visible = True
'Image9.Visible = False
'Image17.Visible = False
'Image18.Visible = True
'Image18.Enabled = False
'Image18.Visible = False
'Image9.Visible = False
'End If

'If year.Caption >= 1000 Then
'Image18.Visible = True
'Image17.Visible = False
'Image9.Visible = False
'End If

'End If
'End If


End Sub

Private Sub Timer10_Timer()
If Label5.Caption = "1c" And Label2.Caption = "Year2" Then
Picture1.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\town2.bmp")
Picture2.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\h2.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g2.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p2.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov2.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t2.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\mar2.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b2.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s2.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\a2.bmp")
End If
If Label5.Caption = "1c" And Label2.Caption = "Year1" Then
Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\town1.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p1.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t1.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g1.bmp")
Picture2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\h1.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\market.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov1.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b1.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s1.bmp")
Picture4.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\siege.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\a1.bmp")
Picture7.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\acadmenby.bmp")
Picture19.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\farm.bmp")
If Label5.Caption = "2c" And Label2.Caption = "Year1" Then
Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\town1.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p1.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t1.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g1.bmp")
Picture2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\h1.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\market.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov1.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b1.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s1.bmp")
Picture4.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\siege.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\townsnow.bmp")
Picture19.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\farm.bmp")
End If
End If
If Label5.Caption = "2c" And Label2.Caption = "year1" Then
Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\town1.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p1.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t1.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g1.bmp")
Picture2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\h1.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\market.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov1.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b1.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s1.bmp")
Picture4.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\siege.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\a1.bmp")
Picture19.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\farm.bmp")
End If
If Label5.Caption = "1a" Then
Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\townsnow.bmp")
End If
'Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\town1.bmp")
If Label3.Caption = "loser" Then
i = sndPlaySound(ByVal CStr(App.Path & "\misnlst1.wav"), SND_ASYNC)
Label3.Caption = "you lost"
Dim Form3 As New Form3
Unload Me
Form3.Show vbModal, Me

End If
If Label3.Caption = "winwin" Then
i = sndPlaySound(ByVal CStr(App.Path & "\misnwon1.wav"), SND_ASYNC)
If Label5.Caption = "1a" Then
Dim Form9 As New Form9
Form9.Label1.Caption = "Mission 2"
Form9.Label2.Visible = False
Form9.Label8.Visible = True
Form9.Label8.Caption = "stop the german advance kill all troops and buildings"

Unload Me
Form9.Show vbModal, Me

End If

End If
If Label3.Caption = hohoho Then

End If
'Picture36.Left = 6000
'Picture36.Height = 200
'Next i
'Picture36.Left = 5000
'Picture36.Height = 500




End Sub

Private Sub Timer11_Timer()
If Label2.Caption = "year10" Then
Command1.Enabled = True
End If

'Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\town1.bmp")
If Label5.Caption = "1a" Then
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\com.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\sam.bmp")
Picture6.Picture = Form90.Image9.Picture
'LoadPicture("c:\windows\desktop\vb projects\war\tech.bmp")
Picture36.Picture = LoadPicture("c:\windows\desktop\vb projects\war\sam.bmp")
End If
If Image155.Visible = False Then
Picture5.Picture = Form90.Image1.Picture  'LoadPicture("c:\windows\desktop\vb projects\war\silo.bmp")

End If
If Image37.Left >= 8880 Then
Image37.Left = 720
End If
If Image38.Left >= 8880 Then
Image38.Left = 840
End If
If Image40.Left >= 8880 Then
Image40.Left = 120
End If
If Image35.Left <= 0 Then
Image35.Left = 8400
End If
If Image36.Left <= 0 Then
Image36.Left = 8400
End If
If Image39.Left <= 0 Then
Image39.Left = 8400
End If
End Sub

Private Sub Timer12_Timer()
Image34.Left = Image34.Left - 90
End Sub

Private Sub Timer13_Timer()
Image35.Visible = True
Image36.Visible = True
Image37.Visible = True
Image38.Visible = True
Image35.Left = Image35.Left - 90
Image36.Left = Image35.Left - 90
Image37.Left = Image37.Left + 90
Image38.Left = Image38.Left + 90
If Picture7.Visible = True Then
Picture5.Visible = True
Picture3.Visible = True
End If
If Image27.Visible = True Then
Image35.Visible = False
Image36.Visible = False
Image37.Visible = False
Image38.Visible = False
Image39.Visible = False
Image40.Visible = False
End If
If Image27.Visible = False Then
Image35.Visible = True
Image36.Visible = True
Image37.Visible = True
Image38.Visible = True
Image39.Visible = True
Image40.Visible = True
End If


End Sub

Private Sub Timer14_Timer()
Image39.Visible = True
Image40.Visible = True
Image39.Left = Image39.Left - 90
Image40.Left = Image40.Left + 90
'MediaPlayer2.FileName = "C:\WINDOWS\Desktop\vb projects\war\wind3.wav"
'MediaPlayer2.FileName = "C:\WINDOWS\Desktop\vb projects\war\tf3.wav"

End Sub


Private Sub Timer15_Timer()
If Label5.Caption = "1c" Then
Call Map2
Timer15.Enabled = False
'Image16.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t2.bmp")
'Image150.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t1.bmp")
'Image48.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t2.bmp")
'Image91.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t1.bmp")
'Image16.OLEDragMode = 0
'Image150.OLEDragMode = 0
'Image48.OLEDragMode = 0
'Image91.OLEDragMode = 0
End If
If Label5.Caption = "1a" Then

Image136.Picture = LoadPicture("c:\windows\desktop\vb projects\war\tr1.bmp")
Image51.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t.bmp")
Image46.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t.bmp")
Image72.Picture = LoadPicture("c:\windows\desktop\vb projects\war\tr1.bmp")
Image136.OLEDragMode = 0
Image51.OLEDragMode = 0
Image46.OLEDragMode = 0
Image72.OLEDragMode = 0
Timer15.Enabled = False
End If
End Sub

Private Sub Timer16_Timer()

Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viechels\alein\commabnfd.bmp")
Picture2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\viechels\alein\o.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\w1.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\q6.bmp")
Picture7.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\q.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\w3.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\alien\q4.bmp")

End Sub

Private Sub Timer17_Timer()
Call silo
Call tech
Call com
Call plant
Call ORE

If Image43.Height = 750 And Image43.Width = 1095 Then
Image43.ToolTipText = "War Factory"
If Image18.Height = 495 And Image18.Width = 975 Then
Image18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image18.Visible = False
End If

End If
If Image42.Height = 750 And Image42.Width = 1095 Then
Image42.ToolTipText = "War Factory"
If Image43.Height = 495 And Image43.Width = 975 Then
Image43.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image43.Visible = False
End If
End If

If Image18.Height = 750 And Image18.Width = 1095 Then
Image18.ToolTipText = "War Factory"
If Image42.Height = 495 And Image42.Width = 975 Then
Image42.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image42.Visible = False
End If
End If
If Image19.Height = 750 And Image19.Width = 1095 Then
Image19.ToolTipText = "War Factory"
If Image18.Height = 495 And Image18.Width = 975 Then
Image18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image18.Visible = False
End If
End If

If Image22.Height = 750 And Image22.Width = 1095 Then

Image22.ToolTipText = "War Factory"
If Image19.Height = 495 And Image19.Width = 975 Then
Image19.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image19.Visible = False
End If
End If

If Image10.Height = 750 And Image10.Width = 1095 Then
Image10.ToolTipText = "War Factory"
If Image22.Height = 495 And Image22.Width = 975 Then
Image22.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image22.Visible = False
End If
End If

If Image11.Height = 750 And Image11.Width = 1095 Then
Image11.ToolTipText = "War Factory"
If Image10.Height = 495 And Image10.Width = 975 Then
Image10.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image10.Visible = False
End If
End If

If Image21.Height = 750 And Image21.Width = 1095 Then
Image21.ToolTipText = "War Factory"
End If
If Image26.Height = 750 And Image26.Width = 1095 Then
Image26.ToolTipText = "War Factory"
End If
If Image23.Height = 750 And Image23.Width = 1095 Then
Image23.ToolTipText = "War Factory"
End If
If Image20.Height = 750 And Image20.Width = 1095 Then
Image20.ToolTipText = "War Factory"
End If
If Image24.Height = 750 And Image24.Width = 1095 Then
Image24.ToolTipText = "War Factory"
End If
If Image2.Height = 750 And Image2.Width = 1095 Then
Image2.ToolTipText = "War Factory"
End If
If Image3.Height = 750 And Image3.Width = 1095 Then
Image3.ToolTipText = "War Factory"
End If
If Image5.Height = 750 And Image5.Width = 1095 Then
Image5.ToolTipText = "War Factory"
End If
If Image4.Height = 750 And Image4.Width = 1095 Then
Image4.ToolTipText = "War Factory"
End If
If Image6.Height = 750 And Image6.Width = 1095 Then
Image6.ToolTipText = "War Factory"
End If
If Image1.Height = 750 And Image1.Width = 1095 Then
Image1.ToolTipText = "War Factory"
End If
If Image17.Height = 750 And Image17.Width = 1095 Then
Image17.ToolTipText = "War Factory"
End If
If Image12.Height = 750 And Image12.Width = 1095 Then
Image12.ToolTipText = "War Factory"
End If
If Image8.Height = 750 And Image8.Width = 1095 Then
Image8.ToolTipText = "War Factory"
End If
If Image7.Height = 750 And Image7.Width = 1095 Then
Image7.ToolTipText = "War Factory"
End If
If Image15.Height = 750 And Image15.Width = 1095 Then
Image15.ToolTipText = "War Factory"
End If
If Image14.Height = 750 And Image14.Width = 1095 Then
Image14.ToolTipText = "War Factory"
End If
If Image9.Height = 750 And Image9.Width = 1095 Then
Image9.ToolTipText = "War Factory"
End If
If Image16.Height = 750 And Image16.Width = 1095 Then
Image16.ToolTipText = "War Factory"
End If
If Image13.Height = 750 And Image13.Width = 1095 Then
Image13.ToolTipText = "War Factory"
End If
If Image28.Height = 750 And Image28.Width = 1095 Then
Image28.ToolTipText = "War Factory"
End If
If Image29.Height = 750 And Image29.Width = 1095 Then
Image29.ToolTipText = "War Factory"
End If
If Image32.Height = 750 And Image32.Width = 1095 Then
Image32.ToolTipText = "War Factory"
End If
If Image133.Height = 750 And Image133.Width = 1095 Then
Image133.ToolTipText = "War Factory"
End If
If Image86.Height = 750 And Image86.Width = 1095 Then
Image86.ToolTipText = "War Factory"
End If
If Image126.Height = 750 And Image126.Width = 1095 Then
Image126.ToolTipText = "War Factory"
End If
If Image151.Height = 750 And Image151.Width = 1095 Then
Image151.ToolTipText = "War Factory"
End If
If Image69.Height = 750 And Image69.Width = 1095 Then
Image69.ToolTipText = "War Factory"
End If
If Image96.Height = 750 And Image96.Width = 1095 Then
Image96.ToolTipText = "War Factory"
End If
If Image78.Height = 750 And Image78.Width = 1095 Then
Image78.ToolTipText = "War Factory"
End If
If Image88.Height = 750 And Image88.Width = 1095 Then
Image88.ToolTipText = "War Factory"
End If
If Image66.Height = 750 And Image66.Width = 1095 Then
Image66.ToolTipText = "War Factory"
End If
If Image118.Height = 750 And Image118.Width = 1095 Then
Image118.ToolTipText = "War Factory"
End If
If Image72.Height = 750 And Image72.Width = 1095 Then
Image72.ToolTipText = "War Factory"
End If
If Image110.Height = 750 And Image110.Width = 1095 Then
Image110.ToolTipText = "War Factory"
End If
If Image135.Height = 750 And Image135.Width = 1095 Then
Image135.ToolTipText = "War Factory"
End If
If Image57.Height = 750 And Image57.Width = 1095 Then
Image57.ToolTipText = "War Factory"
End If
If Image77.Height = 750 And Image77.Width = 1095 Then
Image77.ToolTipText = "War Factory"
End If
If Image85.Height = 750 And Image85.Width = 1095 Then
Image85.ToolTipText = "War Factory"
End If
If Image83.Height = 750 And Image83.Width = 1095 Then
Image83.ToolTipText = "War Factory"
End If
If Image56.Height = 750 And Image56.Width = 1095 Then
Image56.ToolTipText = "War Factory"
End If
If Image93.Height = 750 And Image93.Width = 1095 Then
Image93.ToolTipText = "War Factory"
End If
If Image107.Height = 750 And Image107.Width = 1095 Then
Image107.ToolTipText = "War Factory"
End If
If Image94.Height = 750 And Image94.Width = 1095 Then
Image94.ToolTipText = "War Factory"
End If
If Image30.Height = 750 And Image30.Width = 1095 Then
Image30.ToolTipText = "War Factory"
End If
If Image100.Height = 750 And Image100.Width = 1095 Then
Image100.ToolTipText = "War Factory"
End If
If Image101.Height = 750 And Image101.Width = 1095 Then
Image101.ToolTipText = "War Factory"
End If
If Image84.Height = 750 And Image84.Width = 1095 Then
Image84.ToolTipText = "War Factory"
End If
If Image54.Height = 750 And Image54.Width = 1095 Then
Image54.ToolTipText = "War Factory"
End If
If Image129.Height = 750 And Image129.Width = 1095 Then
Image129.ToolTipText = "War Factory"
End If
If Image76.Height = 750 And Image76.Width = 1095 Then
Image76.ToolTipText = "War Factory"
End If
If Image134.Height = 750 And Image134.Width = 1095 Then
Image134.ToolTipText = "War Factory"
End If
If Image111.Height = 750 And Image111.Width = 1095 Then
Image111.ToolTipText = "War Factory"
End If
If Image74.Height = 750 And Image74.Width = 1095 Then
Image74.ToolTipText = "War Factory"
End If
If Image80.Height = 750 And Image80.Width = 1095 Then
Image80.ToolTipText = "War Factory"
End If
If Image136.Height = 750 And Image136.Width = 1095 Then
Image136.ToolTipText = "War Factory"
End If
If Image144.Height = 750 And Image144.Width = 1095 Then
Image144.ToolTipText = "War Factory"
End If
If Image119.Height = 750 And Image119.Width = 1095 Then
Image119.ToolTipText = "War Factory"
End If
If Image147.Height = 750 And Image147.Width = 1095 Then
Image147.ToolTipText = "War Factory"
End If
If Image128.Height = 750 And Image128.Width = 1095 Then
Image128.ToolTipText = "War Factory"
End If
If Image117.Height = 750 And Image117.Width = 1095 Then
Image117.ToolTipText = "War Factory"
End If
If Image143.Height = 750 And Image143.Width = 1095 Then
Image143.ToolTipText = "War Factory"
End If
If Image55.Height = 750 And Image55.Width = 1095 Then
Image55.ToolTipText = "War Factory"
End If
If Image108.Height = 750 And Image108.Width = 1095 Then
Image108.ToolTipText = "War Factory"
End If
If Image89.Height = 750 And Image89.Width = 1095 Then
Image89.ToolTipText = "War Factory"
End If
If Image91.Height = 750 And Image91.Width = 1095 Then
Image91.ToolTipText = "War Factory"
End If
If Image97.Height = 750 And Image97.Width = 1095 Then
Image97.ToolTipText = "War Factory"
End If
If Image31.Height = 750 And Image31.Width = 1095 Then
Image31.ToolTipText = "War Factory"
End If
If Image70.Height = 750 And Image70.Width = 1095 Then
Image70.ToolTipText = "War Factory"
End If
If Image75.Height = 750 And Image75.Width = 1095 Then
Image75.ToolTipText = "War Factory"
End If
If Image64.Height = 750 And Image64.Width = 1095 Then
Image64.ToolTipText = "War Factory"
End If
If Image104.Height = 750 And Image104.Width = 1095 Then
Image104.ToolTipText = "War Factory"
End If
If Image99.Height = 750 And Image99.Width = 1095 Then
Image99.ToolTipText = "War Factory"
End If
If Image112.Height = 750 And Image112.Width = 1095 Then
Image112.ToolTipText = "War Factory"
End If
If Image92.Height = 750 And Image92.Width = 1095 Then
Image92.ToolTipText = "War Factory"
End If
If Image127.Height = 750 And Image127.Width = 1095 Then
Image127.ToolTipText = "War Factory"
End If
If Image95.Height = 750 And Image95.Width = 1095 Then
Image95.ToolTipText = "War Factory"
End If
If Image124.Height = 750 And Image124.Width = 1095 Then
Image124.ToolTipText = "War Factory"
End If
If Image123.Height = 750 And Image123.Width = 1095 Then
Image123.ToolTipText = "War Factory"
End If
If Image125.Height = 750 And Image125.Width = 1095 Then
Image125.ToolTipText = "War Factory"
End If
If Image48.Height = 750 And Image48.Width = 1095 Then
Image48.ToolTipText = "War Factory"
End If
If Image52.Height = 750 And Image52.Width = 1095 Then
Image52.ToolTipText = "War Factory"
End If
If Image52.Height = 750 And Image52.Width = 1095 Then
Image52.ToolTipText = "War Factory"
End If
If Image47.Height = 750 And Image47.Width = 1095 Then
Image47.ToolTipText = "War Factory"
End If
If Image49.Height = 750 And Image49.Width = 1095 Then
Image49.ToolTipText = "War Factory"
End If
If Image50.Height = 750 And Image50.Width = 1095 Then
Image50.ToolTipText = "War Factory"
End If
If Image53.Height = 750 And Image53.Width = 1095 Then
Image53.ToolTipText = "War Factory"
End If
If Image132.Height = 750 And Image132.Width = 1095 Then
Image132.ToolTipText = "War Factory"
End If
If Image140.Height = 750 And Image140.Width = 1095 Then
Image140.ToolTipText = "War Factory"
End If
If Image59.Height = 750 And Image59.Width = 1095 Then
Image59.ToolTipText = "War Factory"
End If
If Image65.Height = 750 And Image65.Width = 1095 Then
Image65.ToolTipText = "War Factory"
End If
If Image67.Height = 750 And Image67.Width = 1095 Then
Image67.ToolTipText = "War Factory"
End If
If Image109.Height = 750 And Image109.Width = 1095 Then
Image109.ToolTipText = "War Factory"
End If
If Image73.Height = 750 And Image73.Width = 1095 Then
Image73.ToolTipText = "War Factory"
End If
If Image120.Height = 750 And Image120.Width = 1095 Then
Image120.ToolTipText = "War Factory"
End If
If Image58.Height = 750 And Image58.Width = 1095 Then
Image58.ToolTipText = "War Factory"
End If
If Image51.Height = 750 And Image51.Width = 1095 Then
Image51.ToolTipText = "War Factory"
End If
If Image45.Height = 750 And Image45.Width = 1095 Then
Image45.ToolTipText = "War Factory"
End If
If Image71.Height = 750 And Image71.Width = 1095 Then
Image71.ToolTipText = "War Factory"
End If
If Image68.Height = 750 And Image68.Width = 1095 Then
Image68.ToolTipText = "War Factory"
End If
If Image105.Height = 750 And Image105.Width = 1095 Then
Image105.ToolTipText = "War Factory"
End If
If Image106.Height = 750 And Image106.Width = 1095 Then
Image106.ToolTipText = "War Factory"
End If
If Image115.Height = 750 And Image115.Width = 1095 Then
Image115.ToolTipText = "War Factory"
End If
If Image153.Height = 750 And Image153.Width = 1095 Then
Image153.ToolTipText = "War Factory"
End If
If Image139.Height = 750 And Image139.Width = 1095 Then
Image139.ToolTipText = "War Factory"
End If
If Image116.Height = 750 And Image116.Width = 1095 Then
Image116.ToolTipText = "War Factory"
End If
If Image145.Height = 750 And Image145.Width = 1095 Then
Image145.ToolTipText = "War Factory"
End If
If Image142.Height = 750 And Image142.Width = 1095 Then
Image142.ToolTipText = "War Factory"
End If
If Image146.Height = 750 And Image146.Width = 1095 Then
Image146.ToolTipText = "War Factory"
End If
If Image130.Height = 750 And Image130.Width = 1095 Then
Image130.ToolTipText = "War Factory"
End If
If Image121.Height = 750 And Image121.Width = 1095 Then
Image121.ToolTipText = "War Factory"
End If
If Image46.Height = 750 And Image46.Width = 1095 Then
Image46.ToolTipText = "War Factory"
End If
If Image44.Height = 750 And Image44.Width = 1095 Then
Image44.ToolTipText = "War Factory"
End If
If Image63.Height = 750 And Image63.Width = 1095 Then
Image63.ToolTipText = "War Factory"
End If
If Image60.Height = 750 And Image60.Width = 1095 Then
Image60.ToolTipText = "War Factory"
End If
If Image61.Height = 750 And Image61.Width = 1095 Then
Image61.ToolTipText = "War Factory"
End If
If Image62.Height = 750 And Image62.Width = 1095 Then
Image115.ToolTipText = "War Factory"
End If
If Image131.Height = 750 And Image131.Width = 1095 Then
Image131.ToolTipText = "War Factory"
End If
If Image148.Height = 750 And Image148.Width = 1095 Then
Image148.ToolTipText = "War Factory"
End If
If Image154.Height = 750 And Image154.Width = 1095 Then
Image154.ToolTipText = "War Factory"
End If
If Image141.Height = 750 And Image141.Width = 1095 Then
Image141.ToolTipText = "War Factory"
End If
If Image113.Height = 750 And Image113.Width = 1095 Then
Image113.ToolTipText = "War Factory"
End If
If Image152.Height = 750 And Image152.Width = 1095 Then
Image152.ToolTipText = "War Factory"
End If
If Image150.Height = 750 And Image150.Width = 1095 Then
Image150.ToolTipText = "War Factory"
End If
If Image102.Height = 750 And Image102.Width = 1095 Then
Image102.ToolTipText = "War Factory"
End If
If Image114.Height = 750 And Image114.Width = 1095 Then
Image114.ToolTipText = "War Factory"
End If
If Image103.Height = 750 And Image103.Width = 1095 Then
Image103.ToolTipText = "War Factory"
End If
If Image79.Height = 750 And Image79.Width = 1095 Then
Image79.ToolTipText = "War Factory"
End If
If Image141.Height = 750 And Image141.Width = 1095 Then
Image141.ToolTipText = "War Factory"
End If

End Sub

Private Sub Timer18_Timer()
'If Image80.Height = 315 And Image80.Width = 600 Or Image80.Height = 390 And Image80.Width = 345 Or Image80.Height = 375 And Image80.Width = 375 Then
'Call Timer18(1)
'End If
'Case 1
'Static blnface As Boolean
'If blnface = True Then
'Image80.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image80.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If
'End If
If Image79.Height = 315 And Image79.Width = 600 Or Image79.Height = 390 And Image79.Width = 345 Or Image79.Height = 375 And Image79.Width = 375 Then

If blnface = True Then
Image79.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image79.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
If Image73.Height = 315 And Image73.Width = 600 Or Image73.Height = 390 And Image73.Width = 345 Or Image73.Height = 375 And Image73.Width = 375 Then

If blnface = True Then
Image73.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image73.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
'If Image63.Height = 315 And Image63.Width = 600 Or Image63.Height = 390 And Image63.Width = 345 Or Image63.Height = 375 And Image63.Width = 375 Then

'If blnface = True Then
'mage63.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image63.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If
'End If
'If Image61.Height = 315 And Image61.Width = 600 Or Image61.Height = 390 And Image61.Width = 345 Or Image61.Height = 375 And Image61.Width = 375 Then

'If blnface = True Then
'Image61.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image61.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If
'End If

If Image59.Height = 315 And Image59.Width = 600 Or Image59.Height = 390 And Image59.Width = 345 Or Image59.Height = 375 And Image59.Width = 375 Then

If blnface = True Then
Image59.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image59.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
If Image58.Height = 315 And Image58.Width = 600 Or Image58.Height = 390 And Image58.Width = 345 Or Image58.Height = 375 And Image58.Width = 375 Then

If blnface = True Then
Image58.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image58.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
'If Image60.Height = 315 And Image60.Width = 600 Or Image60.Height = 390 And Image60.Width = 345 Or Image60.Height = 375 And Image60.Width = 375 Then

'If blnface = True Then
'Image60.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image60.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If
'End If
'If Image62.Height = 315 And Image62.Width = 600 Or Image62.Height = 390 And Image62.Width = 345 Or Image62.Height = 375 And Image62.Width = 375 Then

'If blnface = True Then
'Image62.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image62.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If
'End If
If Image65.Height = 315 And Image65.Width = 600 Or Image65.Height = 390 And Image65.Width = 345 Or Image65.Height = 375 And Image65.Width = 375 Then

If blnface = True Then
Image65.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image65.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
If Image51.Height = 315 And Image51.Width = 600 Or Image51.Height = 390 And Image51.Width = 345 Or Image51.Height = 375 And Image51.Width = 375 Then

If blnface = True Then
Image51.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image51.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
'If Image52.Height = 315 And Image52.Width = 600 Or Image52.Height = 390 And Image52.Width = 345 Or Image52.Height = 375 And Image52.Width = 375 Then

'If blnface = True Then
'Image52.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image52.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If
'End If
'If Image53.Height = 315 And Image53.Width = 600 Or Image53.Height = 390 And Image53.Width = 345 Or Image53.Height = 375 And Image53.Width = 375 Then

'I'f blnface = True Then
'Image53.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image53.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If
'End If
If Image84.Height = 315 And Image84.Width = 600 Or Image84.Height = 390 And Image84.Width = 345 Or Image84.Height = 375 And Image84.Width = 375 Then

If blnface = True Then
Image84.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image84.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
If Image64.Height = 315 And Image64.Width = 600 Or Image64.Height = 390 And Image64.Width = 345 Or Image64.Height = 375 And Image64.Width = 375 Then

If blnface = True Then
Image64.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
blnface = False
Else
Image64.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
blnface = True
End If
End If
End Sub

Private Sub Timer19_Timer()
Picture2.OLEDragMode = 0
Picture1.OLEDragMode = 0
Picture26.OLEDragMode = 0
Picture28.OLEDragMode = 0
Picture7.OLEDragMode = 0
Picture5.OLEDragMode = 0
Picture3.OLEDragMode = 0
Picture6.OLEDragMode = 0
Picture4.OLEDragMode = 0
Picture9.OLEDragMode = 0
Picture35.OLEDragMode = 0
Picture17.OLEDragMode = 0
Picture19.OLEDragMode = 0
Picture36.OLEDragMode = 0

Timer19.Enabled = False
End Sub

Private Sub Timer2_Timer()
Static blnface As Boolean
If blnface = True Then
Image11.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\rf3333.bmp")
blnface = False
Else
Image11.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\f.bmp")
blnface = True
End If
'S 'tatic blnface As Boolean
'If blnface = True Then
'Image8.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
'blnface = False
'Else
'Image8.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
'blnface = True
'End If




End Sub



Public Sub Power1()
If Power <= 0 Then

i = sndPlaySound(ByVal CStr(App.Path & "\lopower1.wav"), SND_ASYNC)
End If

End Sub
Private Sub image12_KeyDown(KeyCode As Integer, _
Shift As Integer)
  If KeyCode = vbKeyA Then
    Image12.Visible = False
    'MsgBox "You pressed _ the A key."
'i = sndPlaySound(ByVal CStr(App.Path & "\Taunt009.wav"), SND_ASYNC)
End If
End Sub
'Joathan Valentin 1999
'Private Sub Image1_MouseDown(Button As Integer, _
       ' Shift As Integer, X As Single, Y As Single)
   ' Image1.Drag vbBeginDrag
   ' Set Image1.DragIcon = LoadPicture("c:\Program _
      '  files\ Microsoft Visual _
      '  Basic\Icons\Dragdrop\Dragfldr.ico")
'End Sub
    
'Adding a DragOver event procedure to Image2 allows you to terminate dragging when the source enters the target. This example closes the file cabinet when Image1 is passed over Image2.
'Private Sub Image2_DragOver(Source As Control, _
  '      X As Single, Y As Single, State As Integer)
  '  Source.Drag vbEndDrag
  '  Source.Visible = False
 '   Image2.Picture = LoadPicture("c:\Program _
'        files\Microsoft Visual _
        Basic\Icons\Office\Files03a.ico")
'End Sub
    
'A'dding a third Image control to the form demonstrates canceling a drag operation. In this example the Image3 Picture property contains an icon of a trash can. Using the DragOver event and the source argument, dragging the files over Image3 cancels the drag operation.
'Private Sub Image3_DragOver(Source As Control, _
    '   ' X As Single, Y As Single, State As Integer)
    'Source.Drag vbCancel
'End Sub

'jonathan valentin 1999
Private Sub Timer3_Timer()
Image6.Left = Image6.Left + 90
End Sub

Private Sub Timer4_Timer()
''Static blnface As Boolean
''If blnface = True Then
''Image8.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam.bmp")
''blnface = False
''Else
''Image8.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\sam2.bmp")
''blnface = True
''End If
'End If

If Label2.Caption = "Year1" Then
If Label7.Caption >= 100 Then
i = sndPlaySound(ByVal CStr(App.Path & "\5169.wav"), SND_ASYNC)
Timer5.Enabled = False
Label7.Caption = "0"

If Label5.Caption = "1c" Then
Picture1.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\town2.bmp")
Picture2.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\h2.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g2.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p2.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov2.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t2.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\mar2.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b2.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s2.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\a2.bmp")
Label2.Caption = "year2"
End If
End If
End If
If Label2.Caption = "year2" Then
If Label7.Caption >= 100 Then
i = sndPlaySound(ByVal CStr(App.Path & "\5169.wav"), SND_ASYNC)
Timer5.Enabled = False
Label7.Caption = "0"

If Label5.Caption = "1c" Then
Picture1.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\town3.bmp")
Picture2.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\h3.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g2.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p3.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov2.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t3.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\mar 4.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b3.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s2.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\a2.bmp")
Label2.Caption = "year3"
End If
End If
End If
If Label2.Caption = "year3" Then
If Label7.Caption >= 100 Then
i = sndPlaySound(ByVal CStr(App.Path & "\5169.wav"), SND_ASYNC)
Timer5.Enabled = False
Label7.Caption = "0"

If Label5.Caption = "1c" Then
Picture1.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\town4.bmp")
Picture2.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\h4.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g2.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p3.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov2.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t3.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\mar 4.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b3.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s2.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\a2.bmp")
Label2.Caption = "year4"
End If
End If
End If
If Label2.Caption = "year4" Then
If Label7.Caption >= 100 Then
i = sndPlaySound(ByVal CStr(App.Path & "\5169.wav"), SND_ASYNC)
Timer5.Enabled = False
Label7.Caption = "0"

If Label5.Caption = "1c" Then
Picture1.Picture = LoadPicture("C:\windows\desktop\vb projects\war\seven\fegrt.bmp")
Picture2.Picture = LoadPicture("C:\windows\desktop\vb projects\war\buildings\h4.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\g2.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\p3.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\gov2.bmp")
Picture9.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\t3.bmp")
Picture17.Picture = LoadPicture("c:\windows\desktop\vb projects\war\seven\mark4et1.bmp")
Picture26.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\b3.bmp")
Picture28.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\s2.bmp")
Picture35.Picture = LoadPicture("c:\windows\desktop\vb projects\war\buildings\a2.bmp")
Label2.Caption = "year5"
End If
End If
End If
End Sub

Private Sub Timer5_Timer()
Label7.Caption = Label7.Caption + 1
If Label7.Caption >= 100 Then
i = sndPlaySound(ByVal CStr(App.Path & "\5169.wav"), SND_ASYNC)
Timer5.Enabled = False
Label7.Caption = "0"
End If

'If year.Caption <= 100 Then
'Image26.Visible = True
'  'If year.Caption >= 100 Then
  'Image26.Visible = False
 ' End If
''If year.Caption >= 1910 Then
''Image15.Visible = True
'Image9.Visible = False
'Image17.Visible = False
'Image18.Visible = True
'Image18.Enabled = False
''Image18.Visible = False
'Image9.Visible = False
''..End If
''If year.Caption >= 2050 Then
''Image14.Visible = True
''Image2.Visible = False
''Image16.Visible = True
''Image15.Visible = False
'Image9.Visible = False
'Image17.Visible = False
'i ' 'mage18.Visible = False
''End If

End Sub

Private Sub Timer6_Timer()
Label1.Visible = False
''If year.Caption >= 1910 Then
''Image15.Visible = True
'Image9.Visible = False
'Image17.Visible = False
'Image18.Visible = True
'Image18.Enabled = False
''Image18.Visible = False
''Image18.Enabled = False
'Image9.Visible = False
''End If
End Sub

'jonathan valentin aug6,1999
Private Sub Timer7_Timer()
'Private Sub Timer2_Timer()
'Static blnface As Boolean
'If blnface = True Then
'Image24.Picture = LoadPicture("C:\windows\desktop\vb projects\war\b1.bmp")
'blnface = False
'Else
'Image24.Picture = LoadPicture("C:\windows\desktop\vb projects\war\b3.bmp")
'blnface = True
'End If

End Sub

Private Sub Timer8_Timer()
Label3.Visible = False
'Label1.Visible = False
End Sub

Private Sub Timer9_Timer()
Label1.Visible = False
Timer9.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim msgpress As Integer
Select Case Button.Key
Case Is = "a":
i = sndPlaySound(ByVal CStr(App.Path & "\acue12.wav"), SND_ASYNC)
 'dlgCommonDialog.Save
 
Case Is = "b"
i = sndPlaySound(ByVal CStr(App.Path & "\acue11.wav"), SND_ASYNC)
'form2.Show vbModal, Me

Case Is = "c":
i = sndPlaySound(ByVal CStr(App.Path & "\acue2.wav"), SND_ASYNC)
'dlgCommonDialog.ShowOpen
Case Is = "d":
i = sndPlaySound(ByVal CStr(App.Path & "\acue4b.wav"), SND_ASYNC)
'Form4.Show vbModal, Me
Case Is = "e":
i = sndPlaySound(ByVal CStr(App.Path & "\acue18.wav"), SND_ASYNC)
'dlgCommonDialog.ShowPrinter
'Case Is = "EARSE":
Case Is = "f":
MsgBox "need help check read me"
Case Is = "q1":

Case Is = "q2":

Case Is = "q3":

Case Is = "q4":

'Case Is = "STOP":
'Unload Me
'End
'Case Is = "HELP":
'frmAbout.Show vbModal, Me
End Select
End Sub
'End Sub
Private Sub VScroll1_Change()
'Picture37.Top = -VScroll1.Value
Image33.Top = -VScroll1.Value
End Sub
Public Sub pgain()
   response = MsgBox("Are You Sure You Want To Quit?", vbYesNo, "Quit?")
   If response = vbYes Then
   End
    'Call totop
    'score = 0
    'lives = 5
    'life1.Visible = True
    'life2.Visible = True
    'l 'ife3.Visible = True
    'l'ife4.Visible = True
    'life5.Visible = True
    'Timer1.Interval = 125
   Else
    Image103.Visible = True
    End If
End Sub
Public Sub restart()
response = MsgBox("Are You Sure You Want To Restart?", vbYesNo, "Restart?")
   If response = vbYes Then
  
    If Label5.Caption = "1c" Then
Call Map
Timer15.Enabled = False
'Image16.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t2.bmp")
'Image150.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t1.bmp")
'Image48.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t2.bmp")
'Image91.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t1.bmp")
'Image16.OLEDragMode = 0
'Image150.OLEDragMode = 0
'Image48.OLEDragMode = 0
'Image91.OLEDragMode = 0
End If
If Label5.Caption = "1a" Then

Image136.Picture = LoadPicture("c:\windows\desktop\vb projects\war\tr1.bmp")
Image51.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t.bmp")
Image46.Picture = LoadPicture("c:\windows\desktop\vb projects\war\t.bmp")
Image72.Picture = LoadPicture("c:\windows\desktop\vb projects\war\tr1.bmp")
Image136.OLEDragMode = 0
Image51.OLEDragMode = 0
Image46.OLEDragMode = 0
Image72.OLEDragMode = 0
Timer15.Enabled = False
End If
    'Call totop
    'score = 0
    'lives = 5
    'life1.Visible = True
    'life2.Visible = True
    'l 'ife3.Visible = True
    'l'ife4.Visible = True
    'life5.Visible = True
    'Timer1.Interval = 125
   Else
   Unload Me
    Dim Form3 As New Form3
    Form3.Show vbModal, Me
    End If
End Sub


Private Sub vscmap_Change()
    Module1.Map.CurrentY = Me.vscmap.Value
    ' Redraw the map
   ' Call DrawMap
End Sub
Public Sub silo()

If Image43.Height = 420 And Image43.Width = 435 Then
Image43.ToolTipText = "Silo"
If Image18.Height = 495 And Image18.Width = 975 Then
Image18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status2.bmp")
Image18.Visible = False
End If

End If
If Image42.Height = 420 And Image42.Width = 435 Then
Image42.ToolTipText = "Silo"
If Image43.Height = 495 And Image43.Width = 975 Then
Image43.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status2.bmp")
Image43.Visible = False
End If
End If

If Image18.Height = 420 And Image18.Width = 435 Then
Image18.ToolTipText = "silo"
If Image42.Height = 495 And Image42.Width = 975 Then
Image42.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status2.bmp")
Image42.Visible = False
End If
End If
If Image19.Height = 420 And Image19.Width = 435 Then
Image19.ToolTipText = "Silo"
End If
If Image22.Height = 420 And Image22.Width = 435 Then
Image22.ToolTipText = "Silo"
End If
If Image10.Height = 420 And Image10.Width = 435 Then
Image10.ToolTipText = "Silo"
End If
If Image11.Height = 420 And Image11.Width = 435 Then
Image11.ToolTipText = "Silo"
End If
If Image21.Height = 420 And Image21.Width = 435 Then
Image21.ToolTipText = "Silo"
End If
If Image26.Height = 420 And Image26.Width = 435 Then
Image26.ToolTipText = "Silo"
End If
If Image23.Height = 420 And Image23.Width = 435 Then
Image23.ToolTipText = "Silo"
End If
If Image20.Height = 420 And Image20.Width = 435 Then
Image20.ToolTipText = "Silo"
End If
If Image24.Height = 420 And Image24.Width = 435 Then
Image24.ToolTipText = "Silo"
End If
If Image2.Height = 420 And Image2.Width = 435 Then
Image2.ToolTipText = "Silo"
End If
If Image3.Height = 420 And Image3.Width = 435 Then
Image3.ToolTipText = "Silo"
End If
If Image5.Height = 420 And Image5.Width = 435 Then
Image5.ToolTipText = "Silo"
End If
If Image4.Height = 420 And Image4.Width = 435 Then
Image4.ToolTipText = "Silo"
End If
If Image6.Height = 420 And Image6.Width = 435 Then
Image6.ToolTipText = "Silo"
End If
If Image1.Height = 420 And Image1.Width = 435 Then
Image1.ToolTipText = "Silo"
End If
If Image17.Height = 420 And Image17.Width = 435 Then
Image17.ToolTipText = "Silo"
End If
If Image12.Height = 420 And Image12.Width = 435 Then
Image12.ToolTipText = "Silo"
End If
If Image8.Height = 420 And Image8.Width = 435 Then
Image8.ToolTipText = "Silo"
End If
If Image7.Height = 420 And Image7.Width = 435 Then
Image7.ToolTipText = "Silo"
End If
If Image15.Height = 420 And Image15.Width = 435 Then
Image15.ToolTipText = "Silo"
End If
If Image14.Height = 420 And Image14.Width = 435 Then
Image14.ToolTipText = "Silo"
End If
If Image9.Height = 420 And Image9.Width = 435 Then
Image9.ToolTipText = "Silo"
End If
If Image16.Height = 420 And Image16.Width = 435 Then
Image16.ToolTipText = "Silo"
End If
If Image13.Height = 420 And Image13.Width = 435 Then
Image13.ToolTipText = "Silo"
End If
If Image28.Height = 420 And Image28.Width = 435 Then
Image28.ToolTipText = "Silo"
End If
If Image29.Height = 420 And Image29.Width = 435 Then
Image29.ToolTipText = "Silo"
End If
If Image32.Height = 420 And Image32.Width = 435 Then
Image32.ToolTipText = "Silo"
End If
If Image133.Height = 420 And Image133.Width = 435 Then
Image133.ToolTipText = "Silo"
End If
If Image86.Height = 420 And Image86.Width = 435 Then
Image86.ToolTipText = "Silo"
End If
If Image126.Height = 420 And Image126.Width = 435 Then
Image126.ToolTipText = "Silo"
End If
If Image151.Height = 420 And Image151.Width = 435 Then
Image151.ToolTipText = "Silo"
End If
If Image69.Height = 420 And Image69.Width = 435 Then
Image69.ToolTipText = "Silo"
End If
If Image96.Height = 420 And Image96.Width = 435 Then
Image96.ToolTipText = "Silo"
End If
If Image78.Height = 420 And Image78.Width = 435 Then
Image78.ToolTipText = "Silo"
End If
If Image88.Height = 420 And Image88.Width = 435 Then
Image88.ToolTipText = "Silo"
End If
If Image66.Height = 420 And Image66.Width = 435 Then
Image66.ToolTipText = "Silo"
End If
If Image118.Height = 420 And Image118.Width = 435 Then
Image118.ToolTipText = "Silo"
End If
If Image72.Height = 420 And Image72.Width = 435 Then
Image72.ToolTipText = "Silo"
End If
If Image110.Height = 420 And Image110.Width = 435 Then
Image110.ToolTipText = "Silo"
End If
If Image135.Height = 420 And Image135.Width = 435 Then
Image135.ToolTipText = "Silo"
End If
If Image57.Height = 420 And Image57.Width = 435 Then
Image57.ToolTipText = "Silo"
End If
If Image77.Height = 420 And Image77.Width = 435 Then
Image77.ToolTipText = "Silo"
End If
If Image85.Height = 420 And Image85.Width = 435 Then
Image85.ToolTipText = "Silo"
End If
If Image83.Height = 420 And Image83.Width = 435 Then
Image83.ToolTipText = "Silo"
End If
If Image56.Height = 420 And Image56.Width = 435 Then
Image56.ToolTipText = "Silo"
End If
If Image93.Height = 420 And Image93.Width = 435 Then
Image93.ToolTipText = "Silo"
End If
If Image107.Height = 420 And Image107.Width = 435 Then
Image107.ToolTipText = "Silo"
End If
If Image94.Height = 420 And Image94.Width = 435 Then
Image94.ToolTipText = "Silo"
End If
If Image30.Height = 420 And Image30.Width = 435 Then
Image30.ToolTipText = "Silo"
End If
If Image100.Height = 420 And Image100.Width = 435 Then
Image100.ToolTipText = "Silo"
End If
If Image101.Height = 420 And Image101.Width = 435 Then
Image101.ToolTipText = "Silo"
End If
If Image84.Height = 420 And Image84.Width = 435 Then
Image84.ToolTipText = "Silo"
End If
If Image54.Height = 420 And Image54.Width = 435 Then
Image54.ToolTipText = "Silo"
End If
If Image129.Height = 420 And Image129.Width = 435 Then
Image129.ToolTipText = "Silo"
End If
If Image76.Height = 420 And Image76.Width = 435 Then
Image76.ToolTipText = "Silo"
End If
If Image134.Height = 420 And Image134.Width = 435 Then
Image134.ToolTipText = "Silo"
End If
If Image111.Height = 420 And Image111.Width = 435 Then
Image111.ToolTipText = "Silo"
End If
If Image74.Height = 420 And Image74.Width = 435 Then
Image74.ToolTipText = "Silo"
End If
If Image80.Height = 420 And Image80.Width = 435 Then
Image80.ToolTipText = "Silo"
End If
If Image136.Height = 420 And Image136.Width = 435 Then
Image136.ToolTipText = "Silo"
End If
If Image144.Height = 420 And Image144.Width = 435 Then
Image144.ToolTipText = "Silo"
End If
If Image119.Height = 420 And Image119.Width = 435 Then
Image119.ToolTipText = "Silo"
End If
If Image147.Height = 420 And Image147.Width = 435 Then
Image147.ToolTipText = "Silo"
End If
If Image128.Height = 420 And Image128.Width = 435 Then
Image128.ToolTipText = "Silo"
End If
If Image117.Height = 420 And Image117.Width = 435 Then
Image117.ToolTipText = "Silo"
End If
If Image143.Height = 420 And Image143.Width = 435 Then
Image143.ToolTipText = "Silo"
End If
If Image55.Height = 420 And Image55.Width = 435 Then
Image55.ToolTipText = "Silo"
End If
If Image108.Height = 420 And Image108.Width = 435 Then
Image108.ToolTipText = "Silo"
End If
If Image89.Height = 420 And Image89.Width = 435 Then
Image89.ToolTipText = "Silo"
End If
If Image91.Height = 420 And Image91.Width = 435 Then
Image91.ToolTipText = "Silo"
End If
If Image97.Height = 420 And Image97.Width = 435 Then
Image97.ToolTipText = "Silo"
End If
If Image31.Height = 420 And Image31.Width = 435 Then
Image31.ToolTipText = "Silo"
End If

End Sub
Public Sub tech()
If Image43.Height = 720 And Image43.Width = 1170 Then
Image43.ToolTipText = "Tech"
If Image18.Height = 495 And Image18.Width = 975 Then
Image18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image18.Visible = False
End If

End If
If Image42.Height = 720 And Image42.Width = 1170 Then
Image42.ToolTipText = "Tech"
If Image43.Height = 495 And Image43.Width = 975 Then
Image43.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image43.Visible = False
End If
End If

If Image18.Height = 720 And Image18.Width = 1170 Then
Image18.ToolTipText = "Tech"
If Image42.Height = 495 And Image42.Width = 975 Then
Image42.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image42.Visible = False
End If
End If
If Image19.Height = 720 And Image19.Width = 1170 Then
Image19.ToolTipText = "Tech"
End If
If Image22.Height = 720 And Image22.Width = 1170 Then
Image22.ToolTipText = "Tech"
End If
If Image10.Height = 720 And Image10.Width = 1170 Then
Image10.ToolTipText = "Tech"
End If
If Image11.Height = 720 And Image11.Width = 1170 Then
Image11.ToolTipText = "Tech"
End If
If Image21.Height = 720 And Image21.Width = 1170 Then
Image21.ToolTipText = "Tech"
End If
If Image26.Height = 720 And Image26.Width = 1170 Then
Image26.ToolTipText = "Tech"
End If
If Image23.Height = 720 And Image23.Width = 1170 Then
Image23.ToolTipText = "Tech"
End If
If Image20.Height = 720 And Image20.Width = 1170 Then
Image20.ToolTipText = "Tech"
End If
If Image24.Height = 720 And Image24.Width = 1170 Then
Image24.ToolTipText = "Tech"
End If
If Image2.Height = 720 And Image2.Width = 1170 Then
Image2.ToolTipText = "Tech"
End If
If Image3.Height = 720 And Image3.Width = 1170 Then
Image3.ToolTipText = "Tech"
End If
If Image5.Height = 720 And Image5.Width = 1170 Then
Image5.ToolTipText = "Tech"
End If
If Image4.Height = 720 And Image4.Width = 1170 Then
Image4.ToolTipText = "Tech"
End If
If Image6.Height = 720 And Image6.Width = 1170 Then
Image6.ToolTipText = "Tech"
End If
If Image1.Height = 720 And Image1.Width = 1170 Then
Image1.ToolTipText = "Tech"
End If
If Image17.Height = 720 And Image17.Width = 1170 Then
Image17.ToolTipText = "Tech"
End If
If Image12.Height = 720 And Image12.Width = 1170 Then
Image12.ToolTipText = "Tech"
End If
If Image8.Height = 720 And Image8.Width = 1170 Then
Image8.ToolTipText = "Tech"
End If
If Image7.Height = 720 And Image7.Width = 1170 Then
Image7.ToolTipText = "Tech"
End If
If Image15.Height = 720 And Image15.Width = 1170 Then
Image15.ToolTipText = "Tech"
End If
If Image14.Height = 720 And Image14.Width = 1170 Then
Image14.ToolTipText = "Tech"
End If
If Image9.Height = 720 And Image9.Width = 1170 Then
Image9.ToolTipText = "Tech"
End If
If Image16.Height = 720 And Image16.Width = 1170 Then
Image16.ToolTipText = "Tech"
End If
If Image13.Height = 720 And Image13.Width = 1170 Then
Image13.ToolTipText = "Tech"
End If
If Image28.Height = 720 And Image28.Width = 1170 Then
Image28.ToolTipText = "Tech"
End If
If Image29.Height = 720 And Image29.Width = 1170 Then
Image29.ToolTipText = "Tech"
End If
If Image32.Height = 720 And Image32.Width = 1170 Then
Image32.ToolTipText = "Tech"
End If
If Image133.Height = 720 And Image133.Width = 1170 Then
Image133.ToolTipText = "Tech"
End If
If Image86.Height = 720 And Image86.Width = 1170 Then
Image86.ToolTipText = "Tech"
End If
If Image126.Height = 720 And Image126.Width = 1170 Then
Image126.ToolTipText = "Tech"
End If
If Image151.Height = 720 And Image151.Width = 1170 Then
Image151.ToolTipText = "Tech"
End If
If Image69.Height = 720 And Image69.Width = 1170 Then
Image69.ToolTipText = "Tech"
End If
If Image96.Height = 720 And Image96.Width = 1170 Then
Image96.ToolTipText = "Tech"
End If
If Image78.Height = 720 And Image78.Width = 1170 Then
Image78.ToolTipText = "Tech"
End If
If Image88.Height = 720 And Image88.Width = 1170 Then
Image88.ToolTipText = "Tech"
End If
If Image66.Height = 720 And Image66.Width = 1170 Then
Image66.ToolTipText = "Tech"
End If
If Image118.Height = 720 And Image118.Width = 1170 Then
Image118.ToolTipText = "Tech"
End If
If Image72.Height = 720 And Image72.Width = 1170 Then
Image72.ToolTipText = "Tech"
End If
If Image110.Height = 720 And Image110.Width = 1170 Then
Image110.ToolTipText = "Tech"
End If
If Image135.Height = 720 And Image135.Width = 1170 Then
Image135.ToolTipText = "Tech"
End If
If Image57.Height = 720 And Image57.Width = 1170 Then
Image57.ToolTipText = "Tech"
End If
If Image77.Height = 720 And Image77.Width = 1170 Then
Image77.ToolTipText = "Tech"
End If
If Image85.Height = 720 And Image85.Width = 1170 Then
Image85.ToolTipText = "Tech"
End If
If Image83.Height = 720 And Image83.Width = 1170 Then
Image83.ToolTipText = "Tech"
End If
If Image56.Height = 720 And Image56.Width = 1170 Then
Image56.ToolTipText = "Tech"
End If
If Image93.Height = 720 And Image93.Width = 1170 Then
Image93.ToolTipText = "Tech"
End If
If Image107.Height = 720 And Image107.Width = 1170 Then
Image107.ToolTipText = "Tech"
End If
If Image94.Height = 720 And Image94.Width = 1170 Then
Image94.ToolTipText = "Tech"
End If
If Image30.Height = 720 And Image30.Width = 1170 Then
Image30.ToolTipText = "Tech"
End If
If Image100.Height = 720 And Image100.Width = 1170 Then
Image100.ToolTipText = "Tech"
End If
If Image101.Height = 720 And Image101.Width = 1170 Then
Image101.ToolTipText = "Tech"
End If
If Image84.Height = 720 And Image84.Width = 1170 Then
Image84.ToolTipText = "Tech"
End If
If Image54.Height = 720 And Image54.Width = 1170 Then
Image54.ToolTipText = "Tech"
End If
If Image129.Height = 720 And Image129.Width = 1170 Then
Image129.ToolTipText = "Tech"
End If
If Image76.Height = 720 And Image76.Width = 1170 Then
Image76.ToolTipText = "Tech"
End If
If Image134.Height = 720 And Image134.Width = 1170 Then
Image134.ToolTipText = "Tech"
End If
If Image111.Height = 720 And Image111.Width = 1170 Then
Image111.ToolTipText = "Tech"
End If
If Image74.Height = 720 And Image74.Width = 1170 Then
Image74.ToolTipText = "Tech"
End If
If Image80.Height = 720 And Image80.Width = 1170 Then
Image80.ToolTipText = "Tech"
End If
If Image136.Height = 720 And Image136.Width = 1170 Then
Image136.ToolTipText = "Tech"
End If
If Image144.Height = 720 And Image144.Width = 1170 Then
Image144.ToolTipText = "Tech"
End If
If Image119.Height = 720 And Image119.Width = 1170 Then
Image119.ToolTipText = "Tech"
End If
If Image147.Height = 720 And Image147.Width = 1170 Then
Image147.ToolTipText = "Tech"
End If
If Image128.Height = 720 And Image128.Width = 1170 Then
Image128.ToolTipText = "Tech"
End If
If Image117.Height = 720 And Image117.Width = 1170 Then
Image117.ToolTipText = "Tech"
End If
If Image143.Height = 720 And Image143.Width = 1170 Then
Image143.ToolTipText = "Tech"
End If
If Image55.Height = 720 And Image55.Width = 1170 Then
Image55.ToolTipText = "Tech"
End If
If Image108.Height = 720 And Image108.Width = 1170 Then
Image108.ToolTipText = "Tech"
End If
If Image89.Height = 720 And Image89.Width = 1170 Then
Image89.ToolTipText = "Tech"
End If
If Image91.Height = 720 And Image91.Width = 1170 Then
Image91.ToolTipText = "Tech"
End If
If Image97.Height = 720 And Image97.Width = 1170 Then
Image97.ToolTipText = "Tech"
End If
If Image31.Height = 720 And Image31.Width = 1170 Then
Image31.ToolTipText = "Tech"
End If

End Sub
Public Sub com()

If Image43.Height = 780 And Image43.Width = 810 Then
Image43.ToolTipText = "Commications"
If Image18.Height = 495 And Image18.Width = 975 Then
Image18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image18.Visible = False
End If

End If
If Image42.Height = 780 And Image42.Width = 810 Then
Image42.ToolTipText = "Commications"
If Image43.Height = 495 And Image43.Width = 975 Then
Image43.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image43.Visible = False
End If
End If

If Image18.Height = 780 And Image18.Width = 810 Then
Image18.ToolTipText = "Commications"
If Image42.Height = 495 And Image42.Width = 975 Then
Image42.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image42.Visible = False
End If
End If
If Image19.Height = 780 And Image19.Width = 810 Then
Image19.ToolTipText = "Commications"
End If
If Image22.Height = 780 And Image22.Width = 810 Then
Image22.ToolTipText = "Commications"
End If
If Image10.Height = 780 And Image10.Width = 810 Then
Image10.ToolTipText = "Commications"
End If
If Image11.Height = 780 And Image11.Width = 810 Then
Image11.ToolTipText = "Commications"
End If
If Image21.Height = 780 And Image21.Width = 810 Then
Image21.ToolTipText = "Commications"
End If
If Image26.Height = 780 And Image26.Width = 810 Then
Image26.ToolTipText = "Commications"
End If
If Image23.Height = 780 And Image23.Width = 810 Then
Image23.ToolTipText = "Commications"
End If
If Image20.Height = 780 And Image20.Width = 810 Then
Image20.ToolTipText = "Commications"
End If
If Image24.Height = 780 And Image24.Width = 810 Then
Image24.ToolTipText = "Commications"
End If
If Image2.Height = 780 And Image2.Width = 810 Then
Image2.ToolTipText = "Commications"
End If
If Image3.Height = 780 And Image3.Width = 810 Then
Image3.ToolTipText = "Commications"
End If
If Image5.Height = 780 And Image5.Width = 810 Then
Image5.ToolTipText = "Commications"
End If
If Image4.Height = 780 And Image4.Width = 810 Then
Image4.ToolTipText = "Commications"
End If
If Image6.Height = 780 And Image6.Width = 810 Then
Image6.ToolTipText = "Commications"
End If
If Image1.Height = 780 And Image1.Width = 810 Then
Image1.ToolTipText = "Commications"
End If
If Image17.Height = 780 And Image17.Width = 810 Then
Image17.ToolTipText = "Commications"
End If
If Image12.Height = 780 And Image12.Width = 810 Then
Image12.ToolTipText = "Commications"
End If
If Image8.Height = 780 And Image8.Width = 810 Then
Image8.ToolTipText = "Commications"
End If
If Image7.Height = 780 And Image7.Width = 810 Then
Image7.ToolTipText = "Commications"
End If
If Image15.Height = 780 And Image15.Width = 810 Then
Image15.ToolTipText = "Commications"
End If
If Image14.Height = 780 And Image14.Width = 810 Then
Image14.ToolTipText = "Commications"
End If
If Image9.Height = 780 And Image9.Width = 810 Then
Image9.ToolTipText = "Commications"
End If
If Image16.Height = 780 And Image16.Width = 810 Then
Image16.ToolTipText = "Commications"
End If
If Image13.Height = 780 And Image13.Width = 810 Then
Image13.ToolTipText = "Commications"
End If
If Image28.Height = 780 And Image28.Width = 810 Then
Image28.ToolTipText = "Commications"
End If
If Image29.Height = 780 And Image29.Width = 810 Then
Image29.ToolTipText = "Commications"
End If
If Image32.Height = 780 And Image32.Width = 810 Then
Image32.ToolTipText = "Commications"
End If
If Image133.Height = 780 And Image133.Width = 810 Then
Image133.ToolTipText = "Commications"
End If
If Image86.Height = 780 And Image86.Width = 810 Then
Image86.ToolTipText = "Commications"
End If
If Image126.Height = 780 And Image126.Width = 810 Then
Image126.ToolTipText = "Commications"
End If
If Image151.Height = 780 And Image151.Width = 810 Then
Image151.ToolTipText = "Commications"
End If
If Image69.Height = 780 And Image69.Width = 810 Then
Image69.ToolTipText = "Commications"
End If
If Image96.Height = 780 And Image96.Width = 810 Then
Image96.ToolTipText = "Commications"
End If
If Image78.Height = 780 And Image78.Width = 810 Then
Image78.ToolTipText = "Commications"
End If
If Image88.Height = 780 And Image88.Width = 810 Then
Image88.ToolTipText = "Commications"
End If
If Image66.Height = 780 And Image66.Width = 810 Then
Image66.ToolTipText = "Commications"
End If
If Image118.Height = 780 And Image118.Width = 810 Then
Image118.ToolTipText = "Commications"
End If
If Image72.Height = 780 And Image72.Width = 810 Then
Image72.ToolTipText = "Commications"
End If
If Image110.Height = 780 And Image110.Width = 810 Then
Image110.ToolTipText = "Commications"
End If
If Image135.Height = 780 And Image135.Width = 810 Then
Image135.ToolTipText = "Commications"
End If
If Image57.Height = 780 And Image57.Width = 810 Then
Image57.ToolTipText = "Commications"
End If
If Image77.Height = 780 And Image77.Width = 810 Then
Image77.ToolTipText = "Commications"
End If
If Image85.Height = 780 And Image85.Width = 810 Then
Image85.ToolTipText = "Commications"
End If
If Image83.Height = 780 And Image83.Width = 810 Then
Image83.ToolTipText = "Commications"
End If
If Image56.Height = 780 And Image56.Width = 810 Then
Image56.ToolTipText = "Commications"
End If
If Image93.Height = 780 And Image93.Width = 810 Then
Image93.ToolTipText = "Commications"
End If
If Image107.Height = 780 And Image107.Width = 810 Then
Image107.ToolTipText = "Commications"
End If
If Image94.Height = 780 And Image94.Width = 810 Then
Image94.ToolTipText = "Commications"
End If
If Image30.Height = 780 And Image30.Width = 810 Then
Image30.ToolTipText = "Commications"
End If
If Image100.Height = 780 And Image100.Width = 810 Then
Image100.ToolTipText = "Commications"
End If
If Image101.Height = 780 And Image101.Width = 810 Then
Image101.ToolTipText = "Commications"
End If
If Image84.Height = 780 And Image84.Width = 810 Then
Image84.ToolTipText = "Commications"
End If
If Image54.Height = 780 And Image54.Width = 810 Then
Image54.ToolTipText = "Commications"
End If
If Image129.Height = 780 And Image129.Width = 810 Then
Image129.ToolTipText = "Commications"
End If
If Image76.Height = 780 And Image76.Width = 810 Then
Image76.ToolTipText = "Commications"
End If
If Image134.Height = 780 And Image134.Width = 810 Then
Image134.ToolTipText = "Commications"
End If
If Image111.Height = 780 And Image111.Width = 810 Then
Image111.ToolTipText = "Commications"
End If
If Image74.Height = 780 And Image74.Width = 810 Then
Image74.ToolTipText = "Commications"
End If
If Image80.Height = 780 And Image80.Width = 810 Then
Image80.ToolTipText = "Commications"
End If
If Image136.Height = 780 And Image136.Width = 810 Then
Image136.ToolTipText = "Commications"
End If
If Image144.Height = 780 And Image144.Width = 810 Then
Image144.ToolTipText = "Commications"
End If
If Image119.Height = 780 And Image119.Width = 810 Then
Image119.ToolTipText = "Commications"
End If
If Image147.Height = 780 And Image147.Width = 810 Then
Image147.ToolTipText = "Commications"
End If
If Image128.Height = 780 And Image128.Width = 810 Then
Image128.ToolTipText = "Commications"
End If
If Image117.Height = 780 And Image117.Width = 810 Then
Image117.ToolTipText = "Commications"
End If
If Image143.Height = 780 And Image143.Width = 810 Then
Image143.ToolTipText = "Commications"
End If
If Image55.Height = 780 And Image55.Width = 810 Then
Image55.ToolTipText = "Commications"
End If
If Image108.Height = 780 And Image108.Width = 810 Then
Image108.ToolTipText = "Commications"
End If
If Image89.Height = 780 And Image89.Width = 810 Then
Image89.ToolTipText = "Commications"
End If
If Image91.Height = 780 And Image91.Width = 810 Then
Image91.ToolTipText = "Commications"
End If
If Image97.Height = 780 And Image97.Width = 810 Then
Image97.ToolTipText = "Commications"
End If
If Image31.Height = 780 And Image31.Width = 810 Then
Image31.ToolTipText = "Commications"
End If

End Sub
Public Sub plant()
If Image43.Height = 1020 And Image43.Width = 1065 Then
Image43.ToolTipText = "Power Plant"
If Image18.Height = 495 And Image18.Width = 975 Then
Image18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image18.Visible = False
End If

End If
If Image42.Height = 1020 And Image42.Width = 1065 Then
Image42.ToolTipText = "Power Plant"
If Image43.Height = 495 And Image43.Width = 975 Then
Image43.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image43.Visible = False
End If
End If

If Image18.Height = 1020 And Image18.Width = 1065 Then
Image18.ToolTipText = "Power Plant"
If Image42.Height = 495 And Image42.Width = 975 Then
Image42.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image42.Visible = False
End If
End If
If Image19.Height = 1020 And Image19.Width = 1065 Then
Image19.ToolTipText = "Power Plant"
End If
If Image22.Height = 1020 And Image22.Width = 1065 Then
Image22.ToolTipText = "Power Plant"
End If
If Image10.Height = 1020 And Image10.Width = 1065 Then
Image10.ToolTipText = "Power Plant"
End If
If Image11.Height = 1020 And Image11.Width = 1065 Then
Image11.ToolTipText = "Power Plant"
End If
If Image21.Height = 1020 And Image21.Width = 1065 Then
Image21.ToolTipText = "Power Plant"
End If
If Image26.Height = 1020 And Image26.Width = 1065 Then
Image26.ToolTipText = "Power Plant"
End If
If Image23.Height = 1020 And Image23.Width = 1065 Then
Image23.ToolTipText = "Power Plant"
End If
If Image20.Height = 1020 And Image20.Width = 1065 Then
Image20.ToolTipText = "Power Plant"
End If
If Image24.Height = 1020 And Image24.Width = 1065 Then
Image24.ToolTipText = "Power Plant"
End If
If Image2.Height = 1020 And Image2.Width = 1065 Then
Image2.ToolTipText = "Power Plant"
End If
If Image3.Height = 1020 And Image3.Width = 1065 Then
Image3.ToolTipText = "Power Plant"
End If
If Image5.Height = 1020 And Image5.Width = 1065 Then
Image5.ToolTipText = "Power Plant"
End If
If Image4.Height = 1020 And Image4.Width = 1065 Then
Image4.ToolTipText = "Power Plant"
End If
If Image6.Height = 1020 And Image6.Width = 1065 Then
Image6.ToolTipText = "Power Plant"
End If
If Image1.Height = 1020 And Image1.Width = 1065 Then
Image1.ToolTipText = "Power Plant"
End If
If Image17.Height = 1020 And Image17.Width = 1065 Then
Image17.ToolTipText = "Power Plant"
End If
If Image12.Height = 1020 And Image12.Width = 1065 Then
Image12.ToolTipText = "Power Plant"
End If
If Image8.Height = 1020 And Image8.Width = 1065 Then
Image8.ToolTipText = "Power Plant"
End If
If Image7.Height = 1020 And Image7.Width = 1065 Then
Image7.ToolTipText = "Power Plant"
End If
If Image15.Height = 1020 And Image15.Width = 1065 Then
Image15.ToolTipText = "Power Plant"
End If
If Image14.Height = 1020 And Image14.Width = 1065 Then
Image14.ToolTipText = "Power Plant"
End If
If Image9.Height = 1020 And Image9.Width = 1065 Then
Image9.ToolTipText = "Power Plant"
End If
If Image16.Height = 1020 And Image16.Width = 1065 Then
Image16.ToolTipText = "Power Plant"
End If
If Image13.Height = 1020 And Image13.Width = 1065 Then
Image13.ToolTipText = "Power Plant"
End If
If Image28.Height = 1020 And Image28.Width = 1065 Then
Image28.ToolTipText = "Power Plant"
End If
If Image29.Height = 1020 And Image29.Width = 1065 Then
Image29.ToolTipText = "Power Plant"
End If
If Image32.Height = 1020 And Image32.Width = 1065 Then
Image32.ToolTipText = "Power Plant"
End If
If Image133.Height = 1020 And Image133.Width = 1065 Then
Image133.ToolTipText = "Power Plant"
End If
If Image86.Height = 1020 And Image86.Width = 1065 Then
Image86.ToolTipText = "Power Plant"
End If
If Image126.Height = 1020 And Image126.Width = 1065 Then
Image126.ToolTipText = "Power Plant"
End If
If Image151.Height = 1020 And Image151.Width = 1065 Then
Image151.ToolTipText = "Power Plant"
End If
If Image69.Height = 1020 And Image69.Width = 1065 Then
Image69.ToolTipText = "Power Plant"
End If
If Image96.Height = 1020 And Image96.Width = 1065 Then
Image96.ToolTipText = "Power Plant"
End If
If Image78.Height = 1020 And Image78.Width = 1065 Then
Image78.ToolTipText = "Power Plant"
End If
If Image88.Height = 1020 And Image88.Width = 1065 Then
Image88.ToolTipText = "Power Plant"
End If
If Image66.Height = 1020 And Image66.Width = 1065 Then
Image66.ToolTipText = "Power Plant"
End If
If Image118.Height = 1020 And Image118.Width = 1065 Then
Image118.ToolTipText = "Power Plant"
End If
If Image72.Height = 1020 And Image72.Width = 1065 Then
Image72.ToolTipText = "Power Plant"
End If
If Image110.Height = 1020 And Image110.Width = 1065 Then
Image110.ToolTipText = "Power Plant"
End If
If Image135.Height = 1020 And Image135.Width = 1065 Then
Image135.ToolTipText = "Power Plant"
End If
If Image57.Height = 1020 And Image57.Width = 1065 Then
Image57.ToolTipText = "Power Plant"
End If
If Image77.Height = 1020 And Image77.Width = 1065 Then
Image77.ToolTipText = "Power Plant"
End If
If Image85.Height = 1020 And Image85.Width = 1065 Then
Image85.ToolTipText = "Power Plant"
End If
If Image83.Height = 1020 And Image83.Width = 1065 Then
Image83.ToolTipText = "Power Plant"
End If
If Image56.Height = 1020 And Image56.Width = 1065 Then
Image56.ToolTipText = "Power Plant"
End If
If Image93.Height = 1020 And Image93.Width = 1065 Then
Image93.ToolTipText = "Power Plant"
End If
If Image107.Height = 1020 And Image107.Width = 1065 Then
Image107.ToolTipText = "Power Plant"
End If
If Image94.Height = 1020 And Image94.Width = 1065 Then
Image94.ToolTipText = "Commications"
End If
If Image30.Height = 1020 And Image30.Width = 1065 Then
Image30.ToolTipText = "Power Plant"
End If
If Image100.Height = 1020 And Image100.Width = 1065 Then
Image100.ToolTipText = "Power Plant"
End If
If Image101.Height = 1020 And Image101.Width = 1065 Then
Image101.ToolTipText = "Power Plant"
End If
If Image84.Height = 1020 And Image84.Width = 1065 Then
Image84.ToolTipText = "Power Plant"
End If
If Image54.Height = 1020 And Image54.Width = 1065 Then
Image54.ToolTipText = "Power Plant"
End If
If Image129.Height = 1020 And Image129.Width = 1065 Then
Image129.ToolTipText = "Power Plant"
End If
If Image76.Height = 1020 And Image76.Width = 1065 Then
Image76.ToolTipText = "Power Plant"
End If
If Image134.Height = 1020 And Image134.Width = 1065 Then
Image134.ToolTipText = "Power Plant"
End If
If Image111.Height = 1020 And Image111.Width = 1065 Then
Image111.ToolTipText = "Power Plant"
End If
If Image74.Height = 1020 And Image74.Width = 1065 Then
Image74.ToolTipText = "Power Plant"
End If
If Image80.Height = 1020 And Image80.Width = 1065 Then
Image80.ToolTipText = "Power Plant"
End If
If Image136.Height = 1020 And Image136.Width = 1065 Then
Image136.ToolTipText = "Power Plant"
End If
If Image144.Height = 1020 And Image144.Width = 1065 Then
Image144.ToolTipText = "Power Plant"
End If
If Image119.Height = 1020 And Image119.Width = 1065 Then
Image119.ToolTipText = "Power Plant"
End If
If Image147.Height = 1020 And Image147.Width = 1065 Then
Image147.ToolTipText = "Power Plant"
End If
If Image128.Height = 1020 And Image128.Width = 1065 Then
Image128.ToolTipText = "Power Plant"
End If
If Image117.Height = 1020 And Image117.Width = 1065 Then
Image117.ToolTipText = "Power Plant"
End If
If Image143.Height = 1020 And Image143.Width = 1065 Then
Image143.ToolTipText = "Power Plant"
End If
If Image55.Height = 1020 And Image55.Width = 1065 Then
Image55.ToolTipText = "Power Plant"
End If
If Image108.Height = 1020 And Image108.Width = 1065 Then
Image108.ToolTipText = "Power Plant"
End If
If Image89.Height = 1020 And Image89.Width = 1065 Then
Image89.ToolTipText = "Power Plant"
End If
If Image91.Height = 1020 And Image91.Width = 1065 Then
Image91.ToolTipText = "Power Plant"
End If
If Image97.Height = 1020 And Image97.Width = 1065 Then
Image97.ToolTipText = "Power Plant"
End If
If Image31.Height = 1020 And Image31.Width = 1065 Then
Image31.ToolTipText = "Power Plant"
End If
If Image70.Height = 1020 And Image70.Width = 1065 Then
Image70.ToolTipText = "Power Plant"
End If
If Image75.Height = 1020 And Image75.Width = 1065 Then
Image75.ToolTipText = "Power Plant"
End If
If Image64.Height = 1020 And Image64.Width = 1065 Then
Image64.ToolTipText = "Power Plant"
End If
If Image104.Height = 1020 And Image104.Width = 1065 Then
Image104.ToolTipText = "Power Plant"
End If
If Image99.Height = 1020 And Image99.Width = 1065 Then
Image99.ToolTipText = "Power Plant"
End If
If Image112.Height = 1020 And Image112.Width = 1065 Then
Image112.ToolTipText = "Power Plant"
End If
If Image92.Height = 1020 And Image92.Width = 1065 Then
Image92.ToolTipText = "Power Plant"
End If
If Image127.Height = 1020 And Image127.Width = 1065 Then
Image127.ToolTipText = "Power Plant"
End If
If Image95.Height = 1020 And Image95.Width = 1065 Then
Image95.ToolTipText = "Power Plant"
End If
If Image124.Height = 1020 And Image124.Width = 1065 Then
Image124.ToolTipText = "Power Plant"
End If
If Image123.Height = 1020 And Image123.Width = 1065 Then
Image123.ToolTipText = "Power Plant"
End If
If Image125.Height = 1020 And Image125.Width = 1065 Then
Image125.ToolTipText = "Power Plant"
End If
If Image48.Height = 1020 And Image48.Width = 1065 Then
Image48.ToolTipText = "Power Plant"
End If
If Image52.Height = 1020 And Image52.Width = 1065 Then
Image52.ToolTipText = "Power Plant"
End If
If Image52.Height = 1020 And Image52.Width = 1065 Then
Image52.ToolTipText = "Power Plant"
End If
If Image47.Height = 1020 And Image47.Width = 1065 Then
Image47.ToolTipText = "Power Plant"
End If
If Image49.Height = 1020 And Image49.Width = 1065 Then
Image49.ToolTipText = "Power Plant"
End If
If Image50.Height = 1020 And Image50.Width = 1065 Then
Image50.ToolTipText = "Power Plant"
End If
If Image53.Height = 1020 And Image53.Width = 1065 Then
Image53.ToolTipText = "Power Plant"
End If
If Image132.Height = 1020 And Image132.Width = 1065 Then
Image132.ToolTipText = "Power Plant"
End If
If Image140.Height = 1020 And Image140.Width = 1065 Then
Image140.ToolTipText = "Power Plant"
End If
If Image59.Height = 1020 And Image59.Width = 1065 Then
Image59.ToolTipText = "Power Plant"
End If
If Image65.Height = 1020 And Image65.Width = 1065 Then
Image65.ToolTipText = "Power Plant"
End If
If Image67.Height = 1020 And Image67.Width = 1065 Then
Image67.ToolTipText = "Power Plant"
End If
If Image109.Height = 1020 And Image109.Width = 1065 Then
Image109.ToolTipText = "Power Plant"
End If
If Image73.Height = 1020 And Image73.Width = 1065 Then
Image73.ToolTipText = "Power Plant"
End If
If Image120.Height = 1020 And Image120.Width = 1065 Then
Image120.ToolTipText = "Power Plant"
End If
If Image58.Height = 1020 And Image58.Width = 1065 Then
Image58.ToolTipText = "Power Plant"
End If
If Image51.Height = 1020 And Image51.Width = 1065 Then
Image51.ToolTipText = "Power Plant"
End If
If Image45.Height = 1020 And Image45.Width = 1065 Then
Image45.ToolTipText = "Power Plant"
End If
If Image71.Height = 1020 And Image71.Width = 1065 Then
Image71.ToolTipText = "Power Plant"
End If
If Image68.Height = 1020 And Image68.Width = 1065 Then
Image68.ToolTipText = "Power Plant"
End If
If Image105.Height = 1020 And Image105.Width = 1065 Then
Image105.ToolTipText = "Power Plant"
End If
If Image106.Height = 1020 And Image106.Width = 1065 Then
Image106.ToolTipText = "Power Plant"
End If
If Image115.Height = 1020 And Image115.Width = 1065 Then
Image115.ToolTipText = "Power Plant"
End If
If Image153.Height = 1020 And Image153.Width = 1065 Then
Image153.ToolTipText = "Power Plant"
End If
If Image139.Height = 1020 And Image139.Width = 1065 Then
Image139.ToolTipText = "Power Plant"
End If
If Image116.Height = 1020 And Image116.Width = 1065 Then
Image116.ToolTipText = "Power Plant"
End If
If Image145.Height = 1020 And Image145.Width = 1065 Then
Image145.ToolTipText = "Power Plant"
End If
If Image142.Height = 1020 And Image142.Width = 1065 Then
Image142.ToolTipText = "Power Plant"
End If
If Image146.Height = 1020 And Image146.Width = 1065 Then
Image146.ToolTipText = "Power Plant"
End If
If Image130.Height = 1020 And Image130.Width = 1065 Then
Image130.ToolTipText = "Power Plant"
End If
If Image121.Height = 1020 And Image121.Width = 1065 Then
Image121.ToolTipText = "Power Plant"
End If
If Image46.Height = 1020 And Image46.Width = 1065 Then
Image46.ToolTipText = "Power Plant"
End If
If Image44.Height = 1020 And Image44.Width = 1065 Then
Image44.ToolTipText = "Power Plant"
End If
If Image63.Height = 1020 And Image63.Width = 1065 Then
Image63.ToolTipText = "Power Plant"
End If
If Image60.Height = 1020 And Image60.Width = 1065 Then
Image60.ToolTipText = "Power Plant"
End If
If Image61.Height = 1020 And Image61.Width = 1065 Then
Image61.ToolTipText = "Power Plant"
End If
If Image62.Height = 1020 And Image62.Width = 1065 Then
Image115.ToolTipText = "Power Plant"
End If
If Image131.Height = 1020 And Image131.Width = 1065 Then
Image131.ToolTipText = "Power Plant"
End If
If Image148.Height = 1020 And Image148.Width = 1065 Then
Image148.ToolTipText = "Power Plant"
End If
If Image154.Height = 1020 And Image154.Width = 1065 Then
Image154.ToolTipText = "Power Plant"
End If
If Image141.Height = 1020 And Image141.Width = 1065 Then
Image141.ToolTipText = "Power Plant"
End If
If Image113.Height = 1020 And Image113.Width = 1065 Then
Image113.ToolTipText = "Power Plant"
End If
If Image152.Height = 1020 And Image152.Width = 1065 Then
Image152.ToolTipText = "Power Plant"
End If
If Image150.Height = 1020 And Image150.Width = 1065 Then
Image150.ToolTipText = "Power Plant"
End If
If Image102.Height = 1020 And Image102.Width = 1065 Then
Image102.ToolTipText = "Power Plant"
End If
If Image114.Height = 1020 And Image114.Width = 1065 Then
Image114.ToolTipText = "Power Plant"
End If
If Image103.Height = 1020 And Image103.Width = 1065 Then
Image103.ToolTipText = "Power Plant"
End If
If Image79.Height = 1020 And Image79.Width = 1065 Then
Image79.ToolTipText = "Power Plant"
End If
If Image141.Height = 1020 And Image141.Width = 1065 Then
Image141.ToolTipText = "Power Plant"
End If

End Sub
Public Sub ORE()
If Image43.Height = 1260 And Image43.Width = 1230 Then
Image43.ToolTipText = "Refinary"
If Image18.Height = 495 And Image18.Width = 975 Then
Image18.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image18.Visible = False
End If

End If
If Image42.Height = 1260 And Image42.Width = 1065 Then
Image42.ToolTipText = "Power Plant"
If Image43.Height = 495 And Image43.Width = 975 Then
Image43.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image43.Visible = False
End If
End If

If Image18.Height = 1260 And Image18.Width = 1065 Then
Image18.ToolTipText = "Power Plant"
If Image42.Height = 495 And Image42.Width = 975 Then
Image42.Picture = LoadPicture("c:\windows\desktop\vb projects\war\status3.bmp")
Image42.Visible = False
End If
End If
If Image19.Height = 1260 And Image19.Width = 1065 Then
Image19.ToolTipText = "Power Plant"
End If
If Image22.Height = 1260 And Image22.Width = 1065 Then
Image22.ToolTipText = "Power Plant"
End If
If Image10.Height = 1260 And Image10.Width = 1065 Then
Image10.ToolTipText = "Power Plant"
End If
If Image11.Height = 1260 And Image11.Width = 1065 Then
Image11.ToolTipText = "Power Plant"
End If
If Image21.Height = 1260 And Image21.Width = 1065 Then
Image21.ToolTipText = "Power Plant"
End If
If Image26.Height = 1020 And Image26.Width = 1065 Then
Image26.ToolTipText = "Power Plant"
End If
If Image23.Height = 1020 And Image23.Width = 1065 Then
Image23.ToolTipText = "Power Plant"
End If
If Image20.Height = 1020 And Image20.Width = 1065 Then
Image20.ToolTipText = "Power Plant"
End If
If Image24.Height = 1020 And Image24.Width = 1065 Then
Image24.ToolTipText = "Power Plant"
End If
If Image2.Height = 1020 And Image2.Width = 1065 Then
Image2.ToolTipText = "Power Plant"
End If
If Image3.Height = 1020 And Image3.Width = 1065 Then
Image3.ToolTipText = "Power Plant"
End If
If Image5.Height = 1020 And Image5.Width = 1065 Then
Image5.ToolTipText = "Power Plant"
End If
If Image4.Height = 1020 And Image4.Width = 1065 Then
Image4.ToolTipText = "Power Plant"
End If
If Image6.Height = 1020 And Image6.Width = 1065 Then
Image6.ToolTipText = "Power Plant"
End If
If Image1.Height = 1020 And Image1.Width = 1065 Then
Image1.ToolTipText = "Power Plant"
End If
If Image17.Height = 1020 And Image17.Width = 1065 Then
Image17.ToolTipText = "Power Plant"
End If
If Image12.Height = 1020 And Image12.Width = 1065 Then
Image12.ToolTipText = "Power Plant"
End If
If Image8.Height = 1020 And Image8.Width = 1065 Then
Image8.ToolTipText = "Power Plant"
End If
If Image7.Height = 1020 And Image7.Width = 1065 Then
Image7.ToolTipText = "Power Plant"
End If
If Image15.Height = 1020 And Image15.Width = 1065 Then
Image15.ToolTipText = "Power Plant"
End If
If Image14.Height = 1020 And Image14.Width = 1065 Then
Image14.ToolTipText = "Power Plant"
End If
If Image9.Height = 1020 And Image9.Width = 1065 Then
Image9.ToolTipText = "Power Plant"
End If
If Image16.Height = 1020 And Image16.Width = 1065 Then
Image16.ToolTipText = "Power Plant"
End If
If Image13.Height = 1020 And Image13.Width = 1065 Then
Image13.ToolTipText = "Power Plant"
End If
If Image28.Height = 1020 And Image28.Width = 1065 Then
Image28.ToolTipText = "Power Plant"
End If
If Image29.Height = 1020 And Image29.Width = 1065 Then
Image29.ToolTipText = "Power Plant"
End If
If Image32.Height = 1020 And Image32.Width = 1065 Then
Image32.ToolTipText = "Power Plant"
End If
If Image133.Height = 1020 And Image133.Width = 1065 Then
Image133.ToolTipText = "Power Plant"
End If
If Image86.Height = 1020 And Image86.Width = 1065 Then
Image86.ToolTipText = "Power Plant"
End If
If Image126.Height = 1020 And Image126.Width = 1065 Then
Image126.ToolTipText = "Power Plant"
End If
If Image151.Height = 1020 And Image151.Width = 1065 Then
Image151.ToolTipText = "Power Plant"
End If
If Image69.Height = 1020 And Image69.Width = 1065 Then
Image69.ToolTipText = "Power Plant"
End If
If Image96.Height = 1020 And Image96.Width = 1065 Then
Image96.ToolTipText = "Power Plant"
End If
If Image78.Height = 1020 And Image78.Width = 1065 Then
Image78.ToolTipText = "Power Plant"
End If
If Image88.Height = 1020 And Image88.Width = 1065 Then
Image88.ToolTipText = "Power Plant"
End If
If Image66.Height = 1020 And Image66.Width = 1065 Then
Image66.ToolTipText = "Power Plant"
End If
If Image118.Height = 1020 And Image118.Width = 1065 Then
Image118.ToolTipText = "Power Plant"
End If
If Image72.Height = 1020 And Image72.Width = 1065 Then
Image72.ToolTipText = "Power Plant"
End If
If Image110.Height = 1020 And Image110.Width = 1065 Then
Image110.ToolTipText = "Power Plant"
End If
If Image135.Height = 1020 And Image135.Width = 1065 Then
Image135.ToolTipText = "Power Plant"
End If
If Image57.Height = 1020 And Image57.Width = 1065 Then
Image57.ToolTipText = "Power Plant"
End If
If Image77.Height = 1020 And Image77.Width = 1065 Then
Image77.ToolTipText = "Power Plant"
End If
If Image85.Height = 1020 And Image85.Width = 1065 Then
Image85.ToolTipText = "Power Plant"
End If
If Image83.Height = 1020 And Image83.Width = 1065 Then
Image83.ToolTipText = "Power Plant"
End If
If Image56.Height = 1020 And Image56.Width = 1065 Then
Image56.ToolTipText = "Power Plant"
End If
If Image93.Height = 1020 And Image93.Width = 1065 Then
Image93.ToolTipText = "Power Plant"
End If
If Image107.Height = 1020 And Image107.Width = 1065 Then
Image107.ToolTipText = "Power Plant"
End If
If Image94.Height = 1020 And Image94.Width = 1065 Then
Image94.ToolTipText = "Commications"
End If
If Image30.Height = 1020 And Image30.Width = 1065 Then
Image30.ToolTipText = "Power Plant"
End If
If Image100.Height = 1020 And Image100.Width = 1065 Then
Image100.ToolTipText = "Power Plant"
End If
If Image101.Height = 1020 And Image101.Width = 1065 Then
Image101.ToolTipText = "Power Plant"
End If
If Image84.Height = 1020 And Image84.Width = 1065 Then
Image84.ToolTipText = "Power Plant"
End If
If Image54.Height = 1020 And Image54.Width = 1065 Then
Image54.ToolTipText = "Power Plant"
End If
If Image129.Height = 1020 And Image129.Width = 1065 Then
Image129.ToolTipText = "Power Plant"
End If
If Image76.Height = 1020 And Image76.Width = 1065 Then
Image76.ToolTipText = "Power Plant"
End If
If Image134.Height = 1020 And Image134.Width = 1065 Then
Image134.ToolTipText = "Power Plant"
End If
If Image111.Height = 1020 And Image111.Width = 1065 Then
Image111.ToolTipText = "Power Plant"
End If
If Image74.Height = 1020 And Image74.Width = 1065 Then
Image74.ToolTipText = "Power Plant"
End If
If Image80.Height = 1020 And Image80.Width = 1065 Then
Image80.ToolTipText = "Power Plant"
End If
If Image136.Height = 1020 And Image136.Width = 1065 Then
Image136.ToolTipText = "Power Plant"
End If
If Image144.Height = 1020 And Image144.Width = 1065 Then
Image144.ToolTipText = "Power Plant"
End If
If Image119.Height = 1020 And Image119.Width = 1065 Then
Image119.ToolTipText = "Power Plant"
End If
If Image147.Height = 1020 And Image147.Width = 1065 Then
Image147.ToolTipText = "Power Plant"
End If
If Image128.Height = 1020 And Image128.Width = 1065 Then
Image128.ToolTipText = "Power Plant"
End If
If Image117.Height = 1020 And Image117.Width = 1065 Then
Image117.ToolTipText = "Power Plant"
End If
If Image143.Height = 1020 And Image143.Width = 1065 Then
Image143.ToolTipText = "Power Plant"
End If
If Image55.Height = 1020 And Image55.Width = 1065 Then
Image55.ToolTipText = "Power Plant"
End If
If Image108.Height = 1020 And Image108.Width = 1065 Then
Image108.ToolTipText = "Power Plant"
End If
If Image89.Height = 1020 And Image89.Width = 1065 Then
Image89.ToolTipText = "Power Plant"
End If
If Image91.Height = 1020 And Image91.Width = 1065 Then
Image91.ToolTipText = "Power Plant"
End If
If Image97.Height = 1020 And Image97.Width = 1065 Then
Image97.ToolTipText = "Power Plant"
End If
If Image31.Height = 1020 And Image31.Width = 1065 Then
Image31.ToolTipText = "Power Plant"
End If
If Image70.Height = 1020 And Image70.Width = 1065 Then
Image70.ToolTipText = "Power Plant"
End If
If Image75.Height = 1020 And Image75.Width = 1065 Then
Image75.ToolTipText = "Power Plant"
End If
If Image64.Height = 1020 And Image64.Width = 1065 Then
Image64.ToolTipText = "Power Plant"
End If
If Image104.Height = 1020 And Image104.Width = 1065 Then
Image104.ToolTipText = "Power Plant"
End If
If Image99.Height = 1020 And Image99.Width = 1065 Then
Image99.ToolTipText = "Power Plant"
End If
If Image112.Height = 1020 And Image112.Width = 1065 Then
Image112.ToolTipText = "Power Plant"
End If
If Image92.Height = 1020 And Image92.Width = 1065 Then
Image92.ToolTipText = "Power Plant"
End If
If Image127.Height = 1020 And Image127.Width = 1065 Then
Image127.ToolTipText = "Power Plant"
End If
If Image95.Height = 1020 And Image95.Width = 1065 Then
Image95.ToolTipText = "Power Plant"
End If
If Image124.Height = 1020 And Image124.Width = 1065 Then
Image124.ToolTipText = "Power Plant"
End If
If Image123.Height = 1020 And Image123.Width = 1065 Then
Image123.ToolTipText = "Power Plant"
End If
If Image125.Height = 1020 And Image125.Width = 1065 Then
Image125.ToolTipText = "Power Plant"
End If
If Image48.Height = 1020 And Image48.Width = 1065 Then
Image48.ToolTipText = "Power Plant"
End If
If Image52.Height = 1020 And Image52.Width = 1065 Then
Image52.ToolTipText = "Power Plant"
End If
If Image52.Height = 1020 And Image52.Width = 1065 Then
Image52.ToolTipText = "Power Plant"
End If
If Image47.Height = 1020 And Image47.Width = 1065 Then
Image47.ToolTipText = "Power Plant"
End If
If Image49.Height = 1020 And Image49.Width = 1065 Then
Image49.ToolTipText = "Power Plant"
End If
If Image50.Height = 1020 And Image50.Width = 1065 Then
Image50.ToolTipText = "Power Plant"
End If
If Image53.Height = 1020 And Image53.Width = 1065 Then
Image53.ToolTipText = "Power Plant"
End If
If Image132.Height = 1020 And Image132.Width = 1065 Then
Image132.ToolTipText = "Power Plant"
End If
If Image140.Height = 1020 And Image140.Width = 1065 Then
Image140.ToolTipText = "Power Plant"
End If
If Image59.Height = 1020 And Image59.Width = 1065 Then
Image59.ToolTipText = "Power Plant"
End If
If Image65.Height = 1020 And Image65.Width = 1065 Then
Image65.ToolTipText = "Power Plant"
End If
If Image67.Height = 1020 And Image67.Width = 1065 Then
Image67.ToolTipText = "Power Plant"
End If
If Image109.Height = 1020 And Image109.Width = 1065 Then
Image109.ToolTipText = "Power Plant"
End If
If Image73.Height = 1020 And Image73.Width = 1065 Then
Image73.ToolTipText = "Power Plant"
End If
If Image120.Height = 1020 And Image120.Width = 1065 Then
Image120.ToolTipText = "Power Plant"
End If
If Image58.Height = 1020 And Image58.Width = 1065 Then
Image58.ToolTipText = "Power Plant"
End If
If Image51.Height = 1020 And Image51.Width = 1065 Then
Image51.ToolTipText = "Power Plant"
End If
If Image45.Height = 1020 And Image45.Width = 1065 Then
Image45.ToolTipText = "Power Plant"
End If
If Image71.Height = 1020 And Image71.Width = 1065 Then
Image71.ToolTipText = "Power Plant"
End If
If Image68.Height = 1020 And Image68.Width = 1065 Then
Image68.ToolTipText = "Power Plant"
End If
If Image105.Height = 1020 And Image105.Width = 1065 Then
Image105.ToolTipText = "Power Plant"
End If
If Image106.Height = 1020 And Image106.Width = 1065 Then
Image106.ToolTipText = "Power Plant"
End If
If Image115.Height = 1020 And Image115.Width = 1065 Then
Image115.ToolTipText = "Power Plant"
End If
If Image153.Height = 1020 And Image153.Width = 1065 Then
Image153.ToolTipText = "Power Plant"
End If
If Image139.Height = 1020 And Image139.Width = 1065 Then
Image139.ToolTipText = "Power Plant"
End If
If Image116.Height = 1020 And Image116.Width = 1065 Then
Image116.ToolTipText = "Power Plant"
End If
If Image145.Height = 1020 And Image145.Width = 1065 Then
Image145.ToolTipText = "Power Plant"
End If
If Image142.Height = 1020 And Image142.Width = 1065 Then
Image142.ToolTipText = "Power Plant"
End If
If Image146.Height = 1020 And Image146.Width = 1065 Then
Image146.ToolTipText = "Power Plant"
End If
If Image130.Height = 1020 And Image130.Width = 1065 Then
Image130.ToolTipText = "Power Plant"
End If
If Image121.Height = 1020 And Image121.Width = 1065 Then
Image121.ToolTipText = "Power Plant"
End If
If Image46.Height = 1020 And Image46.Width = 1065 Then
Image46.ToolTipText = "Power Plant"
End If
If Image44.Height = 1020 And Image44.Width = 1065 Then
Image44.ToolTipText = "Power Plant"
End If
If Image63.Height = 1020 And Image63.Width = 1065 Then
Image63.ToolTipText = "Power Plant"
End If
If Image60.Height = 1020 And Image60.Width = 1065 Then
Image60.ToolTipText = "Power Plant"
End If
If Image61.Height = 1020 And Image61.Width = 1065 Then
Image61.ToolTipText = "Power Plant"
End If
If Image62.Height = 1020 And Image62.Width = 1065 Then
Image115.ToolTipText = "Power Plant"
End If
If Image131.Height = 1020 And Image131.Width = 1065 Then
Image131.ToolTipText = "Power Plant"
End If
If Image148.Height = 1020 And Image148.Width = 1065 Then
Image148.ToolTipText = "Power Plant"
End If
If Image154.Height = 1020 And Image154.Width = 1065 Then
Image154.ToolTipText = "Power Plant"
End If
If Image141.Height = 1020 And Image141.Width = 1065 Then
Image141.ToolTipText = "Power Plant"
End If
If Image113.Height = 1020 And Image113.Width = 1065 Then
Image113.ToolTipText = "Power Plant"
End If
If Image152.Height = 1020 And Image152.Width = 1065 Then
Image152.ToolTipText = "Power Plant"
End If
If Image150.Height = 1020 And Image150.Width = 1065 Then
Image150.ToolTipText = "Power Plant"
End If
If Image102.Height = 1020 And Image102.Width = 1065 Then
Image102.ToolTipText = "Power Plant"
End If
If Image114.Height = 1020 And Image114.Width = 1065 Then
Image114.ToolTipText = "Power Plant"
End If
If Image103.Height = 1020 And Image103.Width = 1065 Then
Image103.ToolTipText = "Power Plant"
End If
If Image79.Height = 1020 And Image79.Width = 1065 Then
Image79.ToolTipText = "Power Plant"
End If
If Image141.Height = 1020 And Image141.Width = 1065 Then
Image141.ToolTipText = "Power Plant"
End If

End Sub

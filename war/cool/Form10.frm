VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form10 
   Caption         =   "Space"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      Top             =   4560
      Width           =   495
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
      Filename        =   "C:\WINDOWS\Desktop\vb projects\war\solrwind.wav"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   -1  'True
      PlayCount       =   0
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
      Volume          =   -500
      WindowlessVideo =   0   'False
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3100
      Left            =   4680
      Top             =   4560
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   495
      Left            =   5040
      TabIndex        =   26
      Top             =   4560
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   4680
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   4680
      Top             =   4560
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   15000
      TabIndex        =   22
      Top             =   4200
      Width           =   7335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4455
      Left            =   7320
      Max             =   15000
      TabIndex        =   21
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Height          =   735
      Left            =   8280
      Picture         =   "Form10.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "down"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   7560
      Picture         =   "Form10.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "up"
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox Picture5 
      DragMode        =   1  'Automatic
      Height          =   1095
      Left            =   7560
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "Form10.frx":091E
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   18
      ToolTipText     =   "Bunker"
      Top             =   2760
      Width           =   1335
      Begin VB.PictureBox Picture6 
         Height          =   1095
         Left            =   0
         Picture         =   "Form10.frx":2D10
         ScaleHeight     =   1035
         ScaleWidth      =   1275
         TabIndex        =   23
         ToolTipText     =   "Supply depot"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         Begin MSComctlLib.ProgressBar ProgressBar4 
            Height          =   135
            Left            =   -120
            TabIndex        =   30
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label15 
            BackColor       =   &H00000000&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox Picture4 
      DragMode        =   1  'Automatic
      Height          =   975
      Left            =   7560
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "Form10.frx":5312
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   17
      ToolTipText     =   "Factory"
      Top             =   1800
      Width           =   1335
      Begin VB.PictureBox Picture7 
         Height          =   975
         Left            =   0
         Picture         =   "Form10.frx":AA94
         ScaleHeight     =   915
         ScaleWidth      =   1275
         TabIndex        =   24
         ToolTipText     =   "gas"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         Begin ComctlLib.ProgressBar ProgressBar3 
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label Label14 
            BackColor       =   &H00000000&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   7560
      Picture         =   "Form10.frx":E2EA
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   16
      ToolTipText     =   "barrecks"
      Top             =   840
      Width           =   1335
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   135
         Left            =   0
         TabIndex        =   28
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   7560
      Picture         =   "Form10.frx":12A34
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   15
      ToolTipText     =   "Communications"
      Top             =   0
      Width           =   1335
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   0
         TabIndex        =   27
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View"
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   7515
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   1005
         ButtonWidth     =   794
         ButtonHeight    =   847
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   6
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
         EndProperty
      End
      Begin VB.Image Image67 
         Height          =   735
         Left            =   0
         OLEDropMode     =   2  'Automatic
         Top             =   1560
         Width           =   735
      End
      Begin VB.Image Image66 
         Height          =   735
         Left            =   720
         OLEDropMode     =   2  'Automatic
         Top             =   1560
         Width           =   735
      End
      Begin VB.Image Image65 
         Height          =   615
         Left            =   0
         OLEDropMode     =   2  'Automatic
         Top             =   2280
         Width           =   735
      End
      Begin VB.Image Image64 
         Height          =   615
         Left            =   720
         OLEDropMode     =   2  'Automatic
         Top             =   2280
         Width           =   735
      End
      Begin VB.Image Image63 
         Height          =   615
         Left            =   0
         OLEDropMode     =   2  'Automatic
         Top             =   2880
         Width           =   735
      End
      Begin VB.Image Image62 
         Height          =   615
         Left            =   720
         OLEDropMode     =   2  'Automatic
         Top             =   2880
         Width           =   735
      End
      Begin VB.Image Image61 
         Height          =   855
         Left            =   0
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Image60 
         Height          =   855
         Left            =   600
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   855
      End
      Begin VB.Image Image59 
         Height          =   495
         Left            =   1440
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image Image58 
         Height          =   735
         Left            =   0
         OLEDropMode     =   2  'Automatic
         Top             =   0
         Width           =   975
      End
      Begin VB.Image Image57 
         Height          =   735
         Left            =   960
         OLEDropMode     =   2  'Automatic
         Top             =   0
         Width           =   735
      End
      Begin VB.Image Image56 
         Height          =   615
         Left            =   1680
         OLEDropMode     =   2  'Automatic
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image55 
         Height          =   735
         Left            =   2520
         OLEDropMode     =   2  'Automatic
         Top             =   0
         Width           =   735
      End
      Begin VB.Image Image54 
         Height          =   615
         Left            =   2520
         OLEDropMode     =   2  'Automatic
         Top             =   600
         Width           =   735
      End
      Begin VB.Image Image53 
         Height          =   1095
         Left            =   1440
         OLEDropMode     =   2  'Automatic
         Top             =   1200
         Width           =   975
      End
      Begin VB.Image Image52 
         Height          =   735
         Left            =   2400
         OLEDropMode     =   2  'Automatic
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image Image51 
         Height          =   615
         Left            =   1440
         OLEDropMode     =   2  'Automatic
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Image Image50 
         Height          =   615
         Left            =   1440
         OLEDropMode     =   2  'Automatic
         Top             =   2880
         Width           =   855
      End
      Begin VB.Image Image49 
         Height          =   615
         Left            =   0
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   735
      End
      Begin VB.Image Image48 
         Height          =   615
         Left            =   720
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   735
      End
      Begin VB.Image Image47 
         Height          =   615
         Left            =   1440
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image46 
         Height          =   735
         Left            =   3240
         OLEDropMode     =   2  'Automatic
         Top             =   0
         Width           =   735
      End
      Begin VB.Image Image45 
         Height          =   495
         Left            =   3240
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image44 
         Height          =   735
         Left            =   3120
         OLEDropMode     =   2  'Automatic
         Top             =   1200
         Width           =   855
      End
      Begin VB.Image Image43 
         Height          =   735
         Left            =   3960
         OLEDropMode     =   2  'Automatic
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image Image42 
         Height          =   615
         Left            =   3960
         OLEDropMode     =   2  'Automatic
         Top             =   2160
         Width           =   855
      End
      Begin VB.Image Image41 
         Height          =   615
         Left            =   4800
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image40 
         Height          =   495
         Left            =   2280
         OLEDropMode     =   2  'Automatic
         Top             =   3000
         Width           =   855
      End
      Begin VB.Image Image39 
         Height          =   615
         Left            =   2280
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image38 
         Height          =   495
         Left            =   3120
         OLEDropMode     =   2  'Automatic
         Top             =   3000
         Width           =   855
      End
      Begin VB.Image Image37 
         Height          =   735
         Left            =   3960
         OLEDropMode     =   2  'Automatic
         Top             =   2760
         Width           =   855
      End
      Begin VB.Image Image36 
         Height          =   615
         Left            =   3120
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image35 
         Height          =   615
         Left            =   3960
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image34 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   4800
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image33 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   5640
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   735
      End
      Begin VB.Image Image32 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   6360
         OLEDropMode     =   2  'Automatic
         Top             =   3480
         Width           =   855
      End
      Begin VB.Image Image31 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   3960
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   855
      End
      Begin VB.Image Image30 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   6360
         OLEDropMode     =   2  'Automatic
         Top             =   2760
         Width           =   855
      End
      Begin VB.Image Image29 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   5640
         OLEDropMode     =   2  'Automatic
         Top             =   2760
         Width           =   735
      End
      Begin VB.Image Image28 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   4800
         OLEDropMode     =   2  'Automatic
         Top             =   2760
         Width           =   855
      End
      Begin VB.Image Image27 
         DragMode        =   1  'Automatic
         Height          =   615
         Left            =   4800
         OLEDropMode     =   2  'Automatic
         Top             =   2160
         Width           =   855
      End
      Begin VB.Image Image26 
         DragMode        =   1  'Automatic
         Height          =   615
         Left            =   5640
         OLEDropMode     =   2  'Automatic
         Top             =   2160
         Width           =   735
      End
      Begin VB.Image Image25 
         DragMode        =   1  'Automatic
         Height          =   615
         Left            =   6360
         OLEDropMode     =   2  'Automatic
         Top             =   2160
         Width           =   855
      End
      Begin VB.Image Image24 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   4800
         OLEDropMode     =   2  'Automatic
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image Image23 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   4800
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   855
      End
      Begin VB.Image Image22 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   5640
         OLEDropMode     =   2  'Automatic
         Top             =   1440
         Width           =   735
      End
      Begin VB.Image Image21 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   6360
         OLEDropMode     =   2  'Automatic
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image Image20 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   5640
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image19 
         DragMode        =   1  'Automatic
         Height          =   735
         Left            =   6360
         OLEDropMode     =   2  'Automatic
         Top             =   720
         Width           =   855
      End
      Begin VB.Image Image18 
         DragMode        =   1  'Automatic
         Height          =   615
         Left            =   3960
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image17 
         DragMode        =   1  'Automatic
         Height          =   615
         Left            =   4800
         OLEDropMode     =   2  'Automatic
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image16 
         DragMode        =   1  'Automatic
         Height          =   615
         Left            =   5640
         OLEDropMode     =   2  'Automatic
         Top             =   120
         Width           =   735
      End
      Begin VB.Image Image15 
         DragMode        =   1  'Automatic
         Height          =   615
         Left            =   6360
         OLEDropMode     =   2  'Automatic
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image14 
         Height          =   135
         Left            =   2760
         Picture         =   "Form10.frx":15306
         Top             =   1440
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Image Image13 
         Height          =   1095
         Left            =   2640
         ToolTipText     =   "Contruction yard"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Image Image12 
         Height          =   5025
         Left            =   -840
         Picture         =   "Form10.frx":15D44
         Top             =   0
         Visible         =   0   'False
         Width           =   7950
      End
      Begin VB.Image Image10 
         Height          =   3195
         Left            =   3960
         Picture         =   "Form10.frx":980CE
         Top             =   120
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Image Image9 
         Height          =   3600
         Left            =   4080
         Picture         =   "Form10.frx":B9590
         Top             =   0
         Visible         =   0   'False
         Width           =   3330
      End
      Begin VB.Image Image8 
         Height          =   3345
         Left            =   -120
         Picture         =   "Form10.frx":E0812
         Top             =   360
         Visible         =   0   'False
         Width           =   7140
      End
      Begin VB.Image Image7 
         Height          =   1710
         Left            =   480
         Picture         =   "Form10.frx":12E440
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image Image6 
         Height          =   4080
         Left            =   2040
         Picture         =   "Form10.frx":137BEA
         Top             =   1320
         Visible         =   0   'False
         Width           =   4140
      End
      Begin VB.Image Image5 
         Height          =   3435
         Left            =   2640
         Picture         =   "Form10.frx":16EBEC
         Top             =   1800
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.Image Image4 
         Height          =   4665
         Left            =   3000
         Picture         =   "Form10.frx":19454E
         Top             =   0
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Image Image3 
         Height          =   4485
         Left            =   1440
         Picture         =   "Form10.frx":1BC700
         Top             =   -240
         Visible         =   0   'False
         Width           =   4080
      End
      Begin VB.Image Image2 
         Height          =   2025
         Left            =   5160
         Picture         =   "Form10.frx":1F8052
         Top             =   2520
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image1 
         Height          =   8340
         Left            =   1200
         Picture         =   "Form10.frx":203E6C
         Top             =   3120
         Visible         =   0   'False
         Width           =   12525
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Earth"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Moon"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Venus"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Mercury"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Mars"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Jupiter"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Statun"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Uranus"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Neptune"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Pluto"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   4
         Top             =   1920
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   375
      Left            =   3480
      TabIndex        =   36
      Top             =   4560
      Width           =   375
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   26
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":3585BE
            Key             =   "a"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":358D60
            Key             =   "b"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":359832
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":35A544
            Key             =   "d"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":35B0EE
            Key             =   "e"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form10.frx":35BDD8
            Key             =   "f"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image11 
      Height          =   135
      Left            =   2880
      Picture         =   "Form10.frx":35C0F2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "Click at the planet then if you have a rocket ship a constrution yard will be sent"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   3255
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'Dim index As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Form10.WindowState = 1
Label6.Caption = "1"
'Unload Me

End Sub

Private Sub Command2_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image8.Visible = False
Image9.Visible = False
Image10.Visible = False
'Image1.Visible = True
Label1.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Image11.Visible = False
Image12.Visible = False
Image25.Visible = False
Image26.Visible = False
Image27.Visible = False
Image24.Visible = False
Image22.Visible = False
Image21.Visible = False
Image20.Visible = False
Image19.Visible = False
Image23.Visible = False
Image18.Visible = True
Image17.Visible = True
Image16.Visible = True
Image28.Visible = False
Image29.Visible = False
Image30.Visible = False
Image31.Visible = False
Image32.Visible = False
Image15.Visible = False
Image16.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image26.Visible = False
Image27.Visible = False
Image28.Visible = False
Image29.Visible = False
Image30.Visible = False
Image31.Visible = False
Image32.Visible = False
Image33.Visible = False
Image34.Visible = False
Image35.Visible = False
Image36.Visible = False
Image37.Visible = False
Image38.Visible = False
Image39.Visible = False
Image40.Visible = False
Image41.Visible = False
Image42.Visible = False
Image43.Visible = False
Image44.Visible = False
Image45.Visible = False
Image46.Visible = False
Image47.Visible = False
Image48.Visible = False
Image49.Visible = False
Image50.Visible = False
Image51.Visible = False
Image52.Visible = False
Image53.Visible = False
Image54.Visible = False
Image55.Visible = False
Image56.Visible = False
Image57.Visible = False
Image58.Visible = False
Image59.Visible = False
Image60.Visible = False
Image61.Visible = False
Image62.Visible = False
Image63.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
End Sub


Private Sub Command3_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Image12.Visible = True
Image25.Visible = True
Image26.Visible = True
Image27.Visible = True
Image24.Visible = True
Image22.Visible = True
Image21.Visible = True
Image19.Visible = True
Image23.Visible = True
Image20.Visible = True
Image16.Visible = True
Image17.Visible = True
Image18.Visible = True
Image28.Visible = True
Image30.Visible = True
Image31.Visible = True
Image32.Visible = True
Image15.Visible = True
Image29.Visible = True
Image33.Visible = True
Image34.Visible = True
Image35.Visible = True
Image36.Visible = True
Image37.Visible = True
Image38.Visible = True
Image39.Visible = True
Image40.Visible = True
Image41.Visible = True
Image42.Visible = True
Image43.Visible = True
Image44.Visible = True
Image45.Visible = True
Image46.Visible = True
Image47.Visible = True
Image48.Visible = True
Image49.Visible = True
Image50.Visible = True
Image51.Visible = True
Image52.Visible = True
Image53.Visible = True
Image54.Visible = True
Image55.Visible = True
Image56.Visible = True
Image57.Visible = True
Image58.Visible = True
Image59.Visible = True
Image60.Visible = True
Image61.Visible = True
Image62.Visible = True
Image63.Visible = True
Image64.Visible = True
Image65.Visible = True
Image66.Visible = True
Image67.Visible = True
'MediaPlayer1.Enabled = True
'ediaPlayer1.AutoStart = True
MediaPlayer1.Mute = False
End Sub

Private Sub Command4_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
frmAbout.Show vbModal, Me
End Sub

Private Sub Command5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Picture6.Visible = True
Picture7.Visible = True
End Sub

Private Sub Command6_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Picture6.Visible = False
Picture7.Visible = False
End Sub

Private Sub Form_Load()
Image15.Visible = False
Image16.Visible = False
Image17.Visible = False
Image18.Visible = False
Image19.Visible = False
Image20.Visible = False
Image21.Visible = False
Image22.Visible = False
Image23.Visible = False
Image24.Visible = False
Image25.Visible = False
Image26.Visible = False
Image27.Visible = False
Image28.Visible = False
Image29.Visible = False
Image30.Visible = False
Image31.Visible = False
Image32.Visible = False
Image33.Visible = False
Image34.Visible = False
Image35.Visible = False
Image36.Visible = False
Image37.Visible = False
Image38.Visible = False
Image39.Visible = False
Image40.Visible = False
Image41.Visible = False
Image42.Visible = False
Image43.Visible = False
Image44.Visible = False
Image45.Visible = False
Image46.Visible = False
Image47.Visible = False
Image48.Visible = False
Image49.Visible = False
Image50.Visible = False
Image51.Visible = False
Image52.Visible = False
Image53.Visible = False
Image54.Visible = False
Image55.Visible = False
Image56.Visible = False
Image57.Visible = False
Image58.Visible = False
Image59.Visible = False
Image60.Visible = False
Image61.Visible = False
Image62.Visible = False
Image63.Visible = False
Image64.Visible = False
Image65.Visible = False
Image66.Visible = False
Image67.Visible = False
ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
End Sub

Private Sub Image11_Click()
Image13.Visible = False
End Sub

Private Sub Image12_Click()

Image14.Visible = False
End Sub

Private Sub Image13_Click()
Image14.Visible = True
End Sub

Private Sub Image4_Click()
Command3.Visible = True
End Sub

Private Sub Image5_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\step.wav"), SND_ASYNC)
Command3.Visible = True

End Sub


Private Sub Label1_Click()
Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False

End Sub

Private Sub Label10_Click()
Image9.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label11_Click()
Image6.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label12_Click()
Image7.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label13_Click()
If Label13.Caption = 4 Then
  Label13.Caption = 5
  End If
  If Label13.Caption = 3 Then
 Label13.Caption = 4
 End If
 If Label13.Caption = 2 Then
 Label13.Caption = 3
 End If
  If Label13.Caption = 1 Then
 Label13.Caption = 2
 End If
 If Label13.Caption = 0 Then
 Label13.Caption = 1
 End If
End Sub

Private Sub Label14_Click()
If Label14.Caption = 4 Then
  Label14.Caption = 5
  End If
  If Label14.Caption = 3 Then
 Label14.Caption = 4
 End If
 If Label14.Caption = 2 Then
 Label14.Caption = 3
 End If
  If Label14.Caption = 1 Then
 Label14.Caption = 2
 End If
 If Label14.Caption = 0 Then
 Label14.Caption = 1
 End If
End Sub

Private Sub Label15_Click()
If Label15.Caption = 4 Then
  Label15.Caption = 5
  End If
  If Label15.Caption = 3 Then
 Label15.Caption = 4
 End If
 If Label15.Caption = 2 Then
 Label15.Caption = 3
 End If
  If Label15.Caption = 1 Then
 Label15.Caption = 2
 End If
 If Label15.Caption = 0 Then
 Label15.Caption = 1
 End If
End Sub

Private Sub Label3_Click()
Image5.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label4_Click()
Image10.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label5_Click()
Image4.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Image12.Picture = LoadPicture("C:\windows\desktop\vb projects\war\mecury.bmp")
End Sub

Private Sub Label6_Click()
Image3.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label7_Click()
If Label7.Caption = 4 Then
  Label7.Caption = 5
  End If
  If Label7.Caption = 3 Then
 Label7.Caption = 4
 End If
 If Label7.Caption = 2 Then
 Label7.Caption = 3
 End If
  If Label7.Caption = 1 Then
 Label7.Caption = 2
 End If
 If Label7.Caption = 0 Then
 Label7.Caption = 1
 End If
End Sub

Private Sub Label8_Click()
Image2.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub Label9_Click()
Image8.Visible = True
'Image1.Visible = True
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
End Sub

Private Sub HScroll1_Change()
 Image12.Left = -HScroll1.Value
  ' Image12.Top = -HScroll1.Value
    Image25.Left = -HScroll1.Value
    Image21.Left = -HScroll1.Value
    Image22.Left = -HScroll1.Value
    Image24.Left = -HScroll1.Value
    Image19.Left = -HScroll1.Value
    Image26.Left = -HScroll1.Value
    Image18.Left = -HScroll1.Value
    Image27.Left = -HScroll1.Value
    Image16.Left = -HScroll1.Value
    Image17.Left = -HScroll1.Value
    Image15.Left = -HScroll1.Value
    Image20.Left = -HScroll1.Value
    Image20.Left = -HScroll1.Value
    Image29.Left = -HScroll1.Value
    Image30.Left = -HScroll1.Value
    Image31.Left = -HScroll1.Value
    Image32.Left = -HScroll1.Value
    Image33.Left = -HScroll1.Value
End Sub

Private Sub Label9_DblClick()
form2.money = money + 1000
form2.Power = Power + 20
End Sub

Private Sub Picture2_Click()
form2.money.Caption = form2.money.Caption - 1000
form2.Power.Caption = form2.Power.Caption - 20
'money = money - 1500
'Power = Power - 20
 Dim Counter As Integer
    Dim Workarea(30000) As String
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
    If Label7 >= 1 Then
    Label7 = Label7 - 1
    End If
    Picture2.OLEDragMode = 1
    Timer3.Enabled = True
Call form2.Power1
If form2.money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture3_Click()
form2.money.Caption = form2.money.Caption - 1000
form2.Power.Caption = form2.Power.Caption - 20
'money = money - 1500
'Power = Power - 20
 Dim Counter As Integer
    Dim Workarea(30000) As String
    ProgressBar2.min = LBound(Workarea)
    ProgressBar2.Max = UBound(Workarea)
    ProgressBar2.Visible = True

'Set the Progress's Value to Min.
    ProgressBar2.Value = ProgressBar2.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar2.Value = Counter
    Next Counter
    ProgressBar2.Visible = False
    ProgressBar2.Value = ProgressBar2.min
    If Label13 >= 1 Then
    Label13 = Label13 - 1
    End If
Call form2.Power1
If form2.money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture4_Click()
form2.money.Caption = form2.money.Caption - 1000
form2.Power.Caption = form2.Power.Caption - 20
'money = money - 1500
'Power = Power - 20
Call form2.Power1
If form2.money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture5_Click()
form2.money.Caption = form2.money.Caption - 1000
form2.Power.Caption = form2.Power.Caption - 20
'money = money - 1500
'Power = Power - 20
Call form2.Power1
If form2.money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture6_Click()
form2.money.Caption = form2.money.Caption - 1000
form2.Power.Caption = form2.Power.Caption - 20
'money = money - 1500
'Power = Power - 20
 Dim Counter As Integer
    Dim Workarea(30000) As String
    ProgressBar4.min = LBound(Workarea)
    ProgressBar4.Max = UBound(Workarea)
    ProgressBar4.Visible = True

'Set the Progress's Value to Min.
    ProgressBar4.Value = ProgressBar4.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar4.Value = Counter
    Next Counter
    ProgressBar4.Visible = False
    ProgressBar4.Value = ProgressBar4.min
    If Label15 >= 1 Then
    Label15 = Label15 - 1
    End If
Call form2.Power1
Picture6.OLEDragMode = 1
If form2.money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Picture7_Click()
form2.money.Caption = form2.money.Caption - 1000
form2.Power.Caption = form2.Power.Caption - 20
'money = money - 1500
'Power = Power - 20
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar3.min = LBound(Workarea)
    ProgressBar3.Max = UBound(Workarea)
    ProgressBar3.Visible = True

'Set the Progress's Value to Min.
    ProgressBar3.Value = ProgressBar3.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar3.Value = Counter
    Next Counter
    ProgressBar3.Visible = False
    ProgressBar3.Value = ProgressBar3.min
    If Label14 >= 1 Then
    Label14 = Label14 - 1
    End If
    Picture7.OLEDragMode = 1
    
Call form2.Power1
If form2.money = 6000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 12000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
If form2.money = 24000 Then
i = sndPlaySound(ByVal CStr(App.Path & "\silond1.wav"), SND_ASYNC)
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
Picture2 = LoadPicture("c:\windows\desktop\vb projects\war\viechels\a.bmp")
Picture3 = LoadPicture("c:\windows\desktop\vb projects\war\viechels\furture\4.bmp")
Picture4 = LoadPicture("c:\windows\desktop\vb projects\war\viechels\d.bmp")
Picture7 = LoadPicture("c:\windows\desktop\vb projects\war\viechels\furture\gasp.bmp")
Picture5 = LoadPicture("c:\windows\desktop\vb projects\war\viechels\furture\bunker.bmp")
Picture6 = LoadPicture("c:\windows\desktop\vb projects\war\viechels\sluppy.bmp")
End Sub

Private Sub Timer2_Timer()
Static blnface As Boolean
If blnface = True Then
Image66.Picture = LoadPicture("C:\windows\desktop\vb projects\war\viechels\6.bmp")
blnface = False
Else
Image66.Picture = LoadPicture("C:\windows\desktop\vb projects\war\viechels\7.bmp")
blnface = True
End If

End Sub

Private Sub Timer3_Timer()
Picture2.OLEDragMode = 0
Timer3.Enabled = False
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

'Case Is = "STOP":
'Unload Me
'End
'Case Is = "HELP":
'frmAbout.Show vbModal, Me

End Select

End Sub

Private Sub VScroll1_Change()
    Image12.Top = -VScroll1.Value
    Image25.Top = -VScroll1.Value
    Image21.Top = -VScroll1.Value
    Image22.Top = -VScroll1.Value
    Image24.Top = -VScroll1.Value
    Image19.Top = -VScroll1.Value
    Image26.Top = -VScroll1.Value
    Image18.Top = -VScroll1.Value
    Image27.Top = -VScroll1.Value
    Image16.Top = -VScroll1.Value
    Image17.Top = -VScroll1.Value
    Image15.Top = -VScroll1.Value
    Image20.Top = -VScroll1.Value
    Image28.Top = -VScroll1.Value
    Image29.Top = -VScroll1.Value
    Image30.Top = -VScroll1.Value
    Image31.Top = -VScroll1.Value
    Image32.Top = -VScroll1.Value
      Image33.Top = -VScroll1.Value
      Image34.Top = -VScroll1.Value
End Sub

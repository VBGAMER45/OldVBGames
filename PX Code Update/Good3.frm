VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Project X Online http://www.visualbasiczone.com Jonathan Valentin 2001-2003"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Good3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   549
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer TmrSnow 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   720
   End
   Begin VB.CommandButton CmdMenu 
      Caption         =   "Menu"
      Height          =   615
      Left            =   10560
      TabIndex        =   74
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   10560
      TabIndex        =   72
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton CmdQuests 
      Caption         =   "Quests"
      Height          =   495
      Left            =   10560
      TabIndex        =   73
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdScan 
      Caption         =   "Scan"
      Height          =   495
      Left            =   10560
      TabIndex        =   77
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton CmdStealth 
      Caption         =   "Stealth"
      Height          =   495
      Left            =   10560
      TabIndex        =   78
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton CmdAttack 
      Caption         =   "Attack"
      Height          =   495
      Left            =   10560
      TabIndex        =   76
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton CmdGuard 
      Caption         =   "Guard"
      Height          =   735
      Left            =   10560
      TabIndex        =   75
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton CmdItem 
      Caption         =   "Item Change"
      Height          =   615
      Left            =   10560
      TabIndex        =   71
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CmdshowStats 
      Caption         =   "Stats"
      Height          =   495
      Left            =   10560
      TabIndex        =   70
      Top             =   2040
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   1560
      OleObjectBlob   =   "Good3.frx":030A
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   240
      Left            =   3840
      TabIndex        =   69
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicTileType 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   68
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play Orcania"
      Height          =   375
      Left            =   10320
      TabIndex        =   67
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdCenter 
      Caption         =   "&Center "
      Height          =   975
      Left            =   10560
      TabIndex        =   65
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame FrameUsers 
      BackColor       =   &H00000000&
      Caption         =   "Current Users"
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   2760
      TabIndex        =   62
      Top             =   1440
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton CmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   720
         TabIndex        =   64
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ListBox lstConnections 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         ItemData        =   "Good3.frx":050C
         Left            =   120
         List            =   "Good3.frx":050E
         MultiSelect     =   2  'Extended
         TabIndex        =   63
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Timer TimerDraw 
      Enabled         =   0   'False
      Interval        =   2451
      Left            =   360
      Top             =   720
   End
   Begin VB.Frame Light2 
      Caption         =   "Lightning Bolt Color"
      Height          =   975
      Left            =   120
      TabIndex        =   53
      Top             =   3000
      Visible         =   0   'False
      Width           =   3975
      Begin VB.OptionButton OptionColor 
         Caption         =   "B"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   61
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "G"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   60
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "R"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "Y"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "T"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   57
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "P"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   56
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "Grey Scale Light"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   55
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "Grey Scale Dark"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   54
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Light1 
      Caption         =   "Lightning Bolt Stats"
      Height          =   2415
      Left            =   120
      TabIndex        =   39
      Top             =   4080
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CheckBox ChkStart 
         Caption         =   "Always Start Same Location"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   3855
      End
      Begin VB.HScrollBar Velocity 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   20
         Min             =   2
         TabIndex        =   46
         Top             =   960
         Value           =   15
         Width           =   2775
      End
      Begin VB.HScrollBar Velocity 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   45
         Top             =   720
         Value           =   5
         Width           =   2775
      End
      Begin VB.HScrollBar HeightNum 
         Height          =   255
         Left            =   120
         Max             =   200
         TabIndex        =   44
         Top             =   480
         Value           =   190
         Width           =   2775
      End
      Begin VB.HScrollBar Speed 
         Height          =   255
         Left            =   120
         Max             =   299
         TabIndex        =   43
         Top             =   1200
         Value           =   149
         Width           =   2775
      End
      Begin VB.HScrollBar NickNum 
         Height          =   255
         Left            =   120
         Max             =   30
         TabIndex        =   42
         Top             =   240
         Value           =   20
         Width           =   2775
      End
      Begin VB.CheckBox ChkCloud 
         Caption         =   "Show Cloud"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox ChkSplit 
         Caption         =   "Split Bolt"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Label lblVel 
         Caption         =   "V Max: 15"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   52
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblVel 
         Caption         =   "V Min: 5"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   51
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblspeed 
         Caption         =   "Speed: 150"
         Height          =   255
         Left            =   3000
         TabIndex        =   50
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblNicks 
         Caption         =   "Nicks: 20"
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height: 190"
         Height          =   255
         Left            =   3000
         TabIndex        =   48
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.PictureBox picSwan 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3480
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer MonsterAttackTimer 
      Interval        =   1900
      Left            =   840
      Top             =   1200
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   32
      Left            =   1440
      Max             =   255
      Min             =   -255
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Frame FrameStats 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   4800
      TabIndex        =   30
      Top             =   0
      Width           =   3135
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   0
         ScaleHeight     =   117
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   217
         TabIndex        =   31
         Top             =   -360
         Width           =   3255
         Begin VB.CommandButton CmdStats 
            Caption         =   "X"
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   360
            Width           =   375
         End
         Begin VB.Timer StatTimer 
            Interval        =   100
            Left            =   1800
            Top             =   960
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "STA:"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "HP:"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   465
         End
         Begin VB.Shape MANShp 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   150
            Left            =   720
            Top             =   1080
            Width           =   2250
         End
         Begin VB.Shape STAShp 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   150
            Left            =   720
            Top             =   840
            Width           =   2250
         End
         Begin VB.Shape HPShp 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   150
            Left            =   720
            Top             =   600
            Width           =   2250
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Mana:"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   585
         End
         Begin VB.Image BackPic 
            Height          =   1215
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   3255
         End
      End
   End
   Begin VB.PictureBox PicTXT 
      BackColor       =   &H00404040&
      Height          =   2895
      Left            =   600
      ScaleHeight     =   2835
      ScaleWidth      =   6975
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   7030
      Begin VB.VScrollBar TextScroll 
         Height          =   2830
         LargeChange     =   3
         Left            =   6720
         Max             =   0
         TabIndex        =   29
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Holder 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   6735
         TabIndex        =   26
         Top             =   1200
         Width           =   6735
         Begin VB.Label Choice 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   ">>>"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   6450
            WordWrap        =   -1  'True
         End
         Begin VB.Label Pitanje 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   ":::"
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   6420
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Timer TimerAttacklbl 
      Interval        =   1000
      Left            =   840
      Top             =   1680
   End
   Begin VB.CommandButton MakeAFolder 
      Caption         =   "Make A Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   23
      Top             =   15240
      Width           =   1095
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   9840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9840
      TabIndex        =   18
      Text            =   "jonathanv"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9720
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10080
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer tmrSendData 
      Interval        =   1
      Left            =   360
      Top             =   1680
   End
   Begin VB.TextBox txtDialog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6960
      Width           =   8175
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   12
      Top             =   7920
      Width           =   7095
   End
   Begin VB.Timer Information 
      Interval        =   1
      Left            =   360
      Top             =   2160
   End
   Begin VB.Timer ReloadBad 
      Interval        =   9095
      Left            =   360
      Top             =   4440
   End
   Begin VB.Timer Events 
      Interval        =   10
      Left            =   360
      Top             =   1200
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   1
      TabIndex        =   2
      Top             =   6720
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6915
      Left            =   7920
      Max             =   1
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Tiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1980
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   360
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckConnection 
      Index           =   0
      Left            =   840
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblMoney 
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
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   66
      Top             =   6840
      Width           =   3495
   End
   Begin VB.Image ImgTriForce 
      Height          =   750
      Left            =   10560
      Picture         =   "Good3.frx":0510
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image PicItem 
      Height          =   735
      Index           =   0
      Left            =   19500
      Top             =   720
      Width           =   735
   End
   Begin VB.Image PICnpc 
      Height          =   375
      Index           =   0
      Left            =   2760
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape AttackShape 
      BorderColor     =   &H0000FF00&
      Height          =   6930
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Label labBrightness 
      Caption         =   "Brightness-level: 0"
      Height          =   255
      Left            =   1560
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Meter 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image Title 
      Height          =   660
      Left            =   1200
      Picture         =   "Good3.frx":2586
      Top             =   2280
      Width           =   5790
   End
   Begin VB.Label lblAttack 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack Label"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Shape PlayerSquare 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   480
      Left            =   4320
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label MapName 
      BackStyle       =   0  'Transparent
      Caption         =   "newbie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10200
      TabIndex        =   21
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblHost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   10080
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblplayername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblplayername"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   1
      Left            =   10080
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label lblplayername 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   6960
      TabIndex        =   10
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label Lbltop 
      BackStyle       =   0  'Transparent
      Caption         =   "Lbltop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   9960
      TabIndex        =   9
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "LblLeft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9960
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image p21 
      Height          =   480
      Index           =   1
      Left            =   9960
      Picture         =   "Good3.frx":2EC1
      Top             =   19080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTileKind 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TargetShape 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FF0000&
      Height          =   480
      Left            =   10080
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label NowALKDIRECTION 
      BackStyle       =   0  'Transparent
      Caption         =   "NowALKDIRECTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   9240
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Direction 
      BackStyle       =   0  'Transparent
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10080
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   10320
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
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
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
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
      Volume          =   -1170
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   255
      Left            =   10320
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
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
      Filename        =   ""
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
      Volume          =   -530
      WindowlessVideo =   0   'False
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   480
      Left            =   2760
      Top             =   1320
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      FillColor       =   &H00FF0000&
      Height          =   480
      Left            =   2040
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   3720
      Top             =   3360
      Width           =   240
   End
   Begin VB.Image SpellPicture 
      Height          =   1095
      Index           =   0
      Left            =   3840
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Save Game"
      End
      Begin VB.Menu mnuloadgame 
         Caption         =   "Load Game"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnumusic 
         Caption         =   "Music"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusound 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnucredits 
         Caption         =   "Credits"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnukeys 
         Caption         =   "Keys"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuattackmenu 
      Caption         =   "AttackMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuattack 
         Caption         =   "Attack"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuheal 
         Caption         =   "Heal"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnurevive 
         Caption         =   "Revive"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnumagicattack 
         Caption         =   "Magic Attack"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnufollow 
         Caption         =   "Follow"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuactionmenu 
      Caption         =   "ActionMenu"
      Visible         =   0   'False
      Begin VB.Menu mnucuttree 
         Caption         =   "CutDownTree"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnufish 
         Caption         =   "Fish"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnumine 
         Caption         =   "Mine for gold"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnufarm 
         Caption         =   "Plant Seeds"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnudrink 
         Caption         =   "Drink"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuharvest 
         Caption         =   "Harvest"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnupickup 
         Caption         =   "PickUp"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuUse 
         Caption         =   "Use Item on Item"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuseperator 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuextratile 
         Caption         =   "Extra Tile Stuff"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Project X Online the open source project
'Jonathan Valentin 2001 - 2003
'vbbestgame3@hotmail.com
'Game Website at http://x.incyberspace.net
'http://www.visualbasiczone.com
'http://www.projectxonline.net

'To make the game run smoother convert the scirpt events to map files
'and make a sub to read the data

Dim TargetedNpc As Integer
Dim TargetedMonster As Integer
Dim TargetedPlayer As Integer
Dim TargetedItem As Integer
Dim TalkOnce As Integer
Dim FirstTile As Boolean
Dim LightsOn As Boolean
Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hbitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hbitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'the following declares are for cool graphic functions
'Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function FoxTime Lib "FoxCBmp" (Optional ByVal Flags As Long) As Long
Private Declare Function FoxTimer Lib "FoxCBmp" (ByVal Time As Long) As Long
Private Declare Function FoxCounter Lib "FoxCBmp" (Optional ByVal Flags As Long) As Long
Private Declare Function FoxBrightness Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstBmp As Long, ByVal SrcDC As Long, ByVal SrcBmp As Long, ByVal TransColor As Long, ByVal Brightness As Long, Optional ByVal Flags As Long) As Long
Private Declare Function FoxBlendIn Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal SrcDC As Long, ByVal SrcBmp As Long, ByVal TransColor As Long, ByVal Alpha As Byte, Optional ByVal Flags As Long) As Long
Private Declare Function FoxRotate Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal SrcDC As Long, ByVal SrcBmp As Long, ByVal TransColor As Long, ByVal Angle As Double, Optional ByVal Flags As Long) As Long
Private Declare Function FoxBlur Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstBmp As Long, ByVal SrcDC As Long, ByVal SrcBmp As Long, ByVal TransColor As Long, ByVal Blur As Long, Optional ByVal Flags As Long) As Long
'Private Declare Function FoxAlphaBlend Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Alpha As Byte, ByVal TransColor As Long, ByVal Flags As Long) As Long
Private Declare Function FoxAlphaMask Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal MaskDC As Long, ByVal MaskX As Long, ByVal MaskY As Long, ByVal TransColor As Long, ByVal Flags As Long) As Long
Private Declare Function FoxBumpMap Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal TransColor As Long, ByVal Flags As Long) As Long
Private Declare Function FoxMosaic Lib "FoxCBmp.dll" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Level As Long, ByVal TransColor As Long, ByVal Flags As Long) As Long
Dim Slow As Integer
Dim Slow2 As Integer
Dim Team As String
Dim FlagName As String
Dim Linkx As String
Dim linky As String
Dim response As String
Dim wait2 As Integer


Public WithEvents P51  As VB.Image
Attribute P51.VB_VarHelpID = -1
Public WithEvents LblPName As VB.Label
Attribute LblPName.VB_VarHelpID = -1
'Light and Day
Dim TimeofDay3 As Integer 'time of day
Dim SwitchToDay As Boolean

Dim mousex As Integer
Dim mousey As Integer

'Sometimes parameters are sent along with the commands that are sent between computers.
'All parameters will be formatted to be exactly PARAM_LEN characters long to simplify the parsing of commands by the receiving computer(s).
Const PARAM_LEN = 10

'This is a collection of commands and data to be sent to other computers, either the server (if you have connected to one) or to all connected computers (if you are the server).
Dim SendList As New Collection
'This is a collection of commands and data that specifies where to send the items in SendList.
'Each item in SendList has an associated item in SendTo which says to which computer the information in SendList is to be sent.
Dim SendTo As New Collection

'These are used in tracking where your mouse is when drawing pictures.
Dim iX As Integer, iY As Integer

'When this flag is set and the form is unloaded, the confirm quit message will not appear.
'Read more about why this is used in the Form_Unload event.
Dim bJustQuit As Boolean
'below is the lightning stuff
Private Type LightningBolt
    Center As Long
    Inner1 As Long
    Inner2 As Long
    Inner3 As Long
    Inner4 As Long
    Outer1 As Long
    Outer2 As Long
    Outer3 As Long
    Outer4 As Long
    
    Nicks As Long
    VelocityMin As Long
    VelocityMax As Long
    Height As Long
    
    SplitBolt As Boolean
    SameStart As Boolean
    ShowCloud As Boolean
End Type
Dim Bolt As LightningBolt
Dim SendChatTmr As Integer
Dim PlayerMoved As Boolean

Private Type snowtype
X As Integer
y As Integer
End Type
Dim Snow(10000) As snowtype
Private Sub ChkCloud_Click()
If ChkCloud.Value = 0 Then
    Bolt.ShowCloud = False
Else
    Bolt.ShowCloud = True
End If
End Sub

Private Sub ChkSplit_Click()
If ChkSplit.Value = 0 Then
    Bolt.SplitBolt = False
Else
    Bolt.SplitBolt = True
End If
End Sub

Private Sub ChkStart_Click()
If ChkStart.Value = 0 Then
    Bolt.SameStart = False
Else
    Bolt.SameStart = True
End If
End Sub

Public Function DrawBolt(Pic2 As Form)
Dim SM, I, LX, LY, SX, SY, LX2, LY2, SX2, SY2 As Integer
Dim Alter As Boolean
'Pic2.Cls
If Bolt.SameStart = True Then
    SM = Pic2.ScaleWidth / 2
Else
    SM = Int(Rnd * (Pic2.ScaleWidth / 2) + Pic2.ScaleWidth / 4)
End If
Pic2.Forecolor = Bolt.Center
SX = SM
If Bolt.ShowCloud = True Then
    SY = 10
Else
    SY = 0
    SetPixel Pic2.hDC, SM, 0, Bolt.Center
End If
I = Int(Rnd * 2)
If I = 0 Then
    Alter = False
    LX = SX + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
Else
    Alter = True
    LX = SX - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
End If
LY = Bolt.Height / Bolt.Nicks
For I = 0 To Bolt.Nicks
    Pic2.Forecolor = Bolt.Center
    Pic2.Line (SX, SY)-(LX, LY)
    Pic2.Forecolor = Bolt.Inner1
    Pic2.Line (SX - 1, SY)-((SX + LX) / 2 - 1, (SY + LY) / 2)
    Pic2.Line (SX + 1, SY)-((SX + LX) / 2 + 1, (SY + LY) / 2)
    Pic2.Forecolor = Bolt.Inner2
    Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX - 1, LY)
    Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX + 1, LY)
    Pic2.Forecolor = Bolt.Outer1
    Pic2.Line (SX - 2, SY)-((SX + LX) / 2 - 2, (SY + LY) / 2)
    Pic2.Line (SX + 2, SY)-((SX + LX) / 2 + 2, (SY + LY) / 2)
    Pic2.Forecolor = Bolt.Outer2
    Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX - 2, LY)
    Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX + 2, LY)
    If I >= Round((Bolt.Nicks / 2), 0) And Bolt.SplitBolt = True Then
        Pic2.Forecolor = Bolt.Center
        Pic2.Line (SX, SY)-(LX, LY)
        Pic2.Forecolor = Bolt.Inner1
        Pic2.Line (SX2 - 1, SY2)-((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)
        Pic2.Line (SX2 + 1, SY2)-((SX2 + LX2) / 2 + 1, (SY2 + LY2) / 2)
        Pic2.Forecolor = Bolt.Inner2
        Pic2.Line ((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)-(LX2 - 1, LY2)
        Pic2.Line ((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)-(LX2 + 1, LY2)
        Pic2.Forecolor = Bolt.Outer1
        Pic2.Line (SX2 - 2, SY2)-((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)
        Pic2.Line (SX2 + 2, SY2)-((SX2 + LX2) / 2 + 2, (SY2 + LY2) / 2)
        Pic2.Forecolor = Bolt.Outer2
        Pic2.Line ((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)-(LX2 - 2, LY2)
        Pic2.Line ((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)-(LX2 + 2, LY2)
    End If
    If I = Bolt.Nicks Then
        'Do Tail
        If Alter = True Then
            '-
            SX = LX
            SY = LY
            LX = LX - 1
            LY = LY + 2
        Else
            '+
            SX = LX
            SY = LY
            LX = LX + 1
            LY = LY + 2
        End If
        Pic2.Forecolor = Bolt.Inner1
        Pic2.Line (SX, SY)-(LX, LY)
        Pic2.Forecolor = Bolt.Inner3
        Pic2.Line (SX - 1, SY)-((SX + LX) / 2 - 1, (SY + LY) / 2)
        Pic2.Line (SX + 1, SY)-((SX + LX) / 2 + 1, (SY + LY) / 2)
        Pic2.Forecolor = Bolt.Inner4
        Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX - 1, LY)
        Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX + 1, LY)
        Pic2.Forecolor = Bolt.Outer3
        Pic2.Line (SX - 2, SY)-((SX + LX) / 2 - 2, (SY + LY) / 2)
        Pic2.Line (SX + 2, SY)-((SX + LX) / 2 + 2, (SY + LY) / 2)
        Pic2.Forecolor = Bolt.Outer4
        Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX - 2, LY)
        Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX + 2, LY)
        GoTo RefreshPic
    End If
    SX = LX
    SY = LY
    Alter = Int(Rnd * 2)
    If Alter = True Then
        Alter = False
        LX = LX + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
    Else
        Alter = True
        LX = LX - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
    End If
    LY = LY + Bolt.Height / Bolt.Nicks
    If I < Round((Bolt.Nicks / 2), 0) - 1 And Bolt.SplitBolt = True Then
        LX2 = SX
        LY2 = SY
    End If
    If I >= Round((Bolt.Nicks / 2), 0) - 1 And Bolt.SplitBolt = True Then
        SX2 = LX2
        SY2 = LY2
        Alter = Int(Rnd * 2)
        If Alter = True Then
            Alter = False
            LX2 = LX2 + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
        Else
            Alter = True
            LX2 = LX2 - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
        End If
        LY2 = LY2 + Bolt.Height / Bolt.Nicks
    End If
Next I
RefreshPic:

Pic2.Refresh
End Function
Public Sub ProcessData(vsString As String, viConnection As Integer)
'This procedure processes data received from either the server or from connections to the server.
'vsString = the command string being processed
'viConnection = the connection from which the command string was received

Dim I As Integer
Dim sCommand As String
Dim sInstruction As String
Dim sData As String
Dim sFormatStr As String

For I = 1 To PARAM_LEN
    sFormatStr = sFormatStr & "0"
Next I

'Separate commands may be received together so each command is followed by a carriage return.
'So as long as a carriage return is found in the data stream, there must be a command in it so continue processing data.
Do While InStr(1, vsString, vbCrLf)
    
    'Store in sCommand the part of the data stream that contains the first command.
    sCommand = Mid(vsString, 1, InStr(1, vsString, vbCrLf) - 1)
    
    'Each command contains an instruction such as [Message] or [Disconnect].
    'Some commands also contain parameters.
    'Here the instruction part of the command is stored in sInstruction and the rest is stored in sData.
    sInstruction = Mid(sCommand, 1, InStr(1, sCommand, "]"))
    sData = Mid(sCommand, InStr(1, sCommand, "]") + 2, Len(sCommand))
    
    'txtDialog.Text = txtDialog.Text & vsString & vbCrLf
    
    'Branch depending upon the instruction.
    Select Case sInstruction
        Case "[Change Name]"
            'This command is sent by a connecting user when they change their name in their Name text box.  (Only the server will receive such a command.)
            
            'Update their name in the name list.
            lstConnections.List(viConnection) = sData
            'Refresh the name list on all connected computers.
            SendPersonName viConnection
    
         
        Case "[Disconnect]"
            'This is a request to be disconnected from the server.  People cannot disconnect themselves from the server - the server must disconnect them.
            
            'Close the connection.
            sckConnection(viConnection).Close
            'Call the procedure that ensures that one winsock control is listening for connections.
            MakeListen
            'Update the status.
            txtStatus.Text = txtStatus.Text & vbCrLf & lstConnections.List(viConnection) & " disconnected."
            'Reset their name in the name list.
            lstConnections.List(viConnection) = "[Open]"
            'Have the server notify all connected computer that this person has disconnected.
            For I = 0 To giMaxConnections - 1
                If sckConnection(I).State = sckConnected Then
                    SendList.Add "[Disconnected] " & Format(viConnection, sFormatStr)
                    SendTo.Add "sckConnection(" & I & ")"
                End If
            Next I
        Case "[Disconnected]"
            'This command is received when the server notifies someone that someone else has disconnected.
            txtDialog.Text = txtDialog.Text & "disconnected" & vbCrLf
            MsgBox "You have been disconnted!"
            'Update the status.
            txtStatus.Text = txtStatus.Text & vbCrLf & lstConnections.List(sParam(sData, 1)) & " disconnected."
            'Reset their name in the name list.
            lstConnections.List(sParam(sData, 1)) = "[Open]"
            txtDialog.Text = txtDialog.Text & lstConnections.List(sParam(sData, 1)) & " disconnected."
        Case "[Host]"
            'This command is received from the server when one connects to it.  It notifies connecting people of the host's name.
            Dim j As String
            'Update the name of the host.
            lblHost.Caption = sData
            'Update the status.
            txtStatus.Text = txtStatus.Text & vbCrLf & "You have connected to " & sData & "."
            j = "[p]" & "," & MainPlayer.PlayerName & "," & MainPlayer.AccountName
            Form1.SendClientData "p", j
        Case "[Host Change Name]"
            'This command is received from the server when the host changes their name.
            
            'Update the name of the server.
            lblHost.Caption = sData
        Case "[Joined]"
            'This command is sent to the server when someone joins, notifying the server of the name of the person connecting.
            
            'Update the status.
            txtDialog.Text = txtDialog.Text & sData & " has entered the world!." & vbCrLf
            txtStatus.Text = txtStatus.Text & vbCrLf & sData & " joined."
            MediaPlayer2.Filename = App.Path & "\beep.wav"
            
            Monsters.CreatePlayer "newbie", 55, 29, sData
            
            'If you are the server then notify all other connections that someone has joined and send the name of the new connection.
            For I = 0 To giMaxConnections - 1
                If sckConnection(I).State = sckConnected And I <> viConnection Then
                    SendList.Add "[Joined] " & sData
                    SendTo.Add "sckConnection(" & I & ")"
                End If
            Next I
            'Also if you are the server...
            If lstConnections.ListCount > 0 And sckConnect.State <> sckConnected Then
                'Update the name in the connections list.
                lstConnections.List(viConnection) = sData
                'Send the name list to the person connecting.
                SendPeopleList viConnection
                'Notify all open connections of the new person who just joined.
                SendPersonName viConnection
            End If
        Case "[Kicked]"
            'This command is sent by the server notifying connections that someone was kicked.
            
            'Update the status.
            txtStatus.Text = txtStatus.Text & vbCrLf & "Kicked " & lstConnections.List(sParam(sData, 1)) & "."
            
            ' & lstConnections.List(sParam(sData, 1)) & "." & vbCrLf
            MediaPlayer2.Filename = App.Path & "\beep.wav"

            'Reset their name in the name list.
            lstConnections.List(sParam(sData, 1)) = "[Open]"
            txtDialog.Text = txtDialog.Text & "You are kicked off" & vbCrLf
            
       
     
        Case "[Message]"
            'This command is sent when someone enters a message.
        
            'Show the message.
            
            txtDialog.Text = txtDialog.Text & sData & vbNewLine
            MediaPlayer2.Filename = App.Path & "\beep.wav"

            'Notify all open connections of the message.
            For I = 0 To giMaxConnections - 1
                If sckConnection(I).State = sckConnected And I <> viConnection Then
                    SendList.Add "[Message] " & sData
                    SendTo.Add "sckConnection(" & I & ")"
                End If
            Next I
        Case "[People]"
           Dim length As Integer
           Dim slength As Integer
            'This is sent by the server to notify open connections of name changes.
            
            'Update the name list.
            lstConnections.List(sParam(sData, 1)) = sLongParam(sData, 2)
           ' Form1.LoadOtherPlayers
           length = Len(MainPlayer.PlayerName)
           slength = Len(sData)
            If Right(sData, length) = MainPlayer.PlayerName Then
             j = "[p]" & "," & MainPlayer.PlayerName & "," & MainPlayer.AccountName
            Form1.SendClientData "p", j
            End If
            
        Case "[KillPlayer]"
        If MainPlayer.PlayerName = sData Then
        
        MainPlayer.Health = 0
        Call Monsters.CheckDead
        
        End If
        
    End Select
    Call Server.ServerCommands(sInstruction, "", sData)
    
    
    'Remove the processed command from the data stream.
    vsString = Mid(vsString, InStr(1, vsString, vbCrLf) + 2, Len(vsString))
Loop
End Sub


Private Sub CmdAttack_Click()
If MainPlayer.Stamina <= 0 Then
MainPlayer.Stamina = 0
CmdAttack.Enabled = False
Else
CmdAttack.Enabled = True

End If
If MainPlayer.Stamina < 0 Then
MainPlayer.Stamina = 0
Call TOOLS.LoadStats

End If
Monsters.attackMonster
Monsters.DamageforATTack
If OnlineGame = True Then
Server.AttackPlayer
End If
End Sub

Private Sub CmdCenter_Click()
Call Stuff.CenterPerson
Form1.lblPlayerName(0).Top = Form1.Image1.Top - 20
Form1.lblPlayerName(0).Left = Form1.Image1.Left

End Sub

Private Sub cmdClose_Click()
FrameUsers.Visible = False
End Sub


Private Sub CmdGuard_Click()
If CmdGuard.FontBold = True Then
TOOLS.GuardTimer.Enabled = False
CmdGuard.FontBold = False
CmdGuard.Refresh
Exit Sub
End If
If CmdGuard.FontBold = False Then
Monsters.DamageforATTack
TOOLS.GuardTimer.Enabled = True
CmdGuard.FontBold = True
CmdGuard.Refresh
End If
End Sub

Private Sub CmdHelp_Click()
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
TOOLS.Frame2.Visible = True
TOOLS.Frame13.Visible = False
TOOLS.Frame4.Visible = False
TOOLS.Frame1.Visible = False
TOOLS.Frame3.Visible = False
TOOLS.Frame6.Visible = False
Call ZeldaMod
End Sub

Private Sub CmdItem_Click()
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
TOOLS.Frame4.Visible = True
TOOLS.Frame13.Visible = False
TOOLS.Frame2.Visible = False
TOOLS.Frame5.Visible = False
TOOLS.Frame1.Visible = False
TOOLS.Frame3.Visible = False
TOOLS.Frame6.Visible = False
Call ZeldaMod
End Sub

Private Sub CmdMenu_Click()

Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
TOOLS.Frame3.Visible = True
TOOLS.Frame13.Visible = False
TOOLS.Frame4.Visible = False
TOOLS.Frame1.Visible = False
TOOLS.Frame2.Visible = False
TOOLS.Frame6.Visible = False
Call ZeldaMod

End Sub
Private Sub ZeldaMod()
TOOLS.CmdCFull.Visible = True
TOOLS.CmdScan.Visible = False
TOOLS.CmdStealth.Visible = False
TOOLS.CmdQuests.Visible = False
TOOLS.Left = Form1.Width \ 2 - TOOLS.Width \ 2
TOOLS.Height = 8100
TOOLS.CmdAttack.Visible = False
TOOLS.CmdGuard.Visible = False
TOOLS.Width = 3960
TOOLS.Show

End Sub

Private Sub CmdQuests_Click()
TOOLS.Frame13.Visible = True
TOOLS.Frame2.Visible = False
TOOLS.Frame4.Visible = False
TOOLS.Frame1.Visible = False
TOOLS.Frame3.Visible = False
TOOLS.Frame6.Visible = False
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
Call ZeldaMod
End Sub

Private Sub CmdScan_Click()
Monsters.ScanForMonsters
End Sub

Private Sub CmdshowStats_Click()
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
TOOLS.Frame1.Visible = True
TOOLS.Frame13.Visible = False
TOOLS.Frame4.Visible = False
TOOLS.Frame2.Visible = False
TOOLS.Frame3.Visible = False
TOOLS.Frame6.Visible = False
Call ZeldaMod
End Sub

Private Sub CmdStats_Click()
FrameStats.Visible = False
End Sub



Private Sub CmdStealth_Click()
If CmdStealth.FontBold = True Then
    CmdStealth.FontBold = False
    MainPlayer.Stealth = False
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You are no longer sneaking" & vbNewLine
    CmdStealth.Refresh
Exit Sub
End If
If CmdStealth.FontBold = False Then
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You are sneaking" & vbNewLine
    MainPlayer.Stealth = True
    CmdStealth.FontBold = True
    CmdStealth.Refresh
End If
End Sub

Private Sub Command1_Click()

Form1.FontSize = 16
Form1.FontBold = True

Form1.Caption = "The Game is loading Please wait!"
CurrentX = 100
CurrentY = 400
Form1.Print "The Game is loading Please wait!"
Form1.Refresh
imagefx.GrayScale Tiles

'imagefx.Invert Tiles
RedrawMap

Form1.Caption = MainPlayer.PlayerName & " Project X  Online Jonathan Valentin 2001-2002 Verion 0.38"


End Sub

Private Sub Command2_Click()
Call PlayOrcania
End Sub
Private Sub PlayOrcania()
'For Zelda Mode warp to templeoflight
Tiles.Picture = LoadPicture(App.Path & "\tiles2.bmp")
Close
Open App.Path & "\maps\templeoflight.map" For Input As #1: Close
Open App.Path & "\maps\templeoflight.map" For Binary Access Read Lock Read As #1
EraseAll
Get #1, , Map
Close #1
HScroll1.Value = 0
VScroll1.Value = 8
Image1.Left = 240
Image1.Top = 240
MapName.Caption = "templeoflight.map"
MainPlayer.MapName = "templeoflight.map"
RedrawMap

End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'if you do not like i j k l keys
'just change them to vbkeyup vbkeydown vbkeyleft and vbkeyright
If MapEd = True Then 'If in map editor ignore the rest
    Exit Sub
End If
      
        Select Case KeyCode
        
        Case vbKeySpace
            Call RPG.Opendoors
            
        Case vbKeyT
            TalkOnce = 0
            Call Talk
            
        Case vbKeyQ
            TalkOnce = TalkOnce + 1
            If MainPlayer.MapName = "newbie" And TalkOnce = 2 Then
                If MainPlayer.xpos = 41 And MainPlayer.ypos = 40 Or MainPlayer.xpos = 40 And MainPlayer.ypos = 40 Then
                    Call RPG.QuestGenerator
                    TalkOnce = 0
                End If
            End If
            
        Case vbKeyU
            Call ItemInfo.UseSpell
        Case vbKey0
            Call ItemInfo.UseItem
        Case vbKeyEscape
            Unload Me
        Case vbKeyI
            Call MovePlayerUp
            PlayerMoved = True
            Call AdjustItemsandMonsters
        Case vbKeyK
            Call MovePlayerDown
            PlayerMoved = True
            Call AdjustItemsandMonsters
        Case vbKeyJ
            Call MovePlayerLeft
            PlayerMoved = True
            Call AdjustItemsandMonsters
        Case vbKeyL
            Call MovePlayerRight
            PlayerMoved = True
            Call AdjustItemsandMonsters
        Case Else
            PlayerMoved = False
        
        End Select

End Sub
Private Sub MovePlayerRight()

If ZeldaMode = True Then Image1.Picture = teleporter.ZRight.Picture
    Direction.Caption = "right"
    
    Shape2.Left = Snap(Image1.Left + Image1.Width, 32)
    Shape2.Top = Snap(Image1.Top, 32)

    If Map(Linkx + 1, linky).Tiletype = 0 Or Map(Linkx + 1, linky).Tiletype = Tree Or Map(Linkx + 1, linky).Tiletype = Gold Or Map(Linkx + 1, linky).Tiletype = Sign Then
        NowALKDIRECTION.Caption = "Right"
    End If

If NowALKDIRECTION.Caption = "Right" Then
Exit Sub
Else
NowALKDIRECTION.Caption = ""

If Image1.Left < 520 Then
Image1.Left = Image1.Left + 16


Else
If HScroll1.Value < HScroll1.Max Then

 MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
 MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll
If Map(MainPlayer.xpos + 1, MainPlayer.ypos).Tiletype = 0 Or Map(MainPlayer.xpos + 1, MainPlayer.ypos).Tiletype = Tree Or Map(MainPlayer.xpos + 1, MainPlayer.ypos).Tiletype = Sign Then
Else
HScroll1.Value = HScroll1.Value + 1

End If
Else
'If at the edge of a map load the next map
Call Form1.linker
End If
End If
End If



End Sub
Private Sub MovePlayerLeft()

If ZeldaMode = True Then Image1.Picture = teleporter.ZLeft.Picture
Direction.Caption = "left"

Shape2.Left = Snap(Image1.Left - Image1.Width, 32): Shape2.Top = Snap(Image1.Top, 32)
    If Linkx = 0 Then
    Else
    Linkx = Linkx - 1
    End If
If Map(Linkx + 1, linky).Tiletype = NONWALKABLE Or Map(Linkx, linky).Tiletype = Tree Or Map(Linkx, linky).Tiletype = Gold Or Map(Linkx, linky).Tiletype = Sign Then
NowALKDIRECTION.Caption = "Left"
End If

 
If NowALKDIRECTION.Caption = "Left" Then
Exit Sub
Else

NowALKDIRECTION.Caption = ""

If Image1.Left > 10 Then

Image1.Left = Image1.Left - 16

Else

If HScroll1.Value >= 1 Then

 MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
 MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll
If Map(MainPlayer.xpos - 1, MainPlayer.ypos).Tiletype = NONWALKABLE Or Map(MainPlayer.xpos - 1, MainPlayer.ypos).Tiletype = Tree Or Map(MainPlayer.xpos - 1, MainPlayer.ypos).Tiletype = Sign Then
Else
HScroll1.Value = HScroll1.Value - 1

End If
Else
'If at the edge of a map load the next map
Call Form1.linker

End If
End If
End If

End Sub

Private Sub MovePlayerDown()

If ZeldaMode = True Then Image1.Picture = teleporter.ZDown.Picture
    Direction.Caption = "down"

    'Shape2.Left = Snap(Label2.Caption, 32): Shape2.Top = Snap(Label1.Caption + Image1.Height, 32)  '+ 22, 32)
    Shape2.Left = Snap(Image1.Left, 32): Shape2.Top = Snap(Image1.Top + Image1.Height, 32)    '+ 22, 32)

    If Map(Linkx, linky + 1).Tiletype = NONWALKABLE Or Map(Linkx, linky + 1).Tiletype = Tree Or Map(Linkx, linky + 1).Tiletype = Gold Or Map(Linkx, linky + 1).Tiletype = Sign Then
        NowALKDIRECTION.Caption = "Down"
    End If

    If NowALKDIRECTION.Caption = "Down" Then

    Else
        NowALKDIRECTION.Caption = ""
            If Image1.Top < 416 Then
                Image1.Top = Image1.Top + 16
           
            Else
                If VScroll1.Value <= 47 Then

                    MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
                    MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll
                If Map(MainPlayer.xpos, MainPlayer.ypos + 1).Tiletype = NONWALKABLE Or Map(MainPlayer.xpos, MainPlayer.ypos + 1).Tiletype = Tree Or Map(MainPlayer.xpos, MainPlayer.ypos + 1).Tiletype = Sign Then
                Else

                    VScroll1 = VScroll1.Value + 1

                End If
                Else
                'If at the edge of a map load the next map
                Call Form1.linker
                End If

            End If

End If



End Sub

Private Sub MovePlayerUp()
If linky <= 0 Then
    linky = 1
End If
If ZeldaMode = True Then Image1.Picture = teleporter.Zup.Picture

If Map(Linkx, linky - 1).Tiletype = NONWALKABLE Or Map(Linkx, linky - 1).Tiletype = Tree Or Map(Linkx, linky - 1).Tiletype = Gold Or Map(Linkx, linky - 1).Tiletype = Sign Then
NowALKDIRECTION.Caption = "Up"

End If

Shape2.Left = Snap(Image1.Left, 32)
Shape2.Top = Snap(Image1.Top, 32)
Direction.Caption = "up"

If NowALKDIRECTION.Caption = "Up" Then
Exit Sub
Else
NowALKDIRECTION.Caption = ""

If Image1.Top > 4 Then
Image1.Top = Image1.Top - 16

Else
If VScroll1.Value >= 1 Then

 MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
 MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll
If Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiletype = NONWALKABLE Or Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiletype = Tree Or Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiletype = Sign Then
Else

VScroll1 = VScroll1.Value - 1
End If
Else
Call Form1.linker

End If
End If
End If


End Sub

Private Sub AdjustItemsandMonsters()
If PlayerMoved = True Then
    'Adjust the items and Monsters Postions

    Call Stuff.MovenewMonster2
    Call Stuff.MovenewNpc
    Call ItemInfo.MovenewItem
    Call ItemInfo.ShowItem
    'Turn Guard Off
    TOOLS.CmdGuard.FontBold = False
    TOOLS.GuardTimer.Enabled = False
    Title.Visible = False
    Linkx = (Snap(Image1.Left, 32) \ 32) + HScroll
    linky = (Snap(Image1.Top, 32) \ 32) + VScroll
    'Adjust caption on main player
    lblPlayerName(0).Top = Image1.Top - 20
    lblPlayerName(0).Left = Image1.Left
    
    'Check if player is in stealth mode
        If MainPlayer.Stealth = True Then
            If MainPlayer.Stamina >= 1 Then
                MainPlayer.Stamina = MainPlayer.Stamina - 1
                MainPlayer.XP = MainPlayer.XP + 1
                MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 1
                PlayerSkills.StealthNextLevel = PlayerSkills.StealthNextLevel - 1
                Call RPG.CheckIfStatLeveled

            Else
                Form1.txtDialog.Text = Form1.txtDialog.Text & "You have failed to sneak" & vbNewLine
                TOOLS.CmdStealth.FontBold = False
                MainPlayer.Stealth = False
            End If
        End If
        
    PlayerMoved = False
    
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
CurX = X: CurY = y
If MapEd = True Then
'This moves the selected npc when you hold your house on it. In map edit mode

If TargetedNpc = 0 Or TargetedNpc = 1000 Then
Else

On Error GoTo nofile:
I = -1
For I = I + 1 To 250
If PICnpc(TargetedNpc).ToolTipText = NPc(I).PlayerName Then
PICnpc(TargetedNpc).Move X, y
NPc(I).xpos = (Snap(PICnpc(TargetedNpc).Left, 32) \ 32) + HScroll
NPc(I).ypos = (Snap(PICnpc(TargetedNpc).Top, 32) \ 32) + VScroll

Stuff.MovenewNpc
TargetedNpc = 1000
I = 250
End If
Next
nofile:

End If
End If

If MapEd = True Then
On Error GoTo nofile2:
'This moves the selected item when you hold your house on it. In map edit mode

PicItem(TargetedItem).Move X, y
I = -1
For I = I + 1 To 250
If PicItem(TargetedItem).ToolTipText = NewItem(I).ItemName Then
NewItem(I).xpos = (Snap(PicItem(TargetedItem).Left, 32) \ 32) + HScroll
NewItem(I).ypos = (Snap(PicItem(TargetedItem).Top, 32) \ 32) + VScroll
ItemInfo.MovenewItem
TargetedItem = 1000
I = 250
End If
Next
nofile2:

End If

'If mouse moved move square thingy
TargetShape.Left = Snap(X, 32): TargetShape.Top = Snap(y, 32)
mousex = (Snap(TargetShape.Left, 32) \ 32) + HScroll
mousey = (Snap(TargetShape.Top, 32) \ 32) + VScroll

If Button = 2 Then
'Makes all the popup menus
If TOOLS.Frame6.Visible = True Then
mnuseperator.Visible = True
mnuextratile.Visible = True
Else
mnuseperator.Visible = False
mnuextratile.Visible = False
End If
If Map(TOOLS.Label22.Caption, TOOLS.Label23.Caption).Tiletype = 2 Then
PopupMenu mnuactionmenu, , X, y
mnucuttree.Enabled = False
mnufish.Enabled = True
mnumine.Enabled = False
mnufarm.Enabled = False
mnudrink.Enabled = True
mnuharvest.Enabled = False
End If
If Map(TOOLS.Label22.Caption, TOOLS.Label23.Caption).Tiletype = 12 Then
PopupMenu mnuactionmenu, , X, y
mnucuttree.Enabled = False
mnufish.Enabled = False
mnumine.Enabled = True
mnufarm.Enabled = False
mnudrink.Enabled = False
mnuharvest.Enabled = False
End If
If Map(Int(mousex), Int(mousey)).Tiletype = Tree Or Map(TOOLS.Label22.Caption, TOOLS.Label23.Caption).TileX = 352 And Map(TOOLS.Label22.Caption, TOOLS.Label23.Caption).Tiley = 0 Then

mnucuttree.Enabled = True
mnufish.Enabled = False
mnumine.Enabled = False
mnufarm.Enabled = False
mnudrink.Enabled = False
mnuharvest.Enabled = False
PopupMenu mnuactionmenu, , X, y
End If

If Map(TOOLS.Label22.Caption, TOOLS.Label23.Caption).Tiletype = 1 Then
PopupMenu mnuactionmenu, , X, y
mnucuttree.Enabled = False
mnufish.Enabled = False
mnumine.Enabled = False
mnufarm.Enabled = True
mnudrink.Enabled = False
mnuharvest.Enabled = False
End If


If Map(TOOLS.Label22.Caption, TOOLS.Label23.Caption).Tiletype = 16 Then
PopupMenu mnuactionmenu, , X, y
mnucuttree.Enabled = False
mnufish.Enabled = False
mnumine.Enabled = False
mnufarm.Enabled = False
mnudrink.Enabled = False
mnuharvest.Enabled = True
End If
End If
'If button pressed paint tile

If TOOLS.Frame6.Visible = True Then

If Button = 1 And SelectedTool = "Normal Paint" Then
   PaintTile X, y
   
   Shape1.BorderColor = &HFFFFFF
End If

'Fill screen tool - Experimental this does not work (yet)
If Button = 1 And SelectedTool = "Fill Screen" Then
For y = Snap2(y, 32) + VScroll1.Value To Snap2(y, 32) + VScroll1.Value + 15
    For X = Snap2(X, 32) + HScroll1.Value To Snap2(X, 32) + HScroll1.Value + 19
        If Map(X, y).Tiletype = 0 Then
             Map(X, y).Tiletype = 1
             PaintTile X * 32, y * 32
             
        Else
             X = 0: y = y + 1
            If Map(X, y).Tiletype <> 0 Then Exit Sub
            X = -1
        End If
        Shape1.BorderColor = &HFFFFFF
    Next
Next
Me.Refresh
End If
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If TOOLS.Frame6.Visible = True Then

If Button = 1 Then Shape1.BorderColor = &HFF0000

End If
End Sub

Private Sub HScroll2_Change()
    FoxTime 1
    FoxBrightness Tiles.hDC, Tiles.Image.Handle, Tiles.hDC, Tiles.Picture.Handle, 0, HScroll2, 0

    Form1.Refresh
    RedrawMap
    
    labBrightness = "Brightness-level: " & HScroll2
    labBrightness.Refresh

End Sub

Private Sub Information_Timer()
'On Error GoTo nofile
If OnlineGame = True Then
MakeAFolder_Click

End If
Linkx = (Snap(Image1.Left, 32) \ 32) + HScroll
linky = (Snap(Image1.Top, 32) \ 32) + VScroll
Meter.Caption = "Life: " & MainPlayer.Health & " Stamina: " & MainPlayer.Stamina & " Mana " & MainPlayer.Mana
AttackShape.Left = Image1.Left - AttackShape.Width \ 2
AttackShape.Top = Image1.Top - AttackShape.Height \ 2


lblTileKind.Caption = Map(Linkx, linky).Tiletype
LblLeft.Caption = Image1.Left
Lbltop.Caption = Image1.Top
MainPlayer.CurrentWeapon = TOOLS.Weapon.ToolTipText
MainPlayer.MapName = MapName.Caption
MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll
Targetx = (Snap(TargetShape.Left, 32) \ 32) + HScroll
Targety = (Snap(TargetShape.Top, 32) \ 32) + VScroll

lblPlayerName(0).Caption = MainPlayer.PlayerName


Shape2.Left = Snap(Image1.Left, 16): Shape2.Top = Snap(Image1.Top, 16)
Call Stuff.MovenewMonster2
Call Stuff.MovenewNpc
Call ItemInfo.MovenewItem
Call CheckStoreScript

'Check if player is on a posion or hurt tile1
If lblTileKind.Caption = "5" Then
'posion tile! or hurt tile
    wait2 = wait2 + 1
    If wait2 >= 200 Then
        Form1.lblAttack.Visible = True
        Form1.lblAttack = "You lose one life!"
        Form1.TimerAttacklbl.Enabled = True
        MainPlayer.Health = MainPlayer.Health - 1
        Monsters.CheckDead
        TOOLS.LoadStats
        wait2 = 0
    End If
End If
'end hurt tile


If lblTileKind.Caption = "10" Then 'teleporter tile
    teleporter.Show vbModal, Me
End If

TOOLS.Caption = "Location: " & MainPlayer.MapName & " X " & MainPlayer.xpos & " Y " & MainPlayer.ypos
TOOLS.PlayerImage.Picture = Image1.Picture

If MainPlayer.MakeFolder = True Then
MakeAFolder_Click
End If
Form1.KeyPreview = True
End Sub
Private Sub CheckStoreScript()
'Checks if there is a store on that location and if there is brings up the window
'also called if npc is storeowner when you clicked on them
If MapEd = True Then
    Exit Sub
End If
'If frmNPCTrade.Visible = True Then
'Exit Sub
'End If
If ShopWindowOpen = True Then
    Exit Sub
End If

If lblTileKind.Caption = "22" Then
'load custom store

frmNPCTrade.Show
End If

If lblTileKind.Caption = "17" Then
'Store is Ranch

frmNPCTrade.Show
End If
If lblTileKind.Caption = "14" Then
'supply Store

frmNPCTrade.Show
    frmNPCTrade.LoadStore (App.Path & "\scripts\Supply Shop.shop")
    
End If
If lblTileKind.Caption = "13" Then
'food store

frmNPCTrade.Show
End If
If lblTileKind.Caption = "9" Then
'armor store

frmNPCTrade.Show
frmNPCTrade.LoadStore (App.Path & "\scripts\Armor Store.shop")
    
End If
If lblTileKind.Caption = "7" Then
'magic store

frmNPCTrade.Show
frmNPCTrade.LoadStore (App.Path & "\scripts\Magic Store.shop")
    
End If
If lblTileKind.Caption = "8" Then
'weapon store

frmNPCTrade.Show
frmNPCTrade.LoadStore (App.Path & "\scripts\Weapon Store.shop")
    
End If
End Sub

Private Sub MakeAFolder_Click()
If MainPlayer.MakeFolder = True Then
 If sckConnect.State = sckConnected Then

        SendList.Add "[Folder] " & "," & MainPlayer.AccountName & "," & MainPlayer.Password
        
        SendTo.Add "sckConnect"

    End If
     If sckConnect.State = sckConnected Then
        SendList.Add "[Password] " & MainPlayer.Password
        SendTo.Add "sckConnect"
        MainPlayer.MakeFolder = False
    End If

   End If
End Sub

Private Sub mnuattack_Click()
TOOLS.MenuAttack1
End Sub

Private Sub mnucredits_Click()
MsgBox "Jonathan Valentin Programming and game desigin, Rummager graphics,Rpg World Online for some art! Find the game at http://www.rpgwo.com, Mickey for ideas", vbInformation
End Sub

Private Sub mnucuttree_Click()
Treex = TOOLS.Label22.Caption
Treey = TOOLS.Label23.Caption
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "Saw" Then
    If PXItem(I).Quanity > 0 Then

If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call RPG.CutDownTree
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call RPG.CutDownTree
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call RPG.CutDownTree
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.CutDownTree
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.CutDownTree
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.CutDownTree
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call RPG.CutDownTree
End If
    Else
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You need a saw to cut down the tree" & vbNewLine
    End If
        I = MaxItems
End If
Next I
End Sub

Private Sub mnudrink_Click()
If Map(TOOLS.Label22.Caption, TOOLS.Label23.Caption).Tiletype = 2 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You drink the water" & vbNewLine
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "Unable to drink the water" & vbNewLine
End If

End Sub

Private Sub mnuextratile_Click()
AddMonster.TileX.Text = TOOLS.Label22.Caption
AddMonster.Tiley.Text = TOOLS.Label23.Caption
AddMonster.FrameSpecial.Visible = True
AddMonster.Show
End Sub

Private Sub mnufarm_Click()
Treex = TOOLS.Label22.Caption
Treey = TOOLS.Label23.Caption
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "Seeds" Then
  
    If PXItem(I).Quanity > 0 Then
         
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call RPG.PlantSeeds
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call RPG.PlantSeeds
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call RPG.PlantSeeds
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.PlantSeeds
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.PlantSeeds
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.PlantSeeds
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call RPG.PlantSeeds
End If
Else
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You need seeds to plant" & vbNewLine

    End If
    I = MaxItems
End If
Next I
End Sub
Private Sub mnufish_Click()
If TOOLS.Item.ToolTipText = "StickPole" Or TOOLS.Item.ToolTipText = "IronPole" Then
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call Stuff.Fishing
End If
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You need a fishing pole to fish!" & vbNewLine
End If
End Sub

Private Sub mnuharvest_Click()
Treex = TOOLS.Label22.Caption
Treey = TOOLS.Label23.Caption
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call RPG.HarvestSeeds(TOOLS.Label22.Caption, TOOLS.Label23.Caption)
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
    Call RPG.HarvestSeeds(TOOLS.Label22.Caption, TOOLS.Label23.Caption)
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call RPG.HarvestSeeds(TOOLS.Label22.Caption, TOOLS.Label23.Caption)
End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.HarvestSeeds(TOOLS.Label22.Caption, TOOLS.Label23.Caption)
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.HarvestSeeds(TOOLS.Label22.Caption, TOOLS.Label23.Caption)
End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
    Call RPG.HarvestSeeds(TOOLS.Label22.Caption, TOOLS.Label23.Caption)
End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
    Call RPG.HarvestSeeds(TOOLS.Label22.Caption, TOOLS.Label23.Caption)
End If
End Sub

Private Sub mnumine_Click()
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "Pick" Then
    If PXItem(I).Quanity > 0 Then
    
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
'Call Stuff.Mineforgold
Call RPG.Miner(MainPlayer.xpos + 1, MainPlayer.ypos)

End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
'Call Stuff.Mineforgold
Call RPG.Miner(MainPlayer.xpos - 1, MainPlayer.ypos)

End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
'Call Stuff.Mineforgold
Call RPG.Miner(MainPlayer.xpos, MainPlayer.ypos - 1)

End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
'Call Stuff.Mineforgold
Call RPG.Miner(MainPlayer.xpos, MainPlayer.ypos + 1)

End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
'Call Stuff.Mineforgold
Call RPG.Miner(MainPlayer.xpos + 1, MainPlayer.ypos + 1)

End If
If MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
'Call Stuff.Mineforgold
Call RPG.Miner(MainPlayer.xpos - 1, MainPlayer.ypos + 1)

End If
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
'Call Stuff.Mineforgold
Call RPG.Miner(MainPlayer.xpos + 1, MainPlayer.ypos - 1)

End If
If MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
Call Stuff.Mineforgold
'Call RPG.Miner(MainPlayer.xpos, MainPlayer.ypos + 1)

End If
    Else
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You need a pick to mine for gold" & vbNewLine

    End If
End If
Next
End Sub

Private Sub mnumusic_Click()
If mnumusic.Checked = True Then

Form1.MediaPlayer1.Stop

mnumusic.Checked = False
Exit Sub

End If
If mnumusic.Checked = False Then

Form1.MediaPlayer1.Play
mnumusic.Checked = True
End If
End Sub
Private Sub mnupickup_Click()
If MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
Call ItemInfo.PickupItems(MainPlayer.xpos + 1, MainPlayer.ypos)

'End If
ElseIf MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos = TOOLS.Label23.Caption Then
Call ItemInfo.PickupItems(MainPlayer.xpos - 1, MainPlayer.ypos)

'elseEnd If
ElseIf MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
Call ItemInfo.PickupItems(MainPlayer.xpos, MainPlayer.ypos - 1)

'End If
ElseIf MainPlayer.xpos = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
Call ItemInfo.PickupItems(MainPlayer.xpos, MainPlayer.ypos + 1)

'End If
ElseIf MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
Call ItemInfo.PickupItems(MainPlayer.xpos + 1, MainPlayer.ypos + 1)

'End If
ElseIf MainPlayer.xpos - 1 = TOOLS.Label22.Caption And MainPlayer.ypos + 1 = TOOLS.Label23.Caption Then
Call ItemInfo.PickupItems(MainPlayer.xpos - 1, MainPlayer.ypos + 1)

'End If
ElseIf MainPlayer.xpos + 1 = TOOLS.Label22.Caption And MainPlayer.ypos - 1 = TOOLS.Label23.Caption Then
Call ItemInfo.PickupItems(MainPlayer.xpos + 1, MainPlayer.ypos - 1)
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "Move closer to the item!" & vbNewLine

End If
End Sub

Private Sub mnusound_Click()
If mnusound.Checked = True Then
Form1.MediaPlayer2.Stop
mnusound.Checked = False
Exit Sub
End If
If mnusound.Checked = False Then
Form1.MediaPlayer2.Play
mnusound.Checked = True
End If
End Sub



Private Sub Events_Timer()


If MapEd = True Then
Exit Sub
End If
Slow2 = Slow2 + 1
Call CheckForEvents

If MainPlayer.MapName = "c1.map" And MainPlayer.xpos = 18 And MainPlayer.ypos = 10 Then
Call Pics.TeleportPortal

End If

'This is the older custom script system that i had
'Good for maps that have many entrances and exits
If Map(MainPlayer.xpos, MainPlayer.ypos).Tiletype = 20 Then
On Error GoTo nofile434

  Open App.Path & "\scripts\" & MainPlayer.MapName & MainPlayer.xpos & MainPlayer.ypos & ".warp" For Input As #1
    Line Input #1, textline1 'map
    Line Input #1, textline2 'x
    Line Input #1, textline3 'y

   Close #1

Open App.Path & "\maps\" & textline1 For Input As #1: Close
Open App.Path & "\maps\" & textline1 For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Call RPG.FastXYWarp(Int(textline2), Int(textline3))
Call Stuff.CenterPerson
If textline6 = "tiles.bmp" Then
Form1.Tiles.Picture = LoadPicture(App.Path & "\" & textline6)
End If
 MainPlayer.MapName = textline1
Form1.MapName = textline1

nofile434:

 MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
 MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll
Exit Sub
End If


If OnlineGame = True Then
SendChatTmr = SendChatTmr + 1
    If SendChatTmr >= 10 Then
'If oldx = MainPlayer.xpos And oldy = MainPlayer.ypos Then
        If PlayerMoved = True Then
            Call SendPlayerPosition
            SendChatTmr = 0
            PlayerMoved = False
        Else
'oldx = MainPlayer.xpos
'oldx = MainPlayer.ypos

        SendChatTmr = 0
        End If
    End If
End If

ItemInfo.MovenewPlayer
'Animation Loop
If Slow2 > 40 And AnimatedTiles = True Then
Call AnimtaionLoop
Slow2 = 0
RedrawMap
ElseIf AnimatedTiles = False Then
Slow2 = 0
End If

TimeofDay3 = TimeofDay3 + 2

 MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
 MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll

Call CheckExitTile

If MainPlayer.MapName = "palace1" Then
If MainPlayer.ypos >= 15 Then
  Open App.Path & "\maps\newbie" For Input As #1: Close
Open App.Path & "\maps\newbie" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "newbie"
Form1.MapName = "newbie"
Form1.RedrawMap

HScroll1.Value = 12
VScroll1.Value = 23
Image1.Top = 288
Image1.Left = 328
End If

End If



'Disables Attacking if stamina is <=0
If MainPlayer.Stamina <= 0 Then
MainPlayer.Stamina = 0
TOOLS.CmdAttack.Enabled = False
Else
TOOLS.CmdAttack.Enabled = True

End If


'##Begin Night and Day stuff
If TimeofDay3 > 2500 And SwitchToDay = False Then
HScroll2.Value = HScroll2.Value - 1
TimeofDay3 = 0


End If
If HScroll2.Value <= -169 Then
SwitchToDay = True
End If
If HScroll2.Value >= 0 Then
SwitchToDay = False
End If
If TimeofDay3 > 2500 And SwitchToDay = True Then
HScroll2.Value = HScroll2.Value + 1
TimeofDay3 = 0


End If
Me.Visible = True
'##End Night and Day stuff

'Calls additional events
Call MoreEvents


End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
TOOLS.Caption = "Location: " & MainPlayer.MapName & " X " & MainPlayer.xpos & " Y " & MainPlayer.ypos
'Tell the TileX and TileY where they are
CurX = X: CurY = y

'If mouse moved move square thingy

Shape1.Left = Snap(X, 32): Shape1.Top = Snap(y, 32)

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

response = MsgBox("Are you sure you want to exit the world?", vbYesNo, "Quit?")
If response = vbYes Then
txtMessage.Text = MainPlayer.PlayerName & " has left the world!"
cmdSend_Click
Mutiplayer.Timer1.Enabled = True
RPG.SaveCharecter
RPG.SaveMonsterLocation
RPG.SaveNpcLocation
ItemInfo.SaveItemLocation
'Save Charecter
If OnlineGame = True Then
Call Form1.DisconnectCharacter(MainPlayer.PlayerName)
txtMessage.Text = MainPlayer.PlayerName & " has left the world!"
cmdDisconnect_Click
cmdSend_Click
End If
K = 0
For K = K + 1 To 250
newmon = "MonsterImage" & K
On Error Resume Next
  Set P51 = Form1.Controls.Remove(newmon)
Next K

'Erase the arrays from memory
Erase Monster

Erase TrapsArray
Erase NPc
Erase Player


Erase NewItem
Erase Map

If OnlineGame = True Then
    Call Server.SaveChatLog


    Dim I As Integer

'This flag is set after you have been disconnected from the server following an attempt to close the program.
'When this flag is set, the form will just unload - skipping the code below.
    If bJustQuit Then
    'Call the procedure that writes the user preferences to an INI file.
        WritePreferences
        Exit Sub
    End If

'If you are connected to the server, send a request to the server to be disconnected.
'Do not allow the user to quit the program until disconnected from the server.
'If sckConnect.State = sckConnected Then
    'Send the request to be disconnected.
    SendList.Add "[Disconnect]"
    SendTo.Add "sckConnect"
    'Update the status.
    txtStatus.Text = txtStatus.Text & vbCrLf & "Disconnecting..."
    'Set the sckConnect.Tag property to "Quit" so when disconnected, the program knows to quit immediately.
    sckConnect.Tag = "Quit"

    cmdDisconnect_Click
   
'End If

    'Close all connections.
    sckConnect.Close
    For I = 0 To giMaxConnections - 1
        sckConnection(I).Close
    Next I
    
'Call the procedure that writes the user preferences to an INI file.
WritePreferences
End If


Unload TOOLS
Unload Items
Unload Me


Exit Sub
Else
Cancel = True
Exit Sub
End If


End Sub

Private Sub HScroll1_Change()
HScroll = HScroll1.Value
RedrawMap
Stuff.MovenewMonster2
Stuff.MovenewNpc
ItemInfo.MovenewItem
ItemInfo.ShowItem
Call Monsters.MonsterReturn
End Sub
Private Sub HScroll1_Scroll()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
HScroll = HScroll1.Value
'
RedrawMap
Stuff.MovenewMonster2
Stuff.MovenewNpc
ItemInfo.MovenewItem
ItemInfo.ShowItem
Call Monsters.MonsterReturn
End Sub
Private Sub mexit_Click()
response = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit?")
If response = vbYes Then
'Terminate
End
End If
End Sub
Private Sub LoadDefault()

ReDim Map(23 * mWidth, 16 * mHeight) 'Set the ubound of the map array
ReDim SpecialMap(23 * mWidth, 16 * mHeight)
'Set the max of the scroll bars according to the map size
HScroll1.Max = 23 * (mWidth - 1)
VScroll1.Max = 16 * (mHeight - 1)
      
End Sub
Private Sub Form_Load()

Dim I As Integer
SwitchToDay = False
Shape2.Visible = False
Form1.Width = 8280
Image1.Left = 328
Image1.Top = 256
txtIP.Text = MainPlayer.ServerIP

picSwan.Picture = LoadPicture(App.Path & "\scripts\blend.bmp")
BackPic.Picture = LoadPicture(App.Path & "\stone.bmp")
'Load the skins
If SkinFilename = "" Then
Else
Skn.LoadSkin SkinFilename ' Loads another skin into Skin component
Skn.ApplySkin CmdshowStats.hwnd ', SkinFilename
Skn.ApplySkin CmdHelp.hwnd ', SkinFilename
Skn.ApplySkin CmdQuests.hwnd ', SkinFilename
Skn.ApplySkin CmdMenu.hwnd ', SkinFilename
Skn.ApplySkin CmdItem.hwnd ', SkinFilename
Skn.ApplySkin CmdCenter.hwnd ', SkinFilename
Skn.ApplySkin CmdGuard.hwnd
Skn.ApplySkin CmdAttack.hwnd
Skn.ApplySkin CmdStealth.hwnd
Skn.ApplySkin CmdScan.hwnd
Skn.ApplySkin cmdSend.hwnd
Skn.ApplySkin CmdStats.hwnd
'Skn.ApplySkin TOOLS.CmdCFull.hwnd
End If


MainPlayer.Armor = 0

lblAttack.Caption = ""

lblPlayerName(0).Top = Image1.Top - 20
lblPlayerName(0).Left = Image1.Left
txtname.Text = MainPlayer.PlayerName

MediaPlayer2.Filename = App.Path & "\beep.wav"

txtDialog.Text = txtDialog.Text & "Welcome to " & "Project X Online" & vbNewLine
'txtDialog.Text = txtDialog.Text & "Players: " & lstConnections.List

MediaPlayer2.Filename = App.Path & "\beep.wav"
If OnlineGame = True Then
'Call the procedure that gets the user preferences from an INI file.
GetPreferences
'Load enough Winsock controls to handle all connections.
For I = 1 To giMaxConnections - 1
    Load sckConnection(I)
Next I
cmdConnect_Click
End If

    If MainPlayer.Graphic = 1 Then
Form1.Image1.Picture = Pics.Image1.Picture
End If
    If MainPlayer.Graphic = 2 Then
Form1.Image1.Picture = Pics.Image2.Picture
End If
    If MainPlayer.Graphic = 3 Then
Form1.Image1.Picture = Pics.Image3.Picture
End If
    If MainPlayer.Graphic = 4 Then
Form1.Image1.Picture = Pics.Image4.Picture
End If
    If MainPlayer.Graphic = 5 Then
Form1.Image1.Picture = Pics.Image5.Picture
End If
    If MainPlayer.Graphic = 6 Then
Form1.Image1.Picture = Pics.Image6.Picture
End If
    If MainPlayer.Graphic = 7 Then
Form1.Image1.Picture = Pics.Deer.Picture
End If
    If MainPlayer.Graphic = 8 Then
Form1.Image1.Picture = Pics.Chicken.Picture
End If
    If MainPlayer.Graphic = 9 Then
Form1.Image1.Picture = Pics.Leperchaun.Picture
End If
    If MainPlayer.Graphic = 10 Then
Form1.Image1.Picture = Pics.King.Picture
End If
    If MainPlayer.Graphic = 11 Then
Form1.Image1.Picture = Pics.RockMan.Picture
End If
    If MainPlayer.Graphic = 12 Then
Form1.Image1.Picture = Pics.Mutant.Picture
End If
    If MainPlayer.Graphic = 13 Then
Form1.Image1.Picture = Pics.Lion.Picture
End If
    If MainPlayer.Graphic = 14 Then
Form1.Image1.Picture = Pics.LostMage.Picture
End If
    If MainPlayer.Graphic = 15 Then
Form1.Image1.Picture = Pics.MagicTiger.Picture
End If
    If MainPlayer.Graphic = 16 Then
Form1.Image1.Picture = Pics.Undead.Picture
End If
    If MainPlayer.Graphic = 17 Then
Form1.Image1.Picture = Pics.Pirate.Picture
End If

'Verion Info
Form1.Caption = MainPlayer.PlayerName & " Project X  Online Jonathan Valentin 2001-2003 Verion 0.40"



Form1.Show

MediaPlayer1.Filename = App.Path & "\" & "z5oot[2].mid"

Tiles.Picture = LoadPicture(App.Path & "\TILES.BMP")
SelectedTool = "Normal Paint"
mWidth = 4
mHeight = 4
LoadDefault


Form1.KeyPreview = False

 Close
'If Warping = True Then
'Else
'O'n Error GoTo nofile: 'If file doesnt exist dont open it
Open App.Path & "\maps\newbie" For Input As #1: Close
Open App.Path & "\maps\newbie" For Binary Access Read Lock Read As #1

EraseAll
Get #1, , Map
Close
MapName.Caption = "newbie"
HScroll1.Value = "45"
VScroll1.Value = "23"
RedrawMap
RPG.LoadMonsters
RPG.LoadNpcs
ItemInfo.LoadItems
Form1.ShowNpc
'End If
Randomize

'Setup the defualt lightining colors

'Set Center
Bolt.Center = RGB(207, 207, 233)
'Set Inner
Bolt.Inner1 = RGB(176, 176, 207)
Bolt.Inner2 = RGB(176, 176, 192)
Bolt.Inner3 = RGB(160, 160, 176)
Bolt.Inner4 = RGB(145, 145, 176)
'Set Outer
Bolt.Outer1 = RGB(90, 90, 110)
Bolt.Outer2 = RGB(80, 80, 110)
Bolt.Outer3 = RGB(65, 65, 80)
Bolt.Outer4 = RGB(50, 50, 70)
'Set Properties
Bolt.Nicks = NickNum.Value
Bolt.VelocityMin = 3
Bolt.VelocityMax = 15
Bolt.Height = HeightNum.Value
Bolt.SameStart = False
Bolt.ShowCloud = False
Bolt.SplitBolt = True

'Exit Sub
nofile:
RedrawMap

End Sub
Private Sub PaintTile(X As Single, y As Single)
If X <= Form1.ScaleWidth And X >= 0 And y <= Form1.ScaleHeight And y >= 0 Then

    
    BitBlt Form1.hDC, Snap(X, 32), Snap(y, 32), 32, 32, Tiles.hDC, TempTileX, TempTileY, SRCCOPY
    
    'Holds data for where the info should be saved to the map array
    tempx = Snap(X, 32) \ 32 + HScroll1.Value
    tempy = Snap(y, 32) \ 32 + VScroll1.Value

    'Save map info to the map array
    Map(tempx, tempy).TileX = TempTileX
    Map(tempx, tempy).Tiley = TempTileY
If TOOLS.Combo1.Text = "TrapTILE" Then
Map(tempx, tempy).Tiletype = TrapTILE
End If
If TOOLS.Combo1.Text = "ExitTile" Then
Map(tempx, tempy).Tiletype = ExitTile
End If
If TOOLS.Combo1.Text = "LoadWarp" Then
Map(tempx, tempy).Tiletype = LoadWarp
End If
If TOOLS.Combo1.Text = "Destroyable" Then
Map(tempx, tempy).Tiletype = Destroyable
End If
If TOOLS.Combo1.Text = "WeaponStore" Then
Map(tempx, tempy).Tiletype = WeaponStore
End If
If TOOLS.Combo1.Text = "MagicStore" Then
Map(tempx, tempy).Tiletype = MagicStore
End If
 If TOOLS.Combo1.Text = "WarpPortal" Then
 Map(tempx, tempy).Tiletype = WarpPortal
End If
 If TOOLS.Combo1.Text = "HorseRanch" Then
 Map(tempx, tempy).Tiletype = HorseRanch
End If
 If TOOLS.Combo1.Text = "SupplyStore" Then
 Map(tempx, tempy).Tiletype = SupplyStore
End If
If TOOLS.Combo1.Text = "Tree" Then
 Map(tempx, tempy).Tiletype = Tree
End If
If TOOLS.Combo1.Text = "Gold" Then
 Map(tempx, tempy).Tiletype = Gold
End If
If TOOLS.Combo1.Text = "ArmorStore" Then
 Map(tempx, tempy).Tiletype = ArmorStore
End If
If TOOLS.Combo1.Text = "HurtTile" Then
 Map(tempx, tempy).Tiletype = HurtTile
End If
If TOOLS.Combo1.Text = "Swimable" Then
Map(tempx, tempy).Tiletype = SWIMABLE
End If
If TOOLS.Combo1.Text = "Walkable" Then
Map(tempx, tempy).Tiletype = WALKABLE
End If
If TOOLS.Combo1.Text = "NONWALKABLE" Then
Map(tempx, tempy).Tiletype = NONWALKABLE
End If
If TOOLS.Combo1.Text = "DoorTile" Then
Map(tempx, tempy).Tiletype = DoorTile
End If
If TOOLS.Combo1.Text = "DieTile" Then
Map(tempx, tempy).Tiletype = DieTile
End If
If TOOLS.Combo1.Text = "Sign" Then
Dim SignText As String
Map(tempx, tempy).Tiletype = Sign
  F = FreeFile
SignText = InputBox("Please enter the sign text", "Sign Text")
Open (App.Path & "\scripts\" & MainPlayer.MapName & tempx & tempy & ".sign") For Binary Access Write Lock Write As F
Put #F, , SignText & vbCrLf
Close F
End If
    RefreshMini = True
End If
End Sub

Public Sub RedrawMap()
Dim y As Integer
Dim X As Integer
If Warping = True Then Exit Sub
'Draw the map Algorithm
'Form1.KeyPreview = False

Form1.Cls

For y = VScroll1.Value To VScroll1.Value + 16
    For X = HScroll1.Value To HScroll1.Value + 23
        If Map(X, y).Tiletype >= 0 Then 'If the tiletype is anything but 0
       BitBlt Form1.hDC, Snap2(X, 32) - Snap2(HScroll1.Value, 32), Snap2(y, 32) - Snap2(VScroll1.Value, 32), 32, 32, Tiles.hDC, Map(X, y).TileX, Map(X, y).Tiley, SRCCOPY
  
        If ToggleTileType = True Then
        'Shows what type of tile it is
        PicTileType.Cls
        PicTileType.BackColor = vbBlack
        PicTileType.Forecolor = vbWhite
        PicTileType.CurrentX = 0
        PicTileType.CurrentY = 0
        PicTileType.Print Map(X, y).Tiletype
        PicTileType.FontSize = 4
        BitBlt Form1.hDC, Snap2(X, 32) - Snap2(HScroll1.Value, 32), Snap2(y, 32) - Snap2(VScroll1.Value, 32), 16, 10, PicTileType.hDC, 1, 1, SRCCOPY
            End If

        End If
   Next
Next

Form1.ShowMonsters
'Form1.KeyPreview = True
Form1.Refresh
    RefreshMini = True
'Checks if there are lights
Call LightLoop

End Sub

Private Sub mnuabout_Click()
MsgBox "Jonathan Valentin June 17,2001  Revised many times check log November 19, 2001 Revised Feb 4, 2002 Revised Augest 6,2002 Revised Feb 11, 2003", vbInformation
End Sub

Private Sub mnuUse_Click()
If TOOLS.Item.ToolTipText = "" Then
MsgBox "Please select an item to use on this item!"

Else
'Call uses

End If
End Sub




Private Sub MonsterAttackTimer_Timer()
If MapEd = True Then
MonsterAttackTimer.Enabled = False
End If

Slow = Slow + 1
'Does all Monster AI Functions
Call Monsters.MonsterCastSpell
Call Monsters.MonsterAI
Call Monsters.MonsterScan
Call Monsters.MonsterStealth
If TOOLS.CmdGuard.FontBold = True Then
Call TOOLS.GuardMode
End If
If Slow > 4 Then
    Call Monsters.MoveNpc
    Slow = 0
    End If
End Sub

Private Sub PicItem_Click(Index As Integer)
If MapEd = True Then
TargetedItem = Index
Exit Sub

End If
mnucuttree.Visible = False
mnuharvest.Visible = False
mnufish.Visible = False
mnumine.Visible = False
mnufarm.Visible = False
mnudrink.Visible = False
mnupickup.Enabled = True
mnupickup.Visible = True
mnuUse.Visible = True
PopupMenu mnuactionmenu, , CurX, CurY

End Sub

Private Sub PICnpc_Click(Index As Integer)
'this sub targets npcs for gossip and trading
If MapEd = True Then
TargetedNpc = Index

Exit Sub
End If
If PICnpc(Index).ToolTipText = "" Then Exit Sub
I = 0
For I = I + 1 To 250
If PICnpc(Index).ToolTipText = NPc(I).PlayerName Then
TargetedNpc = Index
Form1.txtDialog.Text = Form1.txtDialog.Text & "You have targeted " & NPc(I).PlayerName & vbNewLine
End If

Next
End Sub


Private Sub ReloadBad_Timer()
Dim I As Integer
I = 0
For I = I + 1 To 200 ' Iterate through each element.
  On Error GoTo nofile:
    newmon = "MonsterImage" & I
If Monster(0).Health <= 0 Then
      
     Monster(0).Health = Monster(0).MaxHealth
     
     Monster(0).Stamina = Monster(0).MaxStamina

    End If
  Set P51 = Form1.Controls(newmon)
With P51
If Monster(I).PlayerName = "Elk Warrior Boss" Then
If Bosses(0).Boss1Defeated = True Then Exit Sub
If Monster(I).Health <= 0 Then
Call RPG.Boss1Deafeated
Bosses(0).Boss1Defeated = True
Monster(I).Health = 100

End If
End If
    If Monster(I).Health <= 0 Then ' If Text equals "Hello".
      
     .Picture = LoadPicture(App.Path & "\monsters\" & Monster(I).PlayerName & ".gif")
    
     
     .Tag = ""
     Monster(I).Health = Monster(I).MaxHealth
     
     Monster(I).Stamina = Monster(I).MaxStamina
     
     Monster(I).Mana = Monster(I).MaxMana
       ' Exit loop.
    End If
           .ToolTipText = Monster(I).PlayerName & " Level " & Monster(I).Level & " Health " & Monster(I).Health & "\" & Monster(I).MaxHealth
End With
    If Monster(I).MaxHealth = 0 Then

        Exit For
     End If

Next
nofile:

Exit Sub
End Sub


Private Sub StatTimer_Timer()
   'Makes the cool health mana and stamina bars!
   If MapEd = True Then
   StatTimer.Enabled = False
   Else
   On Error GoTo nofile
   HPShp.Width = (((MainPlayer.Health / 100) / (MainPlayer.MaxHealth / 100)) * 150)

   MANShp.Width = (((MainPlayer.Mana / 100) / (MainPlayer.MaxMana / 100)) * 150)

   STAShp.Width = (((MainPlayer.Stamina / 100) / (MainPlayer.MaxStamina / 100)) * 150)


nofile:
Exit Sub

End If

End Sub
Private Sub TimerAttacklbl_Timer()
lblAttack.Visible = False
TimerAttacklbl.Enabled = False


End Sub




Private Sub TmrSnow_Timer()

For I = 0 To 100
'Me.Caption = I


Snow(I).X = Int(Rnd * Me.ScaleWidth)
Snow(I).y = Int(Rnd * Me.ScaleHeight)
'Me.Caption = Snow(I).X & " " & Snow(I).y

Form1.Circle (Snow(I).X, Snow(I).y), 1, vbWhite


Next
Me.Refresh
'RedrawMap
End Sub

Private Sub VScroll1_Change()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
VScroll = VScroll1.Value
RedrawMap
Stuff.MovenewMonster2
Stuff.MovenewNpc
ItemInfo.MovenewItem
Call Monsters.MonsterReturn
End Sub
Private Sub VScroll1_Scroll()
VScroll = VScroll1.Value
RedrawMap
Stuff.MovenewMonster2
Stuff.MovenewNpc
ItemInfo.MovenewItem
 Call Monsters.MonsterReturn
End Sub

Public Sub EraseAll()
Dim y As Integer
Dim X As Integer
'Erase the map

For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        
            Map(X, y).Tiletype = 0
            Map(X, y).TileX = 0
            Map(X, y).Tiley = 0
        
    Next
Next
'Call EraseSpecialMap
RedrawMap
End Sub
Public Sub EraseSpecialMap()
'You can work on the special map layer it was kind of bug for me
'Dim y As Integer
'Dim X As Integer
'Erase Special map

'For y = 0 To UBound(SpecialMap, 2)
    'For X = 0 To UBound(SpecialMap, 1)
        
            'SpecialMap(X, y).Amountleft = 0
            'SpecialMap(X, y).Brightness = 0
            'SpecialMap(X, y).Choice = 0
            'SpecialMap(X, y).HScroll = 0
            'SpecialMap(X, y).KeyNumber = 0
            'SpecialMap(X, y).Locked = 0
            'SpecialMap(X, y).MapName = ""
            'SpecialMap(X, y).MapTiles = ""
            'SpecialMap(X, y).Mapx = 0
            'SpecialMap(X, y).Mapy = 0
            'SpecialMap(X, y).MineAmount = 0
           ' SpecialMap(X, y).MineType = 0
          '  SpecialMap(X, y).Music = ""
         '   SpecialMap(X, y).PlayerLeft = 0
        '    SpecialMap(X, y).PlayerTop = 0
       '     SpecialMap(X, y).Seeds = 0
      '      SpecialMap(X, y).SoundEffect = ""
     '       SpecialMap(X, y).SoundEffectPlayed = False
    '        SpecialMap(X, y).VenderItem = ""
   '         SpecialMap(X, y).VScroll = 0
  '          SpecialMap(X, y).Weather = ""
 '   Next
'Next
'RedrawMap
End Sub


Private Sub RefreshMiniMap()
MiniMap.Cls
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        If Map(X, y).Tiletype = 1 Then
            For tY = y To y + 1.2
                For tX = X To X + 1.5
                   BitBlt MiniMap.hDC, tX, tY, 1.5, 1.2, TOOLS.Storage.hDC, Map(X, y).TileX + 2, Map(X, y).Tiley + 2, SRCCOPY
                Next
            Next
        End If
    Next
Next
   MiniMap.Refresh
   RefreshMini = False
End Sub


Public Sub linker() 'Links all the maps together

On Error GoTo nofile
  If Direction.Caption = "up" Then
  Open App.Path & "\maps\" & MapName.Caption & ".link" For Input As #1
   Line Input #1, textline1 'Up
   Line Input #1, textline2 'Down
   Line Input #1, textline3 'left
   Line Input #1, textline4 'right

    Close #1


'On Error GoTo nofile: 'If file doesnt exist dont open it
Open App.Path & "\maps\" & textline1 For Input As #1: Close
Open App.Path & "\maps\" & textline1 For Binary Access Read Lock Read As #1
EraseAll
Get #1, , Map
'begin test of extra data
Get #1, , MapExtra
'end test of extra data
Close
RedrawMap
VScroll1.Value = "48"
Image1.Top = "384"
MapName.Caption = textline1
MainPlayer.MapName = textline1
On Error GoTo Nofile9
'Open App.Path & "\maps\" & textline1 & ".spcm" For Input As #1: Close
'Open App.Path & "\maps\" & textline1 & ".spcm" For Binary Access Read Lock Read As #1
'Form1.EraseSpecialMap
'Get #1, , SpecialMap
'Close
End If
If Direction.Caption = "down" Then
  Open App.Path & "\maps\" & MapName.Caption & ".link" For Input As #1
   Line Input #1, textline1 'Up
   Line Input #1, textline2 'Down
   Line Input #1, textline3 'left
   Line Input #1, textline4 'right

    Close #1

    
    Open App.Path & "\maps\" & textline2 For Input As #1: Close
Open App.Path & "\maps\" & textline2 For Binary Access Read Lock Read As #1
EraseAll
Get #1, , Map
'begin test of extra data
Get #1, , MapExtra
'end test of extra data
Close

RedrawMap
VScroll1.Value = "0"

Image1.Top = "16"

MapName.Caption = textline2
MainPlayer.MapName = textline2
On Error GoTo Nofile9
'Open App.Path & "\maps\" & textline2 & ".spcm" For Input As #1: Close
'Open App.Path & "\maps\" & textline2 & ".spcm" For Binary Access Read Lock Read As #1
'Form1.EraseSpecialMap
'Get #1, , SpecialMap
'Close
End If
If Direction.Caption = "left" Then

 Open App.Path & "\maps\" & MapName.Caption & ".link" For Input As #1
   Line Input #1, textline1 'Up
   Line Input #1, textline2 'Down
   Line Input #1, textline3 'left
   Line Input #1, textline4 'right

    Close #1

Open App.Path & "\maps\" & textline3 For Input As #1: Close
Open App.Path & "\maps\" & textline3 For Binary Access Read Lock Read As #1
EraseAll
Get #1, , Map
'begin test of extra data
Get #1, , MapExtra
'end test of extra data
Close
RedrawMap
HScroll1.Value = "65"
Image1.Left = "520"
MapName.Caption = textline3
MainPlayer.MapName = textline3
On Error GoTo Nofile9
'open App.Path & "\maps\" & textline3 & ".spcm" For Input As #1: Close
'Open App.Path & "\maps\" & textline3 & ".spcm" For Binary Access Read Lock Read As #1
'Form1.EraseSpecialMap
'Get #1, , SpecialMap
'Close
End If
If Direction.Caption = "right" Then
Open App.Path & "\maps\" & MapName.Caption & ".link" For Input As #1
   Line Input #1, textline1 'Up
   Line Input #1, textline2 'Down
   Line Input #1, textline3 'left
   Line Input #1, textline4 'right

    Close #1
    
Open App.Path & "\maps\" & textline4 For Input As #1: Close
Open App.Path & "\maps\" & textline4 For Binary Access Read Lock Read As #1
EraseAll
Get #1, , Map
'begin test of extra data
Get #1, , MapExtra
'end test of extra data
Close
RedrawMap
HScroll1.Value = 0
Image1.Left = 8
MapName.Caption = textline4
MainPlayer.MapName = textline4
On Error GoTo Nofile9
'Open App.Path & "\maps\" & textline4 & ".spcm" For Input As #1: Close
'Open App.Path & "\maps\" & textline4 & ".spcm" For Binary Access Read Lock Read As #1
'Form1.EraseSpecialMap
'Get #1, , SpecialMap
'Close
End If

Form1.ShowMonsters


nofile:
Exit Sub

Nofile9:
Form1.EraseSpecialMap
Exit Sub
End Sub



Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
'Data has arrived at the computer connected to the server.

Dim sString As String

'Get the data.
sckConnect.GetData sString, vbString

'Process the data.  Pass -1 for the computer sending the data because it was from the server.
ProcessData sString, -1
End Sub
Private Sub sckConnection_Close(Index As Integer)
'One of the connections to the server was closed.

'Close the connection.
sckConnection(Index).Close

'If a chat room is up, begin to listen again.
If lstConnections.ListCount > 0 And sckConnect.State <> sckConnected Then
    'Call the procedure that ensures that one winsock control is listening for connections.
    MakeListen
End If
End Sub
Private Sub sckConnection_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'A connection was requested from the server.

'Close the port.
sckConnection(Index).Close
'Accept the connection.
sckConnection(Index).Accept requestID

'Send information on the Host to the connection.
SendList.Add "[Host] " & txtname.Text
SendTo.Add "sckConnection(" & Index & ")"

'Call the procedure that ensures that one winsock control is listening for connections.
MakeListen
End Sub
Private Sub sckConnection_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'Data has arrived at the server from an open connection.
Dim sString As String

'Get the data.
sckConnection(Index).GetData sString, vbString

'Process the data.  Pass the index of the connection from which the data came.
ProcessData sString, Index
End Sub
Private Sub sckConnect_Close()
'This occurs when the connection to the server is broken.

'Update the status.
txtStatus.Text = txtStatus.Text & vbCrLf & "Disconnected."
'Close the connection
sckConnect.Close
'Clear the names list.
lstConnections.Clear
'Clear the label holding the host's name.
lblHost.Caption = ""

'Clear stuff to start a new chat room (name list, dialog, etc.)
ClearStuff
response = MsgBox("You have been disconnected!  Do you want to reconnect?", vbYesNo, "Reconnect?")
If response = vbYes Then
'reconnects the user
cmdConnect_Click

Else
End
End If
'If the user tried to quit the program, generating the request to be disconnected, set the flag to quit and unload the form.  (See Form_Unload.)
If sckConnect.Tag = "Quit" Then
    bJustQuit = True
    Unload Me
End If
End Sub

Private Sub tmrSendData_Timer()
'The is the timer that continuously checks for data to send.

'Remembers whether or not something has been sent.
'Only one piece of data can be sent at a time, otherwise the data runs togeter.
Dim bSent As Boolean

'Index variable to determine which piece of data from the queue will be sent.
Dim iSend As Long

'Remembers where the data will be sent.
Dim iConnection As Integer

'Start the index variable at 1.
iSend = 1

'Loop while nothing has been sent and while the index variable is less than the maximum.
Do While bSent = False And iSend <= SendTo.Count
    If SendTo.Item(iSend) = "sckConnect" And sckConnect.State = sckConnected Then
        'Check to see if it is to be sent to the server and make sure the connection is still open.
        
        'Send the data.
        sckConnect.SendData SendList.Item(iSend) & vbCrLf
    
        'Delete the data from the queue.
        SendTo.Remove iSend
        SendList.Remove iSend
        
        'Something has been sent.
        bSent = True
    ElseIf Mid(SendTo.Item(iSend), 1, 13) = "sckConnection" Then
        'Check to see if it is to be sent to one of the connections to you, the server.
    
        'Parse the string containing the name of the connection to determine which connection to send to.
        iConnection = Mid(SendTo.Item(iSend), 15, Len(SendTo.Item(iSend)) - 15)
        
        'Ensure that the connection is open.
        If sckConnection(iConnection).State = sckConnected Then
            'Send the data.
            sckConnection(iConnection).SendData SendList.Item(iSend) & vbCrLf
        
            'Delete the data from the queue.
            SendTo.Remove iSend
            SendList.Remove iSend
            
            'Something has been sent.
            bSent = True
        End If
    End If
    
    'Increment index variable.
    iSend = iSend + 1
Loop
End Sub
Private Sub txtDialog_Change()
'The text in the message dialog text box has changed.

'Put the selection point at the end of the text box so you are seeing the most recent text.
txtDialog.SelStart = Len(txtDialog.Text)
End Sub
Private Sub txtIP_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    'If enter was pressed in the text box that inputs an IP address to connect to, simulate the pressing of the Connect button.
    cmdConnect_Click
    'Make VB think nothing was pressed on the keyboard.  This prevents it from making an annoying beep.
    KeyAscii = 0
End If
End Sub
Private Sub txtMessage_KeyPress(KeyAscii As Integer)


If KeyAscii = vbKeyReturn Then
    'If enter was pressed in the text box that inputs a message to send, simulate the pressing of the Send button.
    Call NoBadChat
    cmdSend_Click
     Call NoBadChat
    'Clear the text box.
    txtMessage.Text = ""
    'Make VB think nothing was pressed on the keyboard.  This prevents it from making an annoying beep.
    KeyAscii = 0
End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim I As Integer

If KeyAscii = vbKeyReturn Then
    'If enter was pressed in the text box that inputs your name, update your name on your screen and on all other computers.
    
    'If you are connected to the server, send the Change Name command.
    If sckConnect.State = sckConnected Then
        SendList.Add "[Change Name] " & txtname.Text
        SendTo.Add "sckConnect"
    End If
    
    'If you are the server, update your name on your screen.
    If lstConnections.ListCount > 0 And sckConnect.State <> sckConnected Then
        lblHost.Caption = txtname.Text
        
        'Send your name to all open connections.
        For I = 0 To giMaxConnections - 1
            If sckConnection(I).State = sckConnected Then
                SendList.Add "[Host Change Name] " & txtname.Text
                SendTo.Add "sckConnection(" & I & ")"
            End If
        Next I
    End If
        
    'Make VB think nothing was pressed on the keyboard.  This prevents it from making an annoying beep.
    KeyAscii = 0
End If
End Sub
Private Sub txtStatus_Change()
'The text in the status text box has changed.

'Put the selection point at the end of the text box so you are seeing the most recent text.
txtStatus.SelStart = Len(txtStatus.Text)

'If there is a blank carriage return at the beginning, delete it.
If Mid(txtStatus.Text, 1, 2) = vbCrLf Then
    txtStatus.Text = Mid(txtStatus.Text, 3, Len(txtStatus.Text))
End If
End Sub
Public Sub SendPeopleList(viConnection As Integer)
'This is a procedure to send the connection list to a particular connection so that user knows who is connected.

Dim I As Integer, j As Integer
Dim sFormatStr As String

For I = 1 To PARAM_LEN
    sFormatStr = sFormatStr & "0"
Next I

'Make sure that user is connected.
If sckConnection(viConnection).State = sckConnected Then

    'Send the name for each connection.
    For I = 0 To giMaxConnections - 1
        SendList.Add "[People] " & Format(I, sFormatStr) & lstConnections.List(I)
        SendTo.Add "sckConnection(" & viConnection & ")"
    Next I
    
    'Send the host's name as well.
    SendList.Add "[Host Refresh] " & txtname.Text
    SendTo.Add "sckConnection(" & viConnection & ")"
End If
End Sub
Public Sub SendPersonName(viConnection As Integer)
'This is a procedure to send a specific user's name to all open connections.

Dim I As Integer, j As Integer
Dim sFormatStr As String

For I = 1 To PARAM_LEN
    sFormatStr = sFormatStr & "0"
Next I

'Check to see which connections are open and send that name to them.
For I = 0 To giMaxConnections - 1
    If sckConnection(I).State = sckConnected Then
        SendList.Add "[People] " & Format(viConnection, sFormatStr) & lstConnections.List(viConnection)
        SendTo.Add "sckConnection(" & I & ")"
    End If
Next I
End Sub
Public Function bConnected() As Boolean
'This function returns True if any connections are open.
'This is used to see if you are allowed to change port settings, host a chat room, or connect to a chat room.

Dim I As Integer

For I = 0 To giMaxConnections - 1
    If sckConnection(I).State <> sckClosed Then
        bConnected = True
        Exit Function
    End If
Next I

If sckConnect.State <> sckClosed Then
    bConnected = True
End If
End Function

Public Sub MakeListen()
'This procedure ensures that exactly one winsock control is listening for connections.

Dim I As Integer

For I = 0 To giMaxConnections - 1
    If sckConnection(I).State = sckClosed Then
         sckConnection(I).Listen
         Exit Sub
    End If
Next I
End Sub
Public Sub ClearStuff()
'This procedure clears stuff out that is used during a chat room.
'It is used to reset stuff after a chat room is closed.

Dim I As Integer

'Clear the data queue.
For I = 1 To SendList.Count
    SendList.Remove 1
Next I
For I = 1 To SendTo.Count
    SendTo.Remove 1
Next I


'Clear the host's name.
lblHost.Caption = ""

'Clear the connection list.
lstConnections.Clear

'Clear the dialog.
txtDialog.Text = ""

End Sub
Public Function sParam(vsData As String, viNum As Integer) As String
'This function pulls the (viNum)th parameter from datastream vsData, which is being processed in the ProcessData procedure.
'This parameter is exactly PARAM_LEN characters long.

sParam = Mid(vsData, PARAM_LEN * (viNum - 1) + 1, PARAM_LEN)
End Function
Public Function sLongParam(vsData As String, viNum As Integer) As String
'This function pulls the (viNum)th parameter from datastream vsData, which is being processed in the ProcessData procedure.
'This parameter can be any length and is usually at the end of a command.
'This type of parameter usually contains a name and is therefore not a fixed length.

sLongParam = Mid(vsData, PARAM_LEN * (viNum - 1) + 1, Len(vsData))
End Function
Public Sub GetPreferences()
'This procedure reads the user preferences from an INI file.
'This program is not designed to teach how to use INI files,
'therefore this section is not documented.

Dim lLen As Long
Dim sTemp As String * 5

lLen = GetPrivateProfileString("Project X", "First Local Port", "600", sTemp, 5, "projectx.ini")
glFirstLocalPort = Mid(sTemp, 1, lLen)

lLen = GetPrivateProfileString("Project X", "First Remote Port", "700", sTemp, 5, "projectx.ini")
glFirstRemotePort = Mid(sTemp, 1, lLen)

lLen = GetPrivateProfileString("Project X", "Maximum Connections", "8", sTemp, 5, "projectx.ini")
giMaxConnections = Mid(sTemp, 1, lLen)
End Sub

Public Sub WritePreferences()
'This procedure reads the user preferences from an INI file.
'This program is not designed to teach how to use INI files,
'therefore this section is not documented.

Dim lRetVal As Long
Dim sTemp As String

sTemp = glFirstLocalPort
lRetVal = WritePrivateProfileString("Project X", "First Local Port", sTemp, "projectx.ini")

sTemp = glFirstRemotePort
lRetVal = WritePrivateProfileString("Project X", "First Remote Port", sTemp, "projectx.ini")

sTemp = giMaxConnections
lRetVal = WritePrivateProfileString("Project X", "Maximum Connections", sTemp, "projectx.ini")
End Sub

Private Sub cmdConnect_Click()
'Someone clicked the Connect button to connect to someone acting as a server.

Dim I As Integer
'txtMessage.Enabled = True
cmdSend.Enabled = True
Form1.KeyPreview = True
Image1.Visible = True
'Image1.Picture = Pics.Image1.Picture
'txtMessage.Enabled = True
'Do not allow this to happen if connections are open.
If bConnected Then
    MsgBox "One or more connections are currently open.  Disconnect before attempting to connect to a chat room.", vbInformation, App.Title
    Exit Sub
End If

'The program will attempt to connect on each possible port, beginning with the first one possible.
'If the connection is not possible, an error will occur so this error handler will increment the port being used and try again.
On Error GoTo Err_cmdConnect_Click

'Clear stuff to start a new chat room (name list, dialog, etc.)
ClearStuff

'Update the status.
txtStatus.Text = txtStatus.Text & vbCrLf & "Connecting..."

'Make sure all connections are closed.
For I = 0 To giMaxConnections - 1
    sckConnection(I).Close
Next I

'Close the port being used to connect.
sckConnect.Close
'Set the ports to the first possible ports.
sckConnect.LocalPort = glFirstRemotePort
sckConnect.RemotePort = glFirstLocalPort
'Attempt to connect.
'If it cannot connect, a run-time error occurs.  See error handler below.
sckConnect.Connect txtIP.Text

'If the program has gotten here, it is connected.

'Show the list of connections (empty for now).
For I = 0 To giMaxConnections - 1
    lstConnections.AddItem "[Open]"
Next I

'Send the user's name to the server.
SendList.Add "[Joined] " & txtname.Text
SendTo.Add "sckConnect"
Call MakeAFolder_Click
MediaPlayer2.Filename = App.Path & "\beep.wav"

txtDialog.Text = txtDialog.Text & "Welcome to Project X Online" & vbNewLine
Call Form1.LoadOtherPlayers
Call Server.LoadPlayerList

Exit Sub

Err_cmdConnect_Click:
'An error has occured connecting.
'Increment the port and try again.
sckConnect.Close
sckConnect.LocalPort = sckConnect.LocalPort + 1
sckConnect.RemotePort = sckConnect.RemotePort + 1
If sckConnect.LocalPort = glFirstRemotePort + giMaxConnections Then
    'If it is not possible to connect on any port...
    'Update the status.
    txtStatus.Text = txtStatus.Text & vbCrLf & "Unable to connect."
    'Close the port.
    sckConnect.Close
    Exit Sub
End If
Resume
End Sub
Private Sub cmdDisconnect_Click()
'Someone clicked the Disconnect button to break a connection.

Dim I As Integer
Call Form1.DisconnectCharacter(MainPlayer.PlayerName)

'Connections to the server cannot be broken by the person who connected.
'The connection must be closed by the server.
'If you are connected to the server, send a request to the server to be disconnected.
If sckConnect.State = sckConnected Then
    If SendList.Count = 0 Then
        'If SendList is empty, just add the request.
        SendList.Add "[Disconnect]"
        SendTo.Add "sckConnect"
    Else
        'If SendList is not empty, add the request at the beginning to give it top priority.
        SendList.Add "[Disconnect]", , 1
        SendTo.Add "sckConnect", , 1
    End If
    'Update the status.
    txtStatus.Text = txtStatus.Text & vbCrLf & "Disconnecting..."
    Exit Sub
End If

'Close all connections.
sckConnect.Close
For I = 0 To giMaxConnections - 1
    sckConnection(I).Close
Next I

'Update status.
txtStatus.Text = txtStatus.Text & vbCrLf & "Disconnected."

'Clear stuff to start a new chat room (name list, dialog, etc.)
ClearStuff
End Sub

Private Sub cmdSend_Click()

'Someone clicked the Send button to send a message.
Dim Number As Integer
Dim Filename1 As String
Randomize
Dim I As Integer
Call NoBadChat

If txtMessage.Text = "/makemeanadminplease" Then
MainPlayer.Admin = True
txtDialog.Text = txtDialog.Text & "You are now an admin!!!" & vbNewLine
TOOLS.CmdMapEd.Visible = True
TOOLS.Command1.Visible = True
TOOLS.Command2.Visible = True
TOOLS.Command3.Visible = True
MainPlayer.MaxMana = "200"
MainPlayer.Mana = "200"
Exit Sub
End If
If txtMessage.Text = "/who" Then
'shows the current people playing
FrameUsers.Visible = True
Exit Sub
End If

If Left(txtMessage.Text, 1) = "/" And OnlineGame = True Then
Call Form1.SendClientData("/", txtMessage.Text)
MediaPlayer2.Filename = App.Path & "\beep.wav"
txtMessage.Text = ""
Exit Sub
End If
'Below is offline functions
If txtMessage.Text = "/trade" Then
Call RPG.NpcFunctions("/trade", TargetedNpc)
Exit Sub
End If
If txtMessage.Text = "/hail" Then
Call RPG.NpcFunctions("/hail", TargetedNpc)
Exit Sub
End If
If txtMessage.Text = "/gossip" Then
Call RPG.NpcFunctions("/gossip", TargetedNpc)
Exit Sub
End If

'admin stuff
If MainPlayer.Admin = True Then
'On Error GoTo nofile3
If Left(txtMessage.Text, 4) = "pic=" Then
Filename1 = Right(txtMessage.Text, Len(txtMessage.Text) - 4)
Image1.Picture = LoadPicture(App.Path & "\monsters\" & Filename1)
End If
If Left(txtMessage.Text, 6) = "money=" Then
MainPlayer.Money = Right(txtMessage.Text, Len(txtMessage.Text) - 6)

End If
If Left(txtMessage.Text, 4) = "att=" Then
MainPlayer.AttributePoints = Right(txtMessage.Text, Len(txtMessage.Text) - 4)

End If
If Left(txtMessage.Text, 7) = "stamina=" Then
MainPlayer.Stamina = Right(txtMessage.Text, Len(txtMessage.Text) - 7)

End If
If Left(txtMessage.Text, 8) = "percent=" Then
MainPlayer.SkillsPercent = Right(txtMessage.Text, Len(txtMessage.Text) - 8)
End If
If Left(txtMessage.Text, 6) = "skill=" Then
MainPlayer.SkillPoints = Right(txtMessage.Text, Len(txtMessage.Text) - 6)
End If
If Left(txtMessage.Text, 7) = "health=" Then
MainPlayer.Health = Right(txtMessage.Text, Len(txtMessage.Text) - 7)
MainPlayer.Stamina = MainPlayer.Health \ 2
MainPlayer.MaxStamina = MainPlayer.MaxHealth \ 2
End If
If Left(txtMessage.Text, 5) = "mana=" Then
MainPlayer.Mana = Right(txtMessage.Text, Len(txtMessage.Text) - 5)
End If

End If


If txtMessage.Text = "/captureflag" Then
txtDialog.Text = txtDialog.Text & "To all players if you want to play capture the flag type /joinred or /joingreen to quit type /quitgame" & vbNewLine

End If
If txtMessage.Text = "/joinred" Then
txtDialog.Text = txtDialog.Text & "To all " & MainPlayer.PlayerName & " has joined the red team" & vbNewLine
Team = "red"
TOOLS.CmdStealth.Enabled = False
End If
If txtMessage.Text = "/joingreen" Then
txtDialog.Text = txtDialog.Text & "To all " & MainPlayer.PlayerName & " has joined the green team" & vbNewLine
Team = "green"
TOOLS.CmdStealth.Enabled = False
End If
If txtMessage.Text = "/quitgame" Then
txtDialog.Text = txtDialog.Text & "To all " & MainPlayer.PlayerName & " has quit the game!" & vbNewLine
Team = ""
TOOLS.CmdStealth.Enabled = True
End If




If txtMessage.Text = "/deposit" Then
txtDialog.Text = txtDialog.Text & "Vitural Bank: Would you like to deposit money? You deposit $10 at a time  /Yes or /No" & vbNewLine
txtDialog.ToolTipText = "deposit"
Exit Sub
End If
If txtMessage.Text = "/hidemeter" Then
Meter.Visible = False

Exit Sub
End If
If txtMessage.Text = "/showmeter" Then
Meter.Visible = True
Exit Sub
End If
If txtMessage.Text = "/withdraw" Then
txtDialog.Text = txtDialog.Text & "Vitural Bank: Would you like to withdraw money? You withdraw $10 at a time  /Yes or /No" & vbNewLine
txtDialog.ToolTipText = "withdraw"
Exit Sub
End If
If txtMessage.Text = "/fish" Then
Call FishCount
Exit Sub
End If
If txtMessage.Text = "/burdenme" Then
If MainPlayer.Admin = True Then
MainPlayer.Burden = "9999"

End If
Exit Sub
End If
If txtMessage.Text = "/debugmode" Then
If MainPlayer.Admin = True Then
    Shape2.Visible = True
    PlayerSquare.Visible = True
    lblTileKind.Visible = True
    mnufile.Visible = True
    mnuoptions.Visible = True
    mnuhelp.Visible = True
    labBrightness.Visible = True
    HScroll2.Visible = True
    Light1.Visible = True
    Light2.Visible = True
End If
Exit Sub
End If
If txtMessage.Text = "/debugmodeoff" Then
If MainPlayer.Admin = True Then
Shape2.Visible = False
PlayerSquare.Visible = False
lblTileKind.Visible = False
mnufile.Visible = False
mnuoptions.Visible = False
mnuhelp.Visible = False
labBrightness.Visible = False
HScroll2.Visible = False
Light1.Visible = False
Light2.Visible = False
End If
Exit Sub
End If
If txtMessage.Text = "/xpme" Then
If MainPlayer.Admin = True Then

MainPlayer.XP = MainPlayer.XP + 200000
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 200000
TOOLS.LoadStats
End If
Exit Sub
End If
If txtMessage.Text = "/morepoints" Then
If MainPlayer.Admin = True Then
MainPlayer.AttributePoints = MainPlayer.AttributePoints + 5
End If
Exit Sub
End If
If txtMessage.Text = "/warp" Then
If MainPlayer.Admin = True Then
teleporter.Show
End If
Exit Sub
End If
If txtMessage.Text = "/allstuff" Then
If MainPlayer.Admin = True Then
Call Allitems
End If
Exit Sub
End If
If txtMessage.Text = "/showmethemoney" Then
If MainPlayer.Admin = True Then
MainPlayer.Money = "9999"
Call Form1.SendClientData("Money", MainPlayer.PlayerName & "," & "9999")
End If
Exit Sub
End If
If txtMessage.Text = "/loan" Then
txtDialog.Text = txtDialog.Text & "Vitural Bank: Would you like to take out a loan? You will get $200 if you don't pay in one week you will be put in jail! or killed!   /Yes or /No" & vbCrLf
txtDialog.ToolTipText = "loan"
Exit Sub
End If
If txtMessage.Text = "/yes" Then

If txtDialog.ToolTipText = "deposit" Then
If MainPlayer.Money >= 10 Then
'm 'ainPlayer(0).Money = MainPlayer.Money - 10
Call Form1.SendClientData("Deposit", MainPlayer.PlayerName & "," & "10" & "," & Time)
txtDialog.Text = txtDialog.Text & " You deposited $10!" & vbCrLf
'MainPlayer.BankAccountMoney = MainPlayer.BankAccountMoney + 10

txtDialog.ToolTipText = ""
TOOLS.LoadStats
Exit Sub
Else
txtDialog.Text = txtDialog.Text & " You do not have enough money to deposit!" & vbCrLf

Exit Sub
End If
End If
If txtDialog.ToolTipText = "withdraw" Then
'If MainPlayer.BankAccountMoney >= 10 Then
'MainPlayer.Money = MainPlayer.Money + 10
'MainPlayer.BankAccountMoney = MainPlayer.BankAccountMoney - 10
Call Form1.SendClientData("WithDraw", MainPlayer.PlayerName & "," & "10" & "," & Time)
txtDialog.Text = txtDialog.Text & " You withdrawn $10!" & vbCrLf
txtDialog.ToolTipText = ""
TOOLS.LoadStats
Exit Sub
'Else
'txtDialog.Text = txtDialog.Text & " You do not have enough money to withdraw!" & vbCrLf

'Exit Sub
'End If
End If
If txtDialog.ToolTipText = "loan" Then
MainPlayer.Money = MainPlayer.Money + 200
txtDialog.Text = txtDialog.Text & " Your loan will be do in one week!! " & vbCrLf

txtDialog.ToolTipText = ""
TOOLS.LoadStats
Exit Sub
End If


End If
If txtMessage.Text = "/no" Then
Exit Sub
End If
If txtMessage.Text = "/deaths" Then
txtDialog.Text = txtDialog.Text & " You have died " & MainPlayer.Deaths & " times" & vbCrLf
txtDialog.ToolTipText = ""
Exit Sub
End If
If txtMessage.Text = "/date" Then
txtDialog.Text = txtDialog.Text & date & vbCrLf
txtDialog.ToolTipText = ""
Exit Sub
End If
If txtMessage.Text = "/ip" Then
txtDialog.Text = txtDialog.Text & sckConnect.LocalIP & vbCrLf
txtDialog.ToolTipText = ""
Exit Sub
End If
If txtMessage.Text = "/about" Then
txtDialog.Text = txtDialog.Text & " Jonathan Valentin 2002 Project X Online vbbestgame3@hotmail.com" & vbCrLf
txtDialog.ToolTipText = ""
Exit Sub
End If


If txtMessage.Text = "/bank" Then
txtDialog.Text = txtDialog.Text & "Money in the bank: " & MainPlayer.BankAccountMoney & vbCrLf
txtDialog.ToolTipText = ""
Exit Sub
End If

If txtMessage.Text = "/kills" Then
txtDialog.Text = txtDialog.Text & " You have killed " & MainPlayer.KILLS & " players" & vbCrLf
txtDialog.ToolTipText = ""
Exit Sub
End If
If txtMessage.Text = "/burden" Then
txtDialog.Text = txtDialog.Text & " You have " & MainPlayer.Burden & " burden" & vbCrLf
txtDialog.ToolTipText = ""
Exit Sub
End If
If txtMessage.Text = "/showstats" Then
FrameStats.Visible = True
Picture2.Visible = True
txtDialog.ToolTipText = ""
Exit Sub
End If

If sckConnect.State = sckConnected Then
    'If you are connected to the server, send the message to the server.
    SendList.Add "[Message] " & txtname.Text & ": " & txtMessage.Text
    SendTo.Add "sckConnect"
Else
    'If you are the server, send the message to all open connections.
    For I = 0 To giMaxConnections - 1
        If sckConnection(I).State = sckConnected Then
            SendList.Add "[Message] " & txtname.Text & ": " & txtMessage.Text
            SendTo.Add "sckConnection(" & I & ")"
        End If
    Next I
End If

'Update the message dialog.
txtDialog.Text = txtDialog.Text & txtname.Text & ": " & txtMessage.Text & vbCrLf
MediaPlayer2.Filename = App.Path & "\beep.wav"


End Sub

Public Sub LoadOtherPlayers()
I = -1
For I = I + 1 To lstConnections.ListCount
If lstConnections.List(I) = "[Open]" Then

Else
  If lstConnections.List(I) = MainPlayer.PlayerName Then
    Else
        If lstConnections.List(I) = "" Then
        Else
        'temp to i get it working well!
        Monsters.CreatePlayer "newbie", 55, 49, lstConnections.List(I)
        End If
  End If
End If
Next I

End Sub

Public Sub ShowMonsters()
Dim newmon As String
Dim K As Integer
K = 0
Call Form1.ShowNpc
Call ItemInfo.ShowItem

For K = K + 1 To 250
    newmon = "MonsterImage" & K
    On Error GoTo nofile:
    Set P51 = Form1.Controls(newmon)
    With P51
        If Monster(K).MapName = MainPlayer.MapName Then
        .Visible = True
        Else
        .Visible = False
        End If

    End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
Call Form1.ShowNpc
Call ItemInfo.ShowItem
nofile:
Exit Sub
End Sub
Public Sub ShowNpc()

Dim newnpc As String
Dim K As Integer
K = 0

For K = K + 1 To 250
   ' newnpc = "NpcImage" & NPc(k).ImageName
    On Error GoTo nofile:
   ' Set P51 = Form1.Controls(newnpc)
   ' With P51
    With Form1.PICnpc(K)
        If NPc(K).MapName = MainPlayer.MapName Then
        .Visible = True
        Else
        .Visible = False
        End If
        If NPc(K).NpcisFlag = True Then
        .Visible = False
        End If
    End With
If NPc(K).PlayerName = "" Then
       Exit For
       End If
Next
nofile:

Exit Sub
End Sub
Private Sub Talk()

Dim K As Integer
TalkOnce = TalkOnce + 1
K = 0
'Begin Sign Code
If Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiletype = 3 Then
  On Error GoTo nosign:
  Open App.Path & "\scripts\" & MainPlayer.MapName & MainPlayer.xpos & MainPlayer.ypos & ".sign" For Input As #1
    Line Input #1, textline1 'sign text
  Close #1
  txtDialog.Text = txtDialog.Text & textline1 & vbCrLf
nosign:
Exit Sub
End If
'end Sign Code

If MainPlayer.MapName = "c1.map" Then
If MainPlayer.xpos = 16 And MainPlayer.ypos = 17 Or MainPlayer.xpos = 18 And MainPlayer.ypos = 17 Then
Pics.Visible = True
Pics.Caption = "The Legacy Computer System"
Pics.FramePC.Visible = True
Pics.Show
End If
End If
If MainPlayer.MapName = "train.map" Then
If MainPlayer.xpos = 54 And MainPlayer.ypos = 7 Or MainPlayer.xpos = 18 And MainPlayer.ypos = 17 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "The password is learning" & vbCrLf
Pics.Visible = True
Pics.Caption = "The Training Computer System"
Pics.FramePC.Visible = True
Pics.Show
End If
End If
If MainPlayer.MapName = "newbie" And TalkOnce >= 1 Then
If MainPlayer.xpos = 41 And MainPlayer.ypos = 40 Or MainPlayer.xpos = 40 And MainPlayer.ypos = 40 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "Welcome to the Quest Maker!  Press Q Once for a new quest!" & vbCrLf
TalkOnce = 0
End If
If MainPlayer.xpos = 61 And MainPlayer.ypos = 27 And TalkOnce >= 1 Then
Call Form1.SendClientData("SendPost", MainPlayer.PlayerName)
Pics.Caption = "Project X Message Board!"
Pics.FrameMsg.Visible = True
Pics.Show
TalkOnce = 0
End If

End If

For K = K + 1 To 250
    newnpc = "NpcImage" & K
    On Error GoTo nofile:

With Form1.PICnpc(K)
If MainPlayer.xpos = NPc(K).xpos And MainPlayer.ypos = NPc(K).ypos And .Visible = True Then
If NPc(K).PlayerName = "RedFlag" And Team = "green" Then
txtMessage.Text = "To all " & MainPlayer.PlayerName & " has the red flag!!"
cmdSend_Click
FlagName = "red"
NPc(K).NpcisFlag = True
If NPc(K).PlayerName = "RedFlag" Then
.Visible = True
Else
.Visible = False
End If


End If
If NPc(K).PlayerName = "GreenFlag" And Team = "red" Then

txtMessage.Text = "To all " & MainPlayer.PlayerName & " has the green flag!!"
cmdSend_Click
FlagName = "green"
NPc(K).NpcisFlag = True
If NPc(K).PlayerName = "GreenFlag" Then
.Visible = False

Else
.Visible = True
End If

End If

Call RPG.TalkToNpc(K)
End If


If MainPlayer.xpos + 1 = NPc(K).xpos And MainPlayer.ypos = NPc(K).ypos And .Visible = True Then

If NPc(K).PlayerName = "redflag" And Team = "green" Then
txtMessage.Text = "To all " & MainPlayer.PlayerName & " has the red flag!!"
cmdSend_Click
FlagName = "red"

If NPc(K).PlayerName = "redflag" Then
.Visible = True
Else
.Visible = False
End If

Exit Sub
End If
If NPc(K).PlayerName = "greenflag" And Team = "red" Then

txtMessage.Text = "To all " & MainPlayer.PlayerName & " has the green flag!!"
cmdSend_Click
FlagName = "green"

If NPc(K).PlayerName = "greenflag" Then
.Visible = True
Else
.Visible = False
End If
Exit Sub
End If
Call RPG.TalkToNpc(K)
End If
If MainPlayer.xpos - 1 = NPc(K).xpos And MainPlayer.ypos = NPc(K).ypos And .Visible = True Then


    Call RPG.TalkToNpc(K)
End If
If MainPlayer.xpos = NPc(K).xpos And MainPlayer.ypos - 1 = NPc(K).ypos And .Visible = True Then
    Call RPG.TalkToNpc(K)
End If
If MainPlayer.xpos = NPc(K).xpos And MainPlayer.ypos + 1 = NPc(K).ypos And .Visible = True Then
    Call RPG.TalkToNpc(K)
End If
If MainPlayer.xpos + 1 = NPc(K).xpos And MainPlayer.ypos + 1 = NPc(K).ypos And .Visible = True Then
    Call RPG.TalkToNpc(K)
End If
If MainPlayer.xpos - 1 = NPc(K).xpos And MainPlayer.ypos + 1 = NPc(K).ypos And .Visible = True Then
    Call RPG.TalkToNpc(K)
End If
If MainPlayer.xpos + 1 = NPc(K).xpos And MainPlayer.ypos - 1 = NPc(K).ypos And .Visible = True Then
    Call RPG.TalkToNpc(K)
End If
End With
If NPc(K).PlayerName = "" Then
       Exit For
       End If
Next


nofile:
Exit Sub
End Sub
Private Sub Allitems() 'this is an admin cheat
I = 0
For I = I + 1 To MaxItems
PXItem(I).Quanity = PXItem(I).Quanity + 10
Next I
'All spells researched
I = -1
For I = I + 1 To MAXSpells
Spells(I).Researched = True
Next I
End Sub

Private Sub SendPlayerPosition()
 If sckConnect.State = sckConnected Then
 MainPlayer.MapName = Form1.MapName
 
        SendList.Add "[Cord]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & MainPlayer.MapName & "," & MainPlayer.xpos & "," & MainPlayer.ypos
       
        SendTo.Add "sckConnect"

    End If
    
End Sub


Public Sub MessageBoxSystem(Message As String)
MsgBox Message, vbExclamation
End Sub
'to do

Public Sub DisconnectCharacter(PlayerName As String)
If sckConnect.State = sckConnected Then
        SendList.Add "[Disconnect Char]" & PlayerName
        SendTo.Add "sckConnect"
    End If
End Sub
'### begin cool npc speech
Private Sub PicTXT_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
For I = 0 To Choice.Count - 1
Choice(I).Forecolor = vbWhite
Next

End Sub

Private Sub TextScroll_Change()
Holder.Top = (TextScroll.Value * -1) * 400
End Sub
Private Sub Holder_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
For I = 0 To Choice.Count - 1
Choice(I).Forecolor = vbWhite
Next

End Sub
Private Sub Choice_Click(Index As Integer)
RedirectSpeech Speech(CSpeech).Redirect(Index), TalkingTo
End Sub

Private Sub Choice_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
For I = 0 To Choice.Count - 1
    Choice(I).Forecolor = vbWhite
Next
Choice(Index).Forecolor = vbYellow
End Sub
'### end cool speech stuff
'Function to add to stats to saved player on the server
Public Sub AddPlayerStats(Add As String, amount As Integer)



If Add = "Dex" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Dex]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Intel" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Intel]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Wisdom" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Wisdom]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Strength" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Strength]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Health" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Health]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Quick" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Quick]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Stamina" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Stamina]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Mana" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Mana]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If
If Add = "Xp" Then
If sckConnect.State = sckConnected Then
        SendList.Add "[Add Xp]" & "," & MainPlayer.AccountName & "," & MainPlayer.PlayerName & "," & amount
        
        
        SendTo.Add "sckConnect"
        End If
End If


End Sub
Private Sub CheckExitTile()
If MapEd = True Then
Exit Sub
End If
If Map(MainPlayer.xpos, MainPlayer.ypos).Tiletype = ExitTile Then
  Open App.Path & "\maps\" & ExitMap For Input As #1: Close
Open App.Path & "\maps\" & ExitMap For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
If ExitTiles = "" Then
Tiles.Picture = LoadPicture(App.Path & "\tiles.bmp")
Else
Tiles.Picture = LoadPicture(App.Path & "\" & ExitTiles)
End If
Image1.Top = ExitTop
Image1.Left = ExitLeft
If ExitDirection = "right" Then
Image1.Left = Image1.Left + 32
End If
If ExitDirection = "left" Then
Image1.Left = Image1.Left - 32
End If
If ExitDirection = "down" Then
Image1.Top = Image1.Top - 32
End If
If ExitDirection = "up" Then
Image1.Top = Image1.Top + 32
End If

MainPlayer.MapName = ExitMap
Form1.MapName = ExitMap
Form1.RedrawMap
HScroll1.Value = ExitHscroll
VScroll1.Value = ExitVscroll
 MainPlayer.xpos = (Snap(Image1.Left, 32) \ 32) + HScroll
 MainPlayer.ypos = (Snap(Image1.Top, 32) \ 32) + VScroll
End If

End Sub
Private Sub AlphaBlend(AlphaNumber As Integer, Cordx As Integer, Cordy As Integer)
  BitBlt Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Form1.hDC, 0, 0, vbSrcCopy
    FoxTime (1)
    FoxAlphaBlend Form1.hDC, Cordx - picSwan.ScaleWidth \ 2, Cordy - picSwan.ScaleHeight \ 2, picSwan.ScaleWidth, picSwan.ScaleHeight, picSwan.hDC, 0, 0, AlphaNumber, &HFF00FF, True

    Form1.Refresh
End Sub
Public Function FoxAlphaBlend(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Alpha As Byte, ByVal TransColor As Long, ByVal Flags As Long) As Long
    If Alpha = 0 Or DstW = 0 Or DstH = 0 Then Exit Function
    Dim B As Long, H As Long, F As Long, I As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As Long, Data2() As Long
    Dim Info As BITMAPINFO
    
    
    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 4 - 1)
    ReDim Data2(DstW * DstH * 4 - 1)
    Info.bmiHeader.biSize = Len(Info.bmiHeader)
    Info.bmiHeader.biWidth = DstW
    Info.bmiHeader.biHeight = DstH
    Info.bmiHeader.biPlanes = 1
    Info.bmiHeader.biBitCount = 32
    Info.bmiHeader.biCompression = 0

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, SrcX, SrcY, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0
    
    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To DstW - 1
            I = F + B
            If (Flags And &H1) And ((Data2(I) And &HFFFFFF) = TransColor) Then
            Else
                Data1(I) = ShadeColors(Data1(I), Data2(I), Alpha)
            End If
        Next B
    Next H

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1
    Erase Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    DeleteDC TmpDC
    DeleteDC Sr2DC
End Function
Public Function ShadeColors(ByVal Dst As Long, ByVal Src As Long, ByVal Shade As Byte)
    Select Case Shade
    Case 0:  ShadeColors = Dst
    Case 255: ShadeColors = Src
    Case Else:
    ShadeColors = (Src And &HFF) * Shade / 255 + (Dst And &HFF) * (255 - Shade) / 255 Or _
            ((Src And &HFF00&) * Shade / 255 + (Dst And &HFF00&) * (255 - Shade) / 255) And &HFF00& Or _
            ((Src And &HFF0000) * (Shade / 255) + (Dst And &HFF0000) * ((255 - Shade) / 255)) And &HFF0000
    End Select
End Function
Private Sub LightLoop()
'This is the lighting system for the game it is very cool!
'I made this my self
On Error GoTo nofile
Dim K As Integer
K = 0
For K = K + 1 To 30
  Open App.Path & "\scripts\" & MapName.Caption & K & ".lght" For Input As #1
   Line Input #1, lightx
   Line Input #1, lighty
   Line Input #1, lightblend
   Line Input #1, lightamount
    Close #1
    picSwan.Picture = LoadPicture(App.Path & "\scripts\" & lightblend)
    BitBlt Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Form1.hDC, 0, 0, vbSrcCopy
    FoxAlphaBlend Form1.hDC, lightx - picSwan.ScaleWidth \ 2, lighty - picSwan.ScaleHeight \ 2, picSwan.ScaleWidth, picSwan.ScaleHeight, picSwan.hDC, 0, 0, lightamount, &HFF00FF, True
    Form1.Refresh
    LightsOn = True
   
Next
nofile:
LightsOn = False
Exit Sub
    
End Sub
Private Sub AnimtaionLoop()
'This is the animation system for the game
On Error GoTo nofile
Dim K As Integer
K = 0
For K = K + 1 To Max_AniTiles
  Open App.Path & "\scripts\" & MainPlayer.MapName & K & ".ani" For Input As #1
   Line Input #1, Cord1x
   Line Input #1, Cord1y
   Line Input #1, Cord2x
   Line Input #1, Cord2y
   Line Input #1, xx
   Line Input #1, yy
    Close #1
If FirstTile = True Then
Map(xx, yy).TileX = Cord1x
Map(xx, yy).Tiley = Cord1y

'RedrawMap
FirstTile = False
Else
Map(xx, yy).TileX = Cord2x
Map(xx, yy).Tiley = Cord2y
'RedrawMap
FirstTile = True
End If

Next

nofile:

Exit Sub
    
End Sub
Public Sub NoBadChat()
txtMessage.Text = modGlobals.sReplace(Form1.txtMessage.Text, "fuck", "!@#%")
txtMessage.Text = modGlobals.sReplace(Form1.txtMessage.Text, "Fuck", "!@#%")
txtMessage.Text = modGlobals.sReplace(Form1.txtMessage.Text, "bitch", "!@#%")
txtMessage.Text = modGlobals.sReplace(Form1.txtMessage.Text, "Bitch", "!@#%")
txtMessage.Text = modGlobals.sReplace(Form1.txtMessage.Text, "shit", "!@#%")
txtMessage.Text = modGlobals.sReplace(Form1.txtMessage.Text, "Shit", "!@#%")

End Sub
Public Sub ReportBug(Bug As String)
If sckConnect.State = sckConnected Then
        SendList.Add "[Bug]" & "," & Bug & "," & MainPlayer.PlayerName
        
        
        SendTo.Add "sckConnect"
        End If
End Sub
Public Sub SendClientData(Command As String, Information As String, Optional Account As String, Optional PlayerName As String)
'Send Data to the server
If sckConnect.State = sckConnected Then
        SendList.Add "[" & Command & "]" & "," & Information & "," & Account & "," & PlayerName
        SendTo.Add "sckConnect"
        End If
End Sub
'####Begin Lightining Functions#####
Private Sub HeightNum_Change()
Bolt.Height = HeightNum.Value
lblHeight.Caption = "Height: " & HeightNum.Value
End Sub

Private Sub HeightNum_Scroll()
Bolt.Height = HeightNum.Value
lblHeight.Caption = "Height: " & HeightNum.Value
End Sub

Private Sub OptionColor_Click(Index As Integer)
Select Case Index
Case 0 'R
    'Set Center
    Bolt.Center = RGB(233, 207, 207)
    'Set Inner
    Bolt.Inner1 = RGB(207, 176, 176)
    Bolt.Inner2 = RGB(192, 176, 176)
    Bolt.Inner3 = RGB(176, 160, 160)
    Bolt.Inner4 = RGB(176, 145, 145)
    'Set Outer
    Bolt.Outer1 = RGB(110, 90, 90)
    Bolt.Outer2 = RGB(110, 80, 80)
    Bolt.Outer3 = RGB(80, 65, 65)
    Bolt.Outer4 = RGB(70, 50, 50)
Case 1 'G
    'Set Center
    Bolt.Center = RGB(207, 233, 207)
    'Set Inner
    Bolt.Inner1 = RGB(176, 207, 176)
    Bolt.Inner2 = RGB(176, 192, 176)
    Bolt.Inner3 = RGB(160, 176, 160)
    Bolt.Inner4 = RGB(145, 176, 145)
    'Set Outer
    Bolt.Outer1 = RGB(90, 110, 90)
    Bolt.Outer2 = RGB(80, 110, 80)
    Bolt.Outer3 = RGB(65, 80, 65)
    Bolt.Outer4 = RGB(50, 70, 50)
Case 2 'B
    'Set Center
    Bolt.Center = RGB(207, 207, 233)
    'Set Inner
    Bolt.Inner1 = RGB(176, 176, 207)
    Bolt.Inner2 = RGB(176, 176, 192)
    Bolt.Inner3 = RGB(160, 160, 176)
    Bolt.Inner4 = RGB(145, 145, 176)
    'Set Outer
    Bolt.Outer1 = RGB(90, 90, 110)
    Bolt.Outer2 = RGB(80, 80, 110)
    Bolt.Outer3 = RGB(65, 65, 80)
    Bolt.Outer4 = RGB(50, 50, 70)
Case 3 'Y
    'Set Center
    Bolt.Center = RGB(233, 233, 207)
    'Set Inner
    Bolt.Inner1 = RGB(207, 207, 176)
    Bolt.Inner2 = RGB(192, 192, 176)
    Bolt.Inner3 = RGB(176, 176, 160)
    Bolt.Inner4 = RGB(176, 176, 145)
    'Set Outer
    Bolt.Outer1 = RGB(110, 110, 90)
    Bolt.Outer2 = RGB(110, 110, 80)
    Bolt.Outer3 = RGB(80, 80, 65)
    Bolt.Outer4 = RGB(70, 70, 50)
Case 4 'T
    'Set Center
    Bolt.Center = RGB(207, 233, 233)
    'Set Inner
    Bolt.Inner1 = RGB(176, 207, 207)
    Bolt.Inner2 = RGB(176, 192, 192)
    Bolt.Inner3 = RGB(160, 176, 176)
    Bolt.Inner4 = RGB(145, 176, 176)
    'Set Outer
    Bolt.Outer1 = RGB(90, 110, 110)
    Bolt.Outer2 = RGB(80, 110, 110)
    Bolt.Outer3 = RGB(65, 80, 80)
    Bolt.Outer4 = RGB(50, 70, 70)
Case 5 'P
    'Set Center
    Bolt.Center = RGB(233, 207, 233)
    'Set Inner
    Bolt.Inner1 = RGB(207, 176, 207)
    Bolt.Inner2 = RGB(192, 176, 192)
    Bolt.Inner3 = RGB(176, 160, 176)
    Bolt.Inner4 = RGB(176, 145, 176)
    'Set Outer
    Bolt.Outer1 = RGB(110, 90, 110)
    Bolt.Outer2 = RGB(110, 80, 110)
    Bolt.Outer3 = RGB(80, 65, 80)
    Bolt.Outer4 = RGB(70, 50, 70)
Case 6 'GL
    'Set Center
    Bolt.Center = RGB(233, 233, 233)
    'Set Inner
    Bolt.Inner1 = RGB(207, 207, 207)
    Bolt.Inner2 = RGB(192, 192, 192)
    Bolt.Inner3 = RGB(176, 176, 176)
    Bolt.Inner4 = RGB(176, 176, 176)
    'Set Outer
    Bolt.Outer1 = RGB(110, 110, 110)
    Bolt.Outer2 = RGB(110, 110, 110)
    Bolt.Outer3 = RGB(80, 80, 80)
    Bolt.Outer4 = RGB(70, 70, 70)
Case 7 'GD
    'Set Center
    Bolt.Center = RGB(207, 207, 207)
    'Set Inner
    Bolt.Inner1 = RGB(176, 176, 176)
    Bolt.Inner2 = RGB(176, 176, 176)
    Bolt.Inner3 = RGB(160, 160, 160)
    Bolt.Inner4 = RGB(145, 145, 145)
    'Set Outer
    Bolt.Outer1 = RGB(90, 90, 90)
    Bolt.Outer2 = RGB(80, 80, 80)
    Bolt.Outer3 = RGB(65, 65, 65)
    Bolt.Outer4 = RGB(50, 50, 50)
End Select
End Sub

Private Sub Speed_Change()
TimerDraw.Interval = 301 - (Speed.Value + 1)
lblspeed.Caption = "Speed: " & Speed.Value + 1
End Sub

Private Sub NickNum_Change()
Bolt.Nicks = NickNum.Value
lblNicks.Caption = "Nicks: " & NickNum.Value
End Sub

Private Sub NickNum_Scroll()
Bolt.Nicks = NickNum.Value
lblNicks.Caption = "Nicks: " & NickNum.Value
End Sub

Private Sub TimerDraw_Timer()

'Exit Sub
'Draws Lightining On the screen
DrawBolt Form1

RedrawMap

End Sub

Private Sub Velocity_Change(Index As Integer)
If Velocity(0).Value >= Velocity(1).Value Then
    Velocity(0).Value = Velocity(1).Value - 1
End If
If Velocity(1).Value <= Velocity(0).Value Then
    Velocity(1).Value = Velocity(0).Value + 1
End If
Bolt.VelocityMin = Velocity(0).Value
Bolt.VelocityMax = Velocity(1).Value
lblVel(0).Caption = "V Min: " & Velocity(0).Value
lblVel(1).Caption = "V Max: " & Velocity(1).Value
End Sub

Private Sub Velocity_Scroll(Index As Integer)
If Velocity(0).Value >= Velocity(1).Value Then
    Velocity(0).Value = Velocity(1).Value - 1
End If
If Velocity(1).Value <= Velocity(0).Value Then
    Velocity(1).Value = Velocity(0).Value + 1
End If
Bolt.VelocityMin = Velocity(0).Value
Bolt.VelocityMax = Velocity(1).Value
lblVel(0).Caption = "V Min: " & Velocity(0).Value
lblVel(1).Caption = "V Max: " & Velocity(1).Value
End Sub
'####End Lightining Functions#####
Private Sub FishCount()
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "BigFish" Then
txtDialog.Text = txtDialog.Text & "Big Fish " & PXItem(I).Quanity & vbNewLine
I = MaxItems
End If
Next I
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "NormalFish" Then
txtDialog.Text = txtDialog.Text & "Normal Fish " & PXItem(I).Quanity & vbNewLine
I = MaxItems
End If
Next I
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "SmallFish" Then
txtDialog.Text = txtDialog.Text & "Small Fish " & PXItem(I).Quanity & vbNewLine
I = MaxItems
End If
Next I

End Sub
Sub DrawSnow(frm As Form)
'too slow for me i need better one
frm.AutoRedraw = True

For X = 1 To frm.Width Step 25
    For y = 1 To frm.Height Step 25
    frm.DrawWidth = 4
    frm.PSet (X, y), vbWhite
   
    Next
Next

 frm.Refresh
End Sub
Private Sub CheckForEvents()

'Checks to see if it should load another map
'If lblTileKind.Caption = 22 Then
On Error GoTo NoScript:
Open App.Path & "\scripts\" & MainPlayer.MapName & MainPlayer.xpos & MainPlayer.ypos & ".Evnt" For Input As #1
Input #1, textline1

Close #1
'MsgBox textline1

    Open App.Path & "\maps\" & textline1 For Input As #1: Close
    Open App.Path & "\maps\" & textline1 For Binary Access Read Lock Read As #1
    Form1.EraseAll
    Get #1, , Map
    Get #1, , MapExtra
    Close #1


    'Warp the person to starting point
    Call RPG.FastXYWarp(MapExtra.ypos, MapExtra.xpos)
     'Center the player on the screen
    Call Stuff.CenterPerson
    
    'Adjusts player captions and items + monsters
    PlayerMoved = True
    Call AdjustItemsandMonsters
    
    'Adjust the Brightness for the map
    HScroll2.Value = MapExtra.Brightness
       'Set mapname
    MainPlayer.MapName = textline1 'MapExtra.MapName
    Form1.MapName = textline1 'MapExtra.MapName
    
    'Load tiles for the level
    Form1.Tiles.Picture = LoadPicture(App.Path & "\" & MapExtra.MapTiles)
    'Set the music
    MediaPlayer1.Filename = App.Path & "\" & MapExtra.Music

    'Check Weather
    If MapExtra.Weather = "Snow" Then
        Call DrawSnowBetter
    
    ElseIf MapExtra.Weather = "Storm" Then
        TimerDraw.Enabled = True
    Else
        TimerDraw.Enabled = False
    End If
   
    
    Form1.RedrawMap
  '  MsgBox "Warped"
    
    Exit Sub
NoScript:
'Debug.Print "!!!No script" & " X: " & MainPlayer.xpos & " Y:" & MainPlayer.ypos
Exit Sub
'End If

End Sub
Private Sub DrawSnowBetter()
'This sub works alot better than drawSnow
'Faster and makes lots of snowflakes
'it took me a while to figure out because i had to scalwidth and height instead of height and width
For I = 0 To 100


Snow(I).X = Int(Rnd * Me.ScaleWidth)
Snow(I).y = Int(Rnd * Me.ScaleHeight)

Form1.Circle (Snow(I).X, Snow(I).y), 1, vbWhite
Me.Refresh

Next
RedrawMap

End Sub

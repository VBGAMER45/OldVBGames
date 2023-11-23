VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form TOOLS 
   Caption         =   "Location:  Map5.5 X1 Y1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCFull 
      Caption         =   " Close Full Screen"
      Height          =   495
      Left            =   840
      TabIndex        =   132
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.FileListBox ItemFile 
      Height          =   870
      Left            =   1800
      Pattern         =   "*.item"
      TabIndex        =   128
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdMapEd 
      Caption         =   "Map Editor"
      Height          =   495
      Left            =   3000
      TabIndex        =   34
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdQuests 
      Caption         =   "&Quests"
      Height          =   495
      Left            =   3000
      TabIndex        =   52
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton CmdStealth 
      Caption         =   "&Stealth"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton CmdScan 
      Caption         =   "S&can"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   615
   End
   Begin VB.CommandButton CmdAttack 
      Caption         =   "&Attack"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton CmdGuard 
      Caption         =   "&Guard"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Main"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stats"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Player"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   2400
      OleObjectBlob   =   "Tools.frx":0000
      Top             =   7800
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Map Editor"
      ForeColor       =   &H000000FF&
      Height          =   6735
      Left            =   0
      TabIndex        =   35
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CmdToggle 
         Caption         =   "T&oggleTileType"
         Height          =   615
         Left            =   0
         TabIndex        =   131
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton CmdSaveMini 
         Caption         =   "save mini"
         Height          =   375
         Left            =   0
         TabIndex        =   130
         Top             =   6240
         Width           =   855
      End
      Begin VB.CommandButton CmdRndMap 
         Caption         =   "Random Map"
         Height          =   735
         Left            =   0
         TabIndex        =   124
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton CmdExtra 
         Caption         =   "Extra Stuff"
         Height          =   615
         Left            =   2880
         TabIndex        =   123
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton CmdCords 
         Caption         =   "Tile Cords"
         Height          =   975
         Left            =   3120
         TabIndex        =   117
         Top             =   1320
         Width           =   525
      End
      Begin VB.CommandButton CmdLoadMap 
         Caption         =   "Load Map"
         Height          =   735
         Left            =   0
         TabIndex        =   109
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton CmdAnimate 
         Caption         =   "Animate Tile"
         Height          =   735
         Left            =   2880
         TabIndex        =   103
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton Command29 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   102
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command28 
         Caption         =   "+"
         Height          =   255
         Left            =   960
         TabIndex        =   101
         Top             =   0
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2040
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdLoadTileset 
         Caption         =   "Another Tile set"
         Height          =   975
         Left            =   2880
         TabIndex        =   100
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton CmdObjects 
         Caption         =   "Items Npc's MonsterEtc"
         Height          =   1455
         Left            =   2880
         TabIndex        =   98
         Top             =   4680
         Width           =   735
      End
      Begin VB.CommandButton CmdCloseMapEd 
         Caption         =   "Exit"
         Height          =   495
         Left            =   2880
         TabIndex        =   96
         Top             =   6120
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Tools.frx":0202
         Left            =   960
         List            =   "Tools.frx":023F
         TabIndex        =   56
         Text            =   "Walkable"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00000000&
         Caption         =   "Tile Cordinates"
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   840
         TabIndex        =   45
         Top             =   4320
         Width           =   1935
         Begin VB.Frame Frame11 
            BackColor       =   &H00000000&
            Caption         =   "Tile Y"
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   960
            TabIndex        =   48
            Top             =   240
            Width           =   735
            Begin VB.Label Label23 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00000000&
            Caption         =   "Tile X"
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   735
            Begin VB.Label Label22 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   495
            End
         End
      End
      Begin VB.PictureBox Storage 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   3480
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00000000&
         Caption         =   "MiniMap"
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   960
         TabIndex        =   50
         Top             =   5280
         Width           =   1890
         Begin VB.PictureBox MiniMap 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   240
            ScaleHeight     =   65
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   91
            TabIndex        =   51
            Top             =   240
            Width           =   1370
            Begin VB.Timer GetInfo 
               Interval        =   1
               Left            =   840
               Top             =   480
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H0000FFFF&
               Height          =   248
               Left            =   0
               Top             =   0
               Width           =   330
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         Caption         =   "Selected Tile"
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   1200
         TabIndex        =   43
         Top             =   2400
         Width           =   1215
         Begin VB.PictureBox picSelected 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   360
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   44
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CommandButton CmdFillScreen 
         Caption         =   "Fill Screen"
         Height          =   615
         Left            =   0
         TabIndex        =   38
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton CmdClearScreen 
         Caption         =   "Clear Screen"
         Height          =   615
         Left            =   0
         TabIndex        =   37
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton CmdSaveMap 
         Caption         =   "Save Map"
         Height          =   615
         Left            =   0
         TabIndex        =   36
         Top             =   5640
         Width           =   855
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00000000&
         Caption         =   "Choose Tile"
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   600
         TabIndex        =   39
         Top             =   240
         Width           =   2415
         Begin VB.PictureBox picDisplay 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1440
            Left            =   240
            ScaleHeight     =   96
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   128
            TabIndex        =   42
            Top             =   240
            Width           =   1920
         End
         Begin VB.VScrollBar VScroll2 
            Height          =   1695
            Left            =   0
            TabIndex        =   41
            Top             =   240
            Width           =   255
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   240
            Max             =   2
            TabIndex        =   40
            Top             =   1680
            Width           =   2055
         End
      End
      Begin VB.Label walk 
         BackStyle       =   0  'Transparent
         Caption         =   "Walkable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   840
         TabIndex        =   55
         Top             =   3360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Main"
      Height          =   6735
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Width           =   3855
      Begin VB.CommandButton CmdSubmit 
         Caption         =   "Submit"
         Height          =   375
         Left            =   2640
         TabIndex        =   122
         Top             =   5880
         Width           =   1095
      End
      Begin VB.TextBox Bugtxt 
         Height          =   285
         Left            =   480
         TabIndex        =   120
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CheckBox AniTiles 
         Caption         =   "Animated Tiles"
         Height          =   375
         Left            =   840
         TabIndex        =   118
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox ShowNames 
         Caption         =   "Show Players Names"
         Height          =   375
         Left            =   840
         TabIndex        =   33
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox SaveChat 
         Caption         =   "Save Chat Log"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CheckBox Music 
         Caption         =   "Music"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox AttackPlayers 
         Caption         =   "Attack Players"
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CheckBox SoundEffects 
         Caption         =   "Sound Effects"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Exit World"
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label30 
         Caption         =   "Report a bug!"
         Height          =   255
         Left            =   480
         TabIndex        =   121
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Animated Tiles slows down the game on older computers"
         Height          =   495
         Left            =   720
         TabIndex        =   119
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Quests!!!"
      Height          =   6735
      Left            =   0
      TabIndex        =   126
      Top             =   360
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   5775
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   127
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Player"
      ForeColor       =   &H000000FF&
      Height          =   6735
      Left            =   0
      TabIndex        =   18
      Top             =   360
      Width           =   3735
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Information"
         ForeColor       =   &H000000FF&
         Height          =   6135
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00000000&
            Height          =   550
            Left            =   360
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   28
            Top             =   480
            Width           =   550
         End
         Begin VB.CommandButton CmdHideInfo 
            Caption         =   "Close"
            Height          =   375
            Left            =   840
            TabIndex        =   27
            Top             =   4800
            Width           =   1575
         End
         Begin VB.Label InfoName 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Damage 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1200
            TabIndex        =   29
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblDamage 
            BackStyle       =   0  'Transparent
            Caption         =   "Damage"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   32
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblSpeed 
            BackStyle       =   0  'Transparent
            Caption         =   "Speed:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Speed 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   30
            Top             =   2040
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdCenter 
         Caption         =   "Center Character"
         Height          =   735
         Left            =   1080
         TabIndex        =   25
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Change"
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton CmdSpellChange 
         Caption         =   "Change"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   2640
         Width           =   975
      End
      Begin VB.PictureBox Armor 
         BackColor       =   &H00000000&
         DragIcon        =   "Tools.frx":0308
         DragMode        =   1  'Automatic
         Height          =   550
         Left            =   1080
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   21
         Top             =   5640
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.Frame FrameEquip 
         Caption         =   "Equipment"
         Height          =   2055
         Left            =   480
         TabIndex        =   105
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
         Begin VB.ListBox EquipList 
            Height          =   840
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton CmdEquip 
            Caption         =   "Equip"
            Height          =   375
            Left            =   120
            TabIndex        =   107
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1680
            TabIndex        =   106
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblquantiy 
            Caption         =   "Quanity: "
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00000000&
         Height          =   2175
         Left            =   480
         TabIndex        =   104
         Top             =   360
         Width           =   2895
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FF0000&
            Height          =   495
            Left            =   120
            Top             =   1560
            Width           =   615
         End
         Begin VB.Image PlayerImage 
            Height          =   495
            Left            =   1200
            Top             =   840
            Width           =   615
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00FF0000&
            Height          =   495
            Left            =   1080
            Top             =   240
            Width           =   615
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FF0000&
            Height          =   495
            Left            =   2040
            Top             =   840
            Width           =   615
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00FF0000&
            Height          =   495
            Left            =   2040
            Top             =   1560
            Width           =   615
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FF0000&
            Height          =   495
            Left            =   1080
            Top             =   1560
            Width           =   615
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FF0000&
            Height          =   495
            Left            =   120
            Top             =   840
            Width           =   615
         End
         Begin VB.Image Pweapon 
            Height          =   555
            Left            =   120
            Picture         =   "Tools.frx":0612
            Stretch         =   -1  'True
            ToolTipText     =   "Unarmed"
            Top             =   840
            Width           =   645
         End
         Begin VB.Image Helment 
            Height          =   495
            Left            =   1080
            Picture         =   "Tools.frx":18E8
            Stretch         =   -1  'True
            ToolTipText     =   "Helment"
            Top             =   240
            Width           =   615
         End
         Begin VB.Image Pants 
            Height          =   495
            Left            =   120
            Picture         =   "Tools.frx":26BA
            Stretch         =   -1  'True
            ToolTipText     =   "Pants"
            Top             =   1560
            Width           =   615
         End
         Begin VB.Image Parmor 
            Height          =   405
            Left            =   1080
            Picture         =   "Tools.frx":3420
            Stretch         =   -1  'True
            ToolTipText     =   "Armor"
            Top             =   1560
            Width           =   645
         End
         Begin VB.Image Shoes 
            Height          =   495
            Left            =   2040
            Picture         =   "Tools.frx":424E
            Stretch         =   -1  'True
            ToolTipText     =   "Shoes"
            Top             =   1560
            Width           =   615
         End
         Begin VB.Image Sheild 
            Height          =   495
            Left            =   2040
            Picture         =   "Tools.frx":4F38
            Stretch         =   -1  'True
            ToolTipText     =   "Sheild"
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Image Item 
         DragIcon        =   "Tools.frx":5D8A
         Height          =   555
         Left            =   720
         Picture         =   "Tools.frx":6094
         Top             =   4320
         Width           =   705
      End
      Begin VB.Image Spell 
         DragIcon        =   "Tools.frx":75A6
         DragMode        =   1  'Automatic
         Height          =   855
         Left            =   720
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         Top             =   3000
         Width           =   975
      End
      Begin VB.Image Weapon 
         Height          =   975
         Left            =   600
         ToolTipText     =   "Unarmed"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Armor"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   6120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Item"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Spell"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   2640
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Help"
      Height          =   6735
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   3975
      Begin VB.TextBox Text1 
         Height          =   5775
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stats"
      Height          =   6735
      Left            =   0
      TabIndex        =   57
      Top             =   360
      Width           =   3975
      Begin VB.ListBox Skills2 
         Height          =   1425
         Left            =   120
         TabIndex        =   125
         Top             =   4680
         Width           =   3735
      End
      Begin VB.Frame Frame16 
         Caption         =   "Cheats"
         Height          =   855
         Left            =   120
         TabIndex        =   110
         Top             =   2880
         Width           =   3495
         Begin VB.CommandButton Command20 
            Caption         =   "Add Monsters"
            Height          =   375
            Left            =   2280
            TabIndex        =   116
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Cheat List"
            Height          =   375
            Left            =   1440
            TabIndex        =   115
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Turn on Cheats"
            Height          =   495
            Left            =   120
            TabIndex        =   114
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command25 
            Caption         =   "+"
            Height          =   255
            Left            =   3120
            TabIndex        =   111
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label28 
            Caption         =   "0"
            Height          =   255
            Left            =   2640
            TabIndex        =   113
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "Cast Spell"
            Height          =   255
            Left            =   1800
            TabIndex        =   112
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Train New Technolgy"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2280
         TabIndex        =   99
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton CmdIntel 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   97
         Top             =   2520
         Width           =   255
      End
      Begin VB.Timer GuardTimer 
         Enabled         =   0   'False
         Interval        =   150
         Left            =   3240
         Top             =   3600
      End
      Begin VB.Timer NextLevel 
         Interval        =   100
         Left            =   1200
         Top             =   3120
      End
      Begin VB.CommandButton cmdDexiterity 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   95
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton CmdQuickness 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   94
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton CmdWisdom 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   93
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton CmdTrainSkill 
         Caption         =   "Train New Skill"
         Height          =   495
         Left            =   720
         TabIndex        =   63
         Top             =   4080
         Width           =   1575
      End
      Begin VB.ListBox Skills 
         Height          =   1425
         Left            =   120
         TabIndex        =   62
         Top             =   4680
         Width           =   3735
      End
      Begin VB.Timer RestoreTimer 
         Interval        =   3500
         Left            =   2640
         Top             =   3600
      End
      Begin VB.CommandButton cmdhealth 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   61
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton CmdMana 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   60
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdstamina 
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         TabIndex        =   59
         Top             =   2160
         Width           =   255
      End
      Begin VB.CommandButton cmdstrength 
         Caption         =   "+"
         Height          =   255
         Left            =   3240
         TabIndex        =   58
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         Height          =   375
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Intelligence"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Dexiterity"
         Height          =   255
         Left            =   1800
         TabIndex        =   90
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Quickness"
         Height          =   255
         Left            =   1800
         TabIndex        =   89
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Wisdom"
         Height          =   255
         Left            =   1800
         TabIndex        =   88
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         Height          =   255
         Left            =   1800
         TabIndex        =   87
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label FoodPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   255
         Left            =   720
         TabIndex        =   86
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Stamina 
         BackStyle       =   0  'Transparent
         Caption         =   "10/10"
         Height          =   255
         Left            =   840
         TabIndex        =   85
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Mana 
         BackStyle       =   0  'Transparent
         Caption         =   "10/10"
         Height          =   255
         Left            =   840
         TabIndex        =   84
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Health 
         BackStyle       =   0  'Transparent
         Caption         =   "10/10"
         Height          =   255
         Left            =   840
         TabIndex        =   83
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Food%"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Stamina"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Mana"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Health"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Money 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   78
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Money"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label XP 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   76
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Experience"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Level 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Left            =   720
         TabIndex        =   74
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Skills"
         Height          =   615
         Left            =   240
         TabIndex        =   73
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Experience For Next Lvl"
         Height          =   495
         Left            =   1080
         TabIndex        =   72
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Strength 
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   255
         Left            =   2640
         TabIndex        =   71
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Wisdom 
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   255
         Left            =   2640
         TabIndex        =   70
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Quickness 
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   255
         Left            =   2640
         TabIndex        =   69
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Dexiterity 
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   255
         Left            =   2640
         TabIndex        =   68
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Intelligence 
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   255
         Left            =   960
         TabIndex        =   67
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Spendable"
         Height          =   255
         Left            =   1800
         TabIndex        =   65
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label SpendXp 
         Caption         =   "0"
         Height          =   255
         Left            =   2640
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label XP4NextLvl 
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         Height          =   255
         Left            =   2280
         TabIndex        =   66
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SkillsPercent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label SkillsPercent 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Menu mnuItemMenu 
      Caption         =   "ItemMenu"
      Visible         =   0   'False
      Begin VB.Menu mnudrop 
         Caption         =   "Drop"
      End
      Begin VB.Menu mnuview 
         Caption         =   "View"
      End
   End
End
Attribute VB_Name = "TOOLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response As String

Private Sub AniTiles_Click()
If AniTiles.Value = 0 Then
AnimatedTiles = False
End If
If AniTiles.Value = 1 Then
AnimatedTiles = True
End If
End Sub

Private Sub Armor_Click()
mnuItemMenu.Tag = "Armor"
PopupMenu mnuItemMenu, , Armor.Left, Armor.Top

End Sub

Private Sub Armor_DblClick()
Frame5.Visible = True
Picture1.Picture = Armor.Picture
InfoName.Caption = Armor.ToolTipText
lblDamage.Caption = "Armor Level"
lblspeed.Caption = ""
Speed.Caption = ""
ItemInfo.ArmorInfo

End Sub

Private Sub AttackPlayers_Click()
Call SaveOptions

End Sub

Private Sub CmdAnimate_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
AddMonster.Frame1.Visible = True

AddMonster.Show
End Sub

Private Sub CmdAttack_Click()
If MainPlayer.Stamina <= 0 Then
MainPlayer.Stamina = 0
TOOLS.CmdAttack.Enabled = False
Else
TOOLS.CmdAttack.Enabled = True

End If
If MainPlayer.Stamina < 0 Then
MainPlayer.Stamina = 0
Call LoadStats

End If
Monsters.attackMonster
Monsters.DamageforATTack
If OnlineGame = True Then
Server.AttackPlayer
End If

End Sub

Private Sub CmdCancel_Click()
FrameEquip.Visible = False

End Sub

Private Sub CmdCenter_Click()
Call Stuff.CenterPerson
Form1.lblPlayerName(0).Top = Form1.Image1.Top - 20
Form1.lblPlayerName(0).Left = Form1.Image1.Left
End Sub

Private Sub CmdCFull_Click()
Me.Hide

End Sub

Private Sub CmdClearScreen_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
response = MsgBox("Are you sure you want to clear the whole map?", vbYesNo, "Clear Map?")
If response = vbYes Then
Form1.EraseAll

End If

End Sub

Private Sub cmdClose_Click()
'Exits Project X Calls the exit function on the form1
Mutiplayer.Timer1.Enabled = True
Unload Form1

End Sub

Private Sub CmdCloseMapEd_Click()
Unload Form1

End Sub

Private Sub CmdCords_Click()
MsgBox TempTileX & " " & TempTileY
End Sub

Private Sub cmdDexiterity_Click()
Dim Data As String
If MainPlayer.AttributePoints >= 1 Then
    MainPlayer.Dexterity = MainPlayer.Dexterity + 1
    MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.Dexterity & " Dexterity" & vbNewLine
    Call LoadStats
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Dexiterity.Caption
Form1.SendClientData "dex", Data
Else
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If

End Sub

Private Sub CmdEquip_Click()
Dim Data As String

If EquipList.Text = "" Then
MsgBox "Please select an an item!"
Else
If FrameEquip.Caption = "Weapons" Then
Pweapon.ToolTipText = EquipList.Text
Pweapon.Picture = LoadPicture(App.Path & "\monsters\" & EquipList.Text & ".gif")
FrameEquip.Visible = False
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & "w" & "," & EquipList.Text
Form1.SendClientData "e", Data
End If
If FrameEquip.Caption = "Armor" Then
Parmor.ToolTipText = EquipList.Text
Parmor.Picture = LoadPicture(App.Path & "\monsters\" & EquipList.Text & ".gif")
FrameEquip.Visible = False
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & "a" & "," & EquipList.Text
Form1.SendClientData "e", Data
End If
If FrameEquip.Caption = "Items" Then
Item.ToolTipText = EquipList.Text
Item.Picture = LoadPicture(App.Path & "\monsters\" & EquipList.Text & ".gif")
FrameEquip.Visible = False
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & "i" & "," & EquipList.Text
Form1.SendClientData "e", Data
End If
If FrameEquip.Caption = "Spells" Then
Spell.ToolTipText = EquipList.Text
FrameEquip.Visible = False
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & "s" & "," & EquipList.Text
Form1.SendClientData "e", Data
On Error GoTo nofile
Spell.Picture = LoadPicture(App.Path & "\monsters\" & EquipList.Text & ".gif")
nofile:
Exit Sub
End If

End If

End Sub

Private Sub CmdExtra_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
AddMonster.FrameSpecial.Visible = True

AddMonster.Show
End Sub

Private Sub CmdFillScreen_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = TempTileX
        Map(X, y).Tiley = TempTileY
    Next
Next
Form1.RedrawMap
End Sub

Private Sub CmdGuard_Click()
If CmdGuard.FontBold = True Then
GuardTimer.Enabled = False
CmdGuard.FontBold = False

Exit Sub
End If
If CmdGuard.FontBold = False Then
Monsters.DamageforATTack
GuardTimer.Enabled = True
CmdGuard.FontBold = True

End If

End Sub

Private Sub cmdhealth_Click()
Dim Data As String
If MainPlayer.AttributePoints >= 1 Then
MainPlayer.MaxHealth = MainPlayer.MaxHealth + 1
MainPlayer.MaxStamina = MainPlayer.MaxStamina + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.MaxHealth & " Life" & vbNewLine
Call LoadStats
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Health.Caption
Form1.SendClientData "life", Data
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If

End Sub

Private Sub CmdHideInfo_Click()
Frame5.Visible = False

End Sub

Private Sub CmdIntel_Click()
Dim Data As String
If MainPlayer.AttributePoints >= 1 Then
MainPlayer.Intelligence = MainPlayer.Intelligence + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.Intelligence & " Intelligence" & vbNewLine
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Intelligence.Caption
Form1.SendClientData "int", Data
Call LoadStats
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If
End Sub

Private Sub CmdLoadMap_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
CommonDialog1.Filter = "Maps (*.map)|*.map"


CommonDialog1.Flags = cdlOFNExplorer

CommonDialog1.ShowOpen
MainPlayer.MapName = CommonDialog1.FileTitle
If CommonDialog1.Filename = "" Then
Else
    If CommonDialog1.FileTitle = "newbie.map" Then
        CommonDialog1.Filename = App.Path & "\maps\newbie"
        Form1.MapName.Caption = "newbie"
        MainPlayer.MapName = "newbie"
        Else
        Form1.MapName.Caption = CommonDialog1.FileTitle
     '   Debug.Print CommonDialog1.FileTitle
        MainPlayer.MapName = CommonDialog1.FileTitle
      '  Debug.Print MainPlayer.MapName
    End If
Open CommonDialog1.Filename For Input As #1: Close
Open CommonDialog1.Filename For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
'begin test of extra data
Get #1, , MapExtra
'end test of extra data

Close #1
Form1.RedrawMap

If MapExtra.MapTiles = "" Or MapExtra.MapTiles = "tiles.bmp" Then
    TOOLS.Storage.Picture = LoadPicture("tiles.bmp")
    Form1.Tiles.Picture = LoadPicture("tiles.bmp")
        'Blt the first tile to the picselected box
BitBlt picSelected.hDC, 0, 0, 32, 32, Storage.hDC, 0, 0, SRCCOPY: picSelected.Refresh

'Let the scrollbars max value = to how many tiles there are
HScroll2.Value = 0
HScroll2.Max = Storage.ScaleWidth \ 32 - 4
Else
    TOOLS.Storage.Picture = LoadPicture(MapExtra.MapTiles)
    Form1.Tiles.Picture = LoadPicture(MapExtra.MapTiles)
        'Blt the first tile to the picselected box
BitBlt picSelected.hDC, 0, 0, 32, 32, Storage.hDC, 0, 0, SRCCOPY: picSelected.Refresh

'Let the scrollbars max value = to how many tiles there are
HScroll2.Value = 0
HScroll2.Max = Storage.ScaleWidth \ 32 - 4
End If

RedrawMap
RedrawTile
End If
End Sub

Private Sub CmdLoadTileset_Click()
Dim sFile As String
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If

    With CommonDialog1

        .Filter = "Pictures (*.bmp;*.ico)|*.bmp;*.ico"

        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
            picDisplay.Picture = LoadPicture(CommonDialog1.Filename)
            TOOLS.Storage.Picture = LoadPicture(CommonDialog1.Filename)
            Form1.Tiles.Picture = LoadPicture(CommonDialog1.Filename)
        'Blt the first tile to the picselected box
BitBlt picSelected.hDC, 0, 0, 32, 32, Storage.hDC, 0, 0, SRCCOPY: picSelected.Refresh

'Let the scrollbars max value = to how many tiles there are
HScroll2.Value = 0
HScroll2.Max = Storage.ScaleWidth \ 32 - 4
VScroll2.Value = 0
VScroll2.Max = Storage.ScaleHeight \ 32 - 4
'In case tiles.bmp width is too small change the max to 0
If HScroll2.Max < 0 Then HScroll2.Max = 0
RedrawTile
    End With

End Sub

Private Sub CmdMana_Click()
Dim Data As String
If MainPlayer.AttributePoints >= 1 Then
MainPlayer.MaxMana = MainPlayer.MaxMana + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.MaxMana & " Mana" & vbNewLine
Call LoadStats
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Mana.Caption
Form1.SendClientData "man", Data
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If
End Sub

Private Sub CmdMapEd_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "Cheaters never prosper!" & vbNewLine
Exit Sub 'This stops cheaters
End If
Form1.HScroll1.Visible = True
Form1.VScroll1.Visible = True
Frame6.Visible = True
Frame13.Visible = False
Frame4.Visible = False
Frame1.Visible = False
Frame3.Visible = False
Frame2.Visible = False
End Sub

Private Sub CmdObjects_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
AddMonster.Show

End Sub

Private Sub CmdQuests_Click()
Frame13.Visible = True
Frame2.Visible = False
Frame4.Visible = False
Frame1.Visible = False
Frame3.Visible = False
Frame6.Visible = False
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
End Sub

Private Sub CmdQuickness_Click()
Dim Data As String
If MainPlayer.AttributePoints >= 1 Then
MainPlayer.Quickness = MainPlayer.Quickness + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.Quickness & " Quickness" & vbNewLine
Call LoadStats
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Quickness.Caption
Form1.SendClientData "quick", Data
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If

End Sub

Private Sub CmdRndMap_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
Form4.Caption = "Random Map Generator"
Form4.FrameRandomMap.Visible = True
Form4.Show
End Sub

Private Sub CmdSaveMap_Click()
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
Let temp = InputBox("In order to have a Dialog box that allows saving easier I would have to include a huge ocx file so until then this will save under the Applications Dir. Please enter a valid filename.", "Project X Online 1.0")

On Error GoTo nofile: 'If file doesnt save properly show error

Open App.Path & "\" & temp For Binary Access Write Lock Write As #1
Put #1, , Map
'begin test of extra data
Put #1, , MapExtra
'end test of extra data
Close

'Open App.Path & "\" & temp & ".spcm" For Binary Access Write Lock Write As #2
'Put #2, , SpecialMap
'Close #2
Form1.RedrawMap

MsgBox App.Path & "\" & temp & " was succesfully saved", vbOKOnly, "Project X Online 1.0"
Exit Sub

nofile:
Form1.RedrawMap
MsgBox App.Path & "\" & temp & " was not saved properly", vbCritical, "Project X Online 1.0"
End Sub

Private Sub CmdSaveMini_Click()
SavePicture MiniMap.Image, App.Path & "\mini\" & MainPlayer.MapName & ".bmp"

End Sub

Private Sub CmdScan_Click()
Monsters.ScanForMonsters
End Sub

Private Sub CmdSpellChange_Click()
lblquantiy.Visible = False
EquipList.Clear
Call ItemInfo.SpellsListing
FrameEquip.Caption = "Spells"
FrameEquip.Visible = True
End Sub

Private Sub cmdstamina_Click()
Dim Data As String
If MainPlayer.AttributePoints >= 1 Then
MainPlayer.MaxStamina = MainPlayer.MaxStamina + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.MaxStamina & " Stamina" & vbNewLine
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Stamina.Caption
Form1.SendClientData "stan", Data
Call LoadStats
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If

End Sub

Private Sub CmdStealth_Click()
If CmdStealth.FontBold = True Then
CmdStealth.FontBold = False
MainPlayer.Stealth = False
Form1.txtDialog.Text = Form1.txtDialog.Text & "You are no longer sneaking" & vbNewLine
Exit Sub
End If
If CmdStealth.FontBold = False Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You are sneaking" & vbNewLine
MainPlayer.Stealth = True
CmdStealth.FontBold = True

End If
End Sub

Private Sub cmdstrength_Click()
Dim Data As String
If MainPlayer.AttributePoints >= 1 Then
MainPlayer.Strength = MainPlayer.Strength + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.Strength & " Strength" & vbNewLine
Call LoadStats
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Strength.Caption
Form1.SendClientData "str", Data
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If

End Sub

Private Sub CmdSubmit_Click()
If OnlineGame = True Then


If Bugtxt.Text = "" Then
MsgBox "Please enter the bug you would like to submit!", vbExclamation, "Report a Bug!"
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "Thankyou for sending in this bug information we will check it out soon!" & vbNewLine
Call Form1.ReportBug(Bugtxt.Text)
End If
Else
MsgBox "Works in online game mode only!", vbExclamation

End If
End Sub

Private Sub CmdToggle_Click()

If ToggleTileType = True Then
ToggleTileType = False
Form1.RedrawMap
Exit Sub
End If
If ToggleTileType = False Then
ToggleTileType = True
Form1.RedrawMap
Exit Sub
End If


End Sub

Private Sub CmdTrainSkill_Click()
Form2.Show

End Sub

Private Sub CmdWisdom_Click()
Dim Data As String
If OnlineGame = False Then
MsgBox "Only works in online play", vbInformation
Exit Sub
End If
If MainPlayer.AttributePoints >= 1 Then
MainPlayer.Wisdom = MainPlayer.Wisdom + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints - 1
MainPlayer.MaxMana = MainPlayer.MaxMana + 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You now have " & MainPlayer.Wisdom & " Wisdom" & vbNewLine
Call LoadStats
Data = MainPlayer.PlayerName & "," & MainPlayer.AccountName & "," & Wisdom.Caption
Form1.SendClientData "wis", Data
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough attribute points!" & vbNewLine


End If

End Sub

Private Sub Combo1_Change()
Call LoadTileCombo
End Sub
Private Sub LoadTileCombo()
If Combo1.Text = "TrapTILE" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "LoadWarp" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "SupplyStore" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "HorseRanch" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "Destroyable" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "WeaponStore" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "MagicStore" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "ArmorStore" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "HurtTile" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "Swimable" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "Walkable" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "NONWALKABLE" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "DoorTile" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "DieTile" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "Sign" Then
walk.Caption = Combo1.Text
End If
If Combo1.Text = "Tree" Then
walk.Caption = Combo1.Text
End If
End Sub
Private Sub Combo1_Click()
Call LoadTileCombo
End Sub

Private Sub Combo1_Scroll()
Call LoadTileCombo
End Sub

Private Sub Command1_Click()
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
Frame4.Visible = True
Frame13.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Frame1.Visible = False
Frame3.Visible = False
Frame6.Visible = False
End Sub




Private Sub Command12_Click()
If OnlineGame = False Then
MsgBox "Enter these cheats in the text box! /xpme /morepoints /showmethemoney /warp /debugmode /debugmodeoff /allstuff   There is also a mapeditor look at the bottom of the screen for a button called map editor if you turn on cheats!"
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "Cheaters Never Prosper!" & vbNewLine
End If
End Sub

Private Sub Command13_Click()
EquipList.Clear
Call ItemInfo.ItemsListing
FrameEquip.Caption = "Items"
FrameEquip.Visible = True

End Sub

Private Sub Command2_Click()
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
Frame1.Visible = True
Frame13.Visible = False
Frame4.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame6.Visible = False
End Sub

Private Sub Command20_Click()
If OnlineGame = False Then

AddMonster.Show
Else

Form1.txtDialog.Text = Form1.txtDialog.Text & "Cheaters Never Prosper!" & vbNewLine

End If
End Sub




Private Sub Command25_Click()

If OnlineGame = False Then
PlayerSkills.CastSpell = PlayerSkills.CastSpell + 10

Label28.Caption = PlayerSkills.CastSpell
Else

Form1.txtDialog.Text = Form1.txtDialog.Text & "Cheaters Never Prosper!" & vbNewLine
End If

End Sub

Private Sub Command26_Click()
Form2.Caption = "Train Technology"
Form2.Frame1.Visible = True
Form2.Show

End Sub


Private Sub Command28_Click()
'Not neeeded if you get vertical scrolling to work
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
Combo1.Top = 240
Frame7.Left = 0
Frame7.Height = 9995
Frame7.Width = 19900
picDisplay.Width = 9990
picDisplay.Height = 10995
Frame9.Visible = False
Frame8.Visible = False
Frame12.Visible = False

TOOLS.Width = 19000

End Sub

Private Sub Command29_Click()
'Not neeeded if you get vertical scrolling to work
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
Frame7.Height = 2290
Frame7.Width = 2450
Frame7.Left = 600
Combo1.Top = 3840
picDisplay.Height = 1425
picDisplay.Width = 1900
Frame9.Visible = True
Frame8.Visible = True
Frame12.Visible = True

End Sub

Private Sub Command3_Click()
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
Frame2.Visible = True
Frame13.Visible = False
Frame4.Visible = False
Frame1.Visible = False
Frame3.Visible = False
Frame6.Visible = False

End Sub


Private Sub Command4_Click()
Form1.Shape1.Visible = False
Form1.HScroll1.Visible = False
Form1.VScroll1.Visible = False
Frame3.Visible = True
Frame13.Visible = False
Frame4.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Frame6.Visible = False
End Sub



Private Sub Command9_Click()
If OnlineGame = False Then
MainPlayer.Admin = True
TOOLS.CmdMapEd.Visible = True
Else

Form1.txtDialog.Text = Form1.txtDialog.Text & "Cheaters Never Prosper!" & vbNewLine
End If

End Sub

Private Sub EquipList_Click()
I = -1
For I = I + 1 To MaxItems
If EquipList.Text = PXItem(I).ItemName Then
lblquantiy.Caption = "Quanity: " & PXItem(I).Quanity
I = MaxItems
End If
Next I
End Sub

Private Sub EquipList_KeyPress(KeyAscii As Integer)
I = -1
For I = I + 1 To MaxItems
If EquipList.Text = PXItem(I).ItemName Then
lblquantiy.Caption = "Quanity: " & PXItem(I).Quanity
I = MaxItems
End If
Next I
End Sub

Private Sub EquipList_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I = -1
For I = I + 1 To MaxItems
If EquipList.Text = PXItem(I).ItemName Then
lblquantiy.Caption = "Quanity: " & PXItem(I).Quanity
I = MaxItems
End If
Next I
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Sends the keystrokes for movement back to the main form
Dim j As String
If KeyCode = vbKeyU Then
ItemInfo.UseSpell

End If
If KeyCode = vbKeyI Then
On Error GoTo no3

AppActivate Form1.Caption
j = "i"

SendKeys j
 
no3:
Exit Sub
End If
If KeyCode = vbKeyJ Then
On Error GoTo no2

AppActivate Form1.Caption
j = "j"

SendKeys j
 
no2:
Exit Sub
End If
If KeyCode = vbKeyK Then
On Error GoTo no1

AppActivate Form1.Caption
j = "k"

SendKeys j
 
no1:
Exit Sub
End If
If KeyCode = vbKeyL Then
On Error GoTo no4

AppActivate Form1.Caption
j = "l"

SendKeys j
 
no4:
Exit Sub
End If
End Sub

Private Sub Form_Load()

If MapEd = True Then
Else
Call PXLoadItems

Call ItemInfo.LoadPlayerItems
Call LoadList(Skills, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt")
Call LoadCharecterInfo

End If
Storage.Picture = LoadPicture(App.Path & "\Tiles.bmp")

'Blt the first tile to the picselected box
BitBlt picSelected.hDC, 0, 0, 32, 32, Storage.hDC, 0, 0, SRCCOPY: picSelected.Refresh

'Let the scrollbars max value = to how many tiles there are
HScroll2.Value = 0
HScroll2.Max = Storage.ScaleWidth \ 32 - 4
'VScroll2.Value = 0
'VScroll2.Max = Storage.ScaleHeight \ 32 - 4
'In case tiles.bmp width is too small change the max to 0
If HScroll2.Max < 0 Then HScroll2.Max = 0
RedrawTile
  
  F = FreeFile
'  On Error GoTo nofile2
Text1.Text = ""
Open App.Path & "\help.txt" For Input As F
Text1.Text = Input(LOF(F), F)
Close F
Text2.Text = ""
Open App.Path & "\quests.txt" For Input As F
Text2.Text = Input(LOF(F), F)
Close F

End Sub

Private Sub Form_Unload(Cancel As Integer)
RPG.SaveCharecter
End Sub


Private Sub GetInfo_Timer()


If Frame6.Visible = True Then Form1.Shape1.Visible = True

If RefreshMini = True Then RefreshMiniMap

'Refresh yellow box
Shape3.Left = HScroll
Shape3.Top = VScroll
'Make the two forms windowstates the same
If ZeldaMode = False Then
TOOLS.WindowState = Form1.WindowState
End If
'Dont move left pos of tools if start is minimized
If Form1.WindowState <> 1 And ZeldaMode = False Then
    TOOLS.Left = Form1.Left + Form1.Width
    TOOLS.Top = Form1.Top
End If

'Show the scroll bars values
Label22.Caption = (Snap(CurX, 32) \ 32) + HScroll
Label23.Caption = (Snap(CurY, 32) \ 32) + VScroll
End Sub

Private Sub GuardTimer_Timer()
Call Monsters.DamageforATTack
If OnlineGame = True Then
Server.AttackPlayer
End If
End Sub

Private Sub Helment_Click()
lblquantiy.Visible = True
EquipList.Clear
FrameEquip.Caption = "Helments"
FrameEquip.Visible = True

End Sub

Private Sub HScroll2_Change()
RedrawTile
End Sub

Private Sub HScroll2_Scroll()
RedrawTile
End Sub

Private Sub Item_DblClick()
Frame5.Visible = True
Picture1.Picture = Item.Picture
InfoName.Caption = Item.ToolTipText
lblDamage.Caption = ""
lblspeed.Caption = ""
Damage.Caption = ""
Speed.Caption = ""

End Sub

Private Sub Item_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 3 Then
mnuItemMenu.Tag = "Item"
PopupMenu mnuItemMenu, vbPopupMenuRightButton, Item.Left, Item.Top

End If
If Button = 1 Then

EquipList.Clear
Call ItemInfo.ItemsListing
FrameEquip.Caption = "Items"
FrameEquip.Visible = True
Else
mnuItemMenu.Tag = "item"
PopupMenu mnuItemMenu, vbPopupMenuCenterAlign, Item.Left + Item.Width, Item.Top + Item.Height


End If
End Sub

Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'This will check to make sure box doesnt go off screen it also moves box
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
If Button = 1 Then
    tHScroll = X: tVScroll = y
    If tHScroll < 0 Then tHScroll = 0
    If tVScroll < 0 Then tVScroll = 0
    If tHScroll + Shape3.Width > MiniMap.ScaleWidth Then tHScroll = MiniMap.ScaleWidth - Shape3.Width
    If tVScroll + Shape3.Height > MiniMap.ScaleHeight Then tVScroll = MiniMap.ScaleHeight - Shape3.Height
    HScroll = tHScroll: VScroll = tVScroll
Form1.HScroll1.Value = HScroll
Form1.VScroll1.Value = VScroll

    
End If
End Sub

Private Sub MiniMap_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'This will check to make sure box doesnt go off screen it also moves box
If OnlineGame = True And MainPlayer.Admin = False Then
Exit Sub 'This stops cheaters
End If
If Button = 1 Then
    tHScroll = X: tVScroll = y
    If tHScroll < 0 Then tHScroll = 0
    If tVScroll < 0 Then tVScroll = 0
    If tHScroll + Shape3.Width > MiniMap.ScaleWidth Then tHScroll = MiniMap.ScaleWidth - Shape3.Width
    If tVScroll + Shape3.Height > MiniMap.ScaleHeight Then tVScroll = MiniMap.ScaleHeight - Shape3.Height
    HScroll = tHScroll: VScroll = tVScroll
   Form1.HScroll1.Value = HScroll
Form1.VScroll1.Value = VScroll
End If
End Sub

Private Sub mnudrop_Click()

If mnuItemMenu.Tag = "Weapon" Then
'DropItem2 MainPlayer.xpos + 1, MainPlayer.ypos + 1, Weapon.ToolTipText, 1
'Form1.SendItem Weapon.ToolTipText, MainPlayer.MapName, MainPlayer.xpos + 1, MainPlayer.ypos + 1
If TOOLS.Pweapon.ToolTipText = "" Or TOOLS.Pweapon.ToolTipText = "Unarmed" Then Exit Sub
'DropItem2 MainPlayer.xpos + 1, MainPlayer.ypos + 1, Pweapon.ToolTipText, 1
ItemInfo.DropItem MainPlayer.xpos + 1, MainPlayer.ypos + 1, Pweapon.ToolTipText

'Form1.SendItem Pweapon.ToolTipText, MainPlayer.MapName, MainPlayer.xpos + 1, MainPlayer.ypos + 1
TOOLS.Pweapon.ToolTipText = "Unarmed"
TOOLS.Pweapon.Picture = LoadPicture(App.Path & "\monsters\Unarmed.gif")

End If
If mnuItemMenu.Tag = "Spell" Then

End If
If mnuItemMenu.Tag = "Armor" Then
ItemInfo.DropItem MainPlayer.xpos + 1, MainPlayer.ypos + 1, Parmor.ToolTipText

'Form1.SendItem Parmor.ToolTipText, MainPlayer.MapName, MainPlayer.xpos + 1, MainPlayer.ypos + 1
TOOLS.Item.ToolTipText = ""
TOOLS.Item.Picture = Armor.Picture

End If
If mnuItemMenu.Tag = "item" Then

If TOOLS.Item.ToolTipText = "" Then Exit Sub
'ItemInfo.DropItem2 MainPlayer.xpos + 1, MainPlayer.ypos + 1, Item.ToolTipText, 1
ItemInfo.DropItem MainPlayer.xpos + 1, MainPlayer.ypos + 1, Item.ToolTipText

'Form1.SendItem Item.ToolTipText, MainPlayer.MapName, MainPlayer.xpos + 1, MainPlayer.ypos + 1
TOOLS.Item.ToolTipText = ""
TOOLS.Item.Picture = Armor.Picture

End If
End Sub

Private Sub mnuview_Click()
If mnuItemMenu.Tag = "Weapon" Then
    Frame5.Visible = True
    Picture1.Picture = Pweapon.Picture
    InfoName.Caption = Pweapon.ToolTipText
    lblDamage.Caption = "Damage"
    lblspeed.Caption = "Speed"
    Call ItemInfo.WeaponInfo
End If
If mnuItemMenu.Tag = "Spell" Then
    Frame5.Visible = True
    Picture1.Picture = Spell.Picture
    InfoName.Caption = Spell.ToolTipText
    Call ItemInfo.SpellInfo
End If
If mnuItemMenu.Tag = "Armor" Then
    Frame5.Visible = True
    Picture1.Picture = Parmor.Picture
    InfoName.Caption = Parmor.ToolTipText
    lblDamage.Caption = "Armor Level"
    lblspeed.Caption = ""
    Speed.Caption = ""
    ItemInfo.ArmorInfo
End If
If mnuItemMenu.Tag = "Item" Then
    Frame5.Visible = True
    Picture1.Picture = Item.Picture
    InfoName.Caption = Item.ToolTipText
    lblDamage.Caption = ""
    lblspeed.Caption = ""
    Damage.Caption = ""
    Speed.Caption = ""
End If
End Sub

Private Sub Music_Click()
If Music.Value = 1 Then
Form1.MediaPlayer1.Play
Form1.MediaPlayer1.AutoStart = True
End If
If Music.Value = 0 Then
Form1.MediaPlayer1.Stop
Form1.MediaPlayer1.AutoStart = False

End If
Call SaveOptions
End Sub
Private Sub SaveOptions()
  F = FreeFile
Open (App.Path & "\options.txt") For Binary Access Write Lock Write As F
    Put #F, , "ShowPlayerNames" & vbNewLine
    Put #F, , ShowNames.Value & vbNewLine
    Put #F, , "AttackPlayers" & vbNewLine
    Put #F, , AttackPlayers.Value & vbNewLine
    Put #F, , "SoundEffects" & vbNewLine
    Put #F, , SoundEffects.Value & vbNewLine
    Put #F, , "Music" & vbNewLine
    Put #F, , Music.Value & vbNewLine
    Put #F, , "SaveChatLog" & vbNewLine
    Put #F, , SaveChat.Value & vbNewLine
Close F
End Sub
Private Sub NextLevel_Timer()
If MainPlayer.Xp4nextLevel <= 0 Then

MainPlayer.Xp4nextLevel = 0
MainPlayer.Xp4nextLevel = MainPlayer.Level * 2000
MainPlayer.SkillPoints = MainPlayer.SkillPoints + 1
MainPlayer.AttributePoints = MainPlayer.AttributePoints + 5
MainPlayer.Level = MainPlayer.Level + 1
Call LoadStats

Form1.txtDialog.Text = Form1.txtDialog.Text & "You are now level " & MainPlayer.Level & vbNewLine


End If
End Sub

Private Sub Pants_Click()
lblquantiy.Visible = True
EquipList.Clear
FrameEquip.Caption = "Pants"
FrameEquip.Visible = True
End Sub


Private Sub Parmor_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 3 Then
PopupMenu mnuItemMenu, vbPopupMenuRightButton, Parmor.Left, Parmor.Top
End If
If Button = 1 Then

EquipList.Clear
lblquantiy.Visible = True
Call ItemInfo.ArmorList
FrameEquip.Caption = "Armor"
FrameEquip.Visible = True
Else
mnuItemMenu.Tag = "Armor"
PopupMenu mnuItemMenu, vbPopupMenuCenterAlign, Parmor.Left + Parmor.Width, Parmor.Top + Parmor.Height


End If
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'If user presses left M. Button. then blit the pic to the Select Pic Box
If Button = 1 Then
    BitBlt picSelected.hDC, 0, 0, 32, 32, picDisplay.hDC, Snap(X, 32), Snap(y, 32), SRCCOPY: picSelected.Refresh
    TempTileX = Snap(X, 32) + HScroll2.Value * 32: TempTileY = Snap(y, 32)

End If
End Sub



Private Sub Pweapon_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 3 Then
PopupMenu mnuItemMenu, vbPopupMenuRightButton, Pweapon.Left, Pweapon.Top
End If
If Button = 1 Then
lblquantiy.Visible = True
EquipList.Clear
EquipList.AddItem "Unarmed"
Call ItemInfo.EquipmentList
FrameEquip.Caption = "Weapons"
FrameEquip.Visible = True
Else
mnuItemMenu.Tag = "Weapon"
PopupMenu mnuItemMenu, vbPopupMenuCenterAlign, Pweapon.Left + Pweapon.Width, Pweapon.Top + Pweapon.Height


End If
End Sub

Private Sub RestoreTimer_Timer()
'This timer heals the monster and the player over time. Health,Mana, and Stamina
Dim I As Integer
If MainPlayer.Health < MainPlayer.MaxHealth Then
MainPlayer.Health = MainPlayer.Health + 1
End If
If MainPlayer.Stamina < MainPlayer.MaxStamina Then
MainPlayer.Stamina = MainPlayer.Stamina + 1
End If
If MainPlayer.Mana < MainPlayer.MaxMana Then

MainPlayer.Mana = MainPlayer.Mana + 1

End If
Call LoadStats
I = 0
For I = I + 1 To 250 ' Iterate through each element.
On Error GoTo nofile:
    newmon = "MonsterImage" & I
If Monster(0).Health <= 0 Then
      
If Monster(0).Health <= 0 Then
Else
If Monster(0).Health < Monster(0).MaxHealth Then
Monster(0).Health = Monster(0).Health + 1

End If

End If
If Monster(0).Stamina <= 0 Then
Else
If Monster(0).Stamina < Monster(0).MaxStamina Then
Monster(0).Stamina = Monster(0).Stamina + 1

End If

End If
 

     
    End If
  Set P51 = Form1.Controls(newmon)
With P51

If Monster(I).Health <= 0 Then
Else
If Monster(I).Health < Monster(I).MaxHealth Then
Monster(I).Health = Monster(I).Health + 1

End If

End If

If Monster(I).Stamina <= 0 Then
Monster(I).Stamina = Monster(I).Stamina = 0
Else
If Monster(I).Stamina < Monster(I).MaxStamina Then
Monster(I).Stamina = Monster(I).Stamina + 1

End If

End If
If Monster(I).Mana <= 0 Then
Else
If Monster(I).Mana < Monster(I).MaxMana Then
Monster(I).Mana = Monster(I).Mana + 1

End If

End If

           .ToolTipText = Monster(I).PlayerName & " Level " & Monster(I).Level & " Health " & Monster(I).Health & "\" & Monster(I).MaxHealth
Call LoadStats
End With
    If Monster(I).MaxHealth = 0 Then

        Exit For
     End If
    
Next
Call LoadStats
nofile:
Exit Sub
End Sub

Private Sub SaveChat_Click()
Call SaveOptions
End Sub

Private Sub Sheild_Click()
lblquantiy.Visible = True
EquipList.Clear
FrameEquip.Caption = "Sheilds"
FrameEquip.Visible = True
End Sub

Private Sub Shoes_Click()
lblquantiy.Visible = True
EquipList.Clear
FrameEquip.Caption = "Shoes"
FrameEquip.Visible = True
End Sub

Private Sub ShowNames_Click()
If ShowNames.Value = 1 Then
Form1.lblPlayerName(0).Visible = True
End If
If ShowNames.Value = 0 Then
Form1.lblPlayerName(0).Visible = False
End If
Call SaveOptions
End Sub
Private Sub SoundEffects_Click()
If SoundEffects.Value = 1 Then
Form1.MediaPlayer2.Play
Form1.MediaPlayer2.AutoStart = True
End If
If SoundEffects.Value = 0 Then
Form1.MediaPlayer2.Stop
Form1.MediaPlayer2.AutoStart = False

End If
Call SaveOptions
End Sub

Private Sub Spell_Click()
mnuItemMenu.Tag = "Spell"
PopupMenu mnuItemMenu, , Spell.Left, Spell.Top

End Sub

Private Sub Spell_DblClick()
Frame5.Visible = True
Picture1.Picture = Spell.Picture
InfoName.Caption = Spell.ToolTipText
Call ItemInfo.SpellInfo

End Sub

Private Sub VScroll2_Change()
'RedrawTile2
End Sub

Private Sub VScroll2_Scroll()
'RedrawTile2
End Sub

Private Sub Weapon_Click()
mnuItemMenu.Tag = "Weapon"
PopupMenu mnuItemMenu, , Weapon.Left, Weapon.Top

End Sub

Private Sub Weapon_DblClick()
Frame5.Visible = True
Picture1.Picture = Weapon.Picture
InfoName.Caption = Weapon.ToolTipText
lblDamage.Caption = "Damage"
lblspeed.Caption = "Speed"

End Sub
Private Sub RefreshMiniMap()
'Refresh Mini Map
MiniMap.Cls
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        If Map(X, y).Tiletype >= 0 Then
            For tY = y To y + 1.2
                For tX = X To X + 1.5
                   BitBlt TOOLS.MiniMap.hDC, tX, tY, 1.5, 1.2, Storage.hDC, Map(X, y).TileX + 2, Map(X, y).Tiley + 2, SRCCOPY
                Next
            Next
        End If
    Next
Next
   MiniMap.Refresh
   RefreshMini = False
End Sub
Private Sub RedrawTile()
'The orginal redraw tile function
'An algorithnm that will blit tiles to the picDisplay box according to the scroll bars values

picDisplay.Cls
tempx = 0
For y = 0 To 64 Step 32
    For X = HScroll2.Value * 32 To HScroll2.Value * 32 + 96 Step 32
        BitBlt picDisplay.hDC, tempx, y, 32, 32, Storage.hDC, X, y, SRCCOPY
        'Because X will try to blit offscreen
        tempx = tempx + 32
    Next
tempx = 0
Next

picDisplay.Refresh
End Sub
Private Sub RedrawTile33()
'i am redoing this to added vertical tile support
'An algorithnm that will blit tiles to the picDisplay box according to the scroll bars values

picDisplay.Cls
tempx = 0
tempy = 0
For y = VScroll2.Value * 32 To VScroll2.Value * 32 + 96 Step 32
    For X = HScroll2.Value * 32 To HScroll2.Value * 32 + 96 Step 32
    Form1.Caption = tempx & " : : " & tempy
        BitBlt picDisplay.hDC, tempx, tempy, 32, 32, Storage.hDC, X, y, SRCCOPY
        'Because X will try to blit offscreen
        tempx = tempx + 32
    Next
    'Form1.Caption = tempx & " : : " & tempy
    tempy = tempy + 32
tempx = 0

Next
tempy = 0
picDisplay.Refresh
End Sub
Private Sub RedrawTile2()
'An algorithnm that will blit tiles to the picDisplay box according to the scroll bars values


picDisplay.Cls
tempx = 0
For y = 0 To 96 Step 32
    For X = VScroll2.Value * 32 To VScroll2.Value * 32 + 96 Step 32
        BitBlt picDisplay.hDC, y, tempx, 32, 32, Storage.hDC, y, X, SRCCOPY
        'Because X will try to blit offscreen
        tempx = tempx + 32
    Next
tempx = 0
Next

picDisplay.Refresh


End Sub
Public Sub LoadCharecterInfo()
If MapEd = True Then
Else
If OnlineGame = True Then Exit Sub
On Error GoTo nofile
Dim inputfile As Integer
inputfile = 1
Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".char" For Input As #inputfile
Input #inputfile, MainPlayer.Level, MainPlayer.Admin, MainPlayer.XP, MainPlayer.Xp4nextLevel, MainPlayer.MapName, MainPlayer.xpos, MainPlayer.ypos
Input #inputfile, MainPlayer.VScroll, MainPlayer.HScroll, MainPlayer.Graphic, MainPlayer.Money, MainPlayer.Health, MainPlayer.Stamina, MainPlayer.Mana
Input #inputfile, MainPlayer.MaxHealth, MainPlayer.MaxStamina, MainPlayer.MaxMana, MainPlayer.Strength, MainPlayer.Quickness, MainPlayer.Dexterity, MainPlayer.Wisdom
Input #inputfile, MainPlayer.Intelligence, MainPlayer.Burden, MainPlayer.KILLS, MainPlayer.Deaths, MainPlayer.SkillPoints, MainPlayer.SkillsPercent, MainPlayer.SpendXp, MainPlayer.BankAccountMoney
Input #inputfile, MainPlayer.AttributePoints, PlayerSkills.Axe, PlayerSkills.Blacksmith, PlayerSkills.Bombs, PlayerSkills.Bow, PlayerSkills.Builder, PlayerSkills.CastSpell, PlayerSkills.Dagger, PlayerSkills.Fail
Input #inputfile, PlayerSkills.Farmer, PlayerSkills.Hunter, PlayerSkills.Lumberjack, PlayerSkills.Mace, PlayerSkills.MagicDefense, PlayerSkills.MeleeDefense, PlayerSkills.Miner, PlayerSkills.Run, PlayerSkills.Scan, PlayerSkills.Spear
Input #inputfile, PlayerSkills.Staff, PlayerSkills.Stealth, PlayerSkills.Swim, PlayerSkills.Sword, PlayerSkills.TheForce, PlayerSkills.UnarmedAttack
Close #inputfile

'Here is a more secure way to store data just
'uncommented and comment the above input staments
''Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".char" For Binary Access Read Lock Read As #1
''Get #1, , MainPlayer
''Close #1
''Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".char" For Binary Access Read Lock Read As #2
''Get #2, , PlayerSkills
''Close #2

Call LoadStats
Call SkillStatLevel("All")

Call LoadStats
'Warping = True
'Loads the last map the player was on and his cordinates
Open App.Path & "\maps\" & MainPlayer.MapName For Input As #1: Close
Open App.Path & "\maps\" & MainPlayer.MapName For Binary Access Read Lock Read As #1
'
'Form1.EraseAll
Get #1, , Map
Close
'Warping = False
Form1.MapName.Caption = MainPlayer.MapName
Call RPG.FastXYWarp(MainPlayer.xpos, MainPlayer.ypos)
Call Stuff.CenterPerson
Debug.Print "###Player Stats Loaded###"
End If
Exit Sub
nofile:
Debug.Print "!!!Player Stats NOT Loaded!!!"
Exit Sub

End Sub
Sub LoadList(Lst As ListBox, file As String)

On Error GoTo error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
Lst.AddItem A$
Loop
Close 1
Exit Sub
error:
X = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub
Public Sub ReLoadSkills()
Skills.Clear
If MapEd = True Then
Else
Call LoadList(Skills, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt")
End If

End Sub
Public Sub LoadStats()
TOOLS.Level.Caption = MainPlayer.Level
TOOLS.Health.Caption = MainPlayer.Health & "/" & MainPlayer.MaxHealth
TOOLS.Stamina.Caption = MainPlayer.Stamina & "/" & MainPlayer.MaxStamina
TOOLS.Mana.Caption = MainPlayer.Mana & "/" & MainPlayer.MaxMana
TOOLS.Money.Caption = MainPlayer.Money
TOOLS.XP.Caption = MainPlayer.XP
TOOLS.SpendXp.Caption = MainPlayer.SpendXp
TOOLS.XP4NextLvl.Caption = MainPlayer.Xp4nextLevel
TOOLS.Dexiterity.Caption = MainPlayer.Dexterity
TOOLS.Strength.Caption = MainPlayer.Strength
TOOLS.Quickness.Caption = MainPlayer.Quickness
TOOLS.Wisdom.Caption = MainPlayer.Wisdom
TOOLS.Intelligence.Caption = MainPlayer.Intelligence
TOOLS.SpendXp.Caption = MainPlayer.SpendXp
TOOLS.SkillsPercent = MainPlayer.SkillsPercent
Call UpdateSkills
End Sub
Public Sub MenuAttack1()
CmdAttack_Click


End Sub
Public Sub GuardMode()
Call Monsters.ScanForMonsters
Call Monsters.attackMonster
End Sub
Private Sub UpdateSkills()
Call CheckSkills
Skills2.Clear
If TrueSkills.Axe = True Then
Skills2.AddItem ("Axe " & PlayerSkills.Axe)
End If
If TrueSkills.Bombs = True Then
Skills2.AddItem ("Bombs " & PlayerSkills.Bombs)
End If
If TrueSkills.Bow = True Then
Skills2.AddItem ("Bow " & PlayerSkills.Bow)
End If
If TrueSkills.CastSpell = True Then
Skills2.AddItem ("Cast Spell " & PlayerSkills.CastSpell)
End If
If TrueSkills.Dagger = True Then
Skills2.AddItem ("Dagger " & PlayerSkills.Dagger)
End If
If TrueSkills.Fail = True Then
Skills2.AddItem ("Fail " & PlayerSkills.Fail)
End If
If TrueSkills.Mace = True Then
Skills2.AddItem ("Mace " & PlayerSkills.Mace)
End If
If TrueSkills.MagicDefense = True Then
Skills2.AddItem ("Magic Defense " & PlayerSkills.MagicDefense)
End If
If TrueSkills.MeleeDefense = True Then
Skills2.AddItem ("Melee Defense " & PlayerSkills.MeleeDefense)
End If
If TrueSkills.Spear = True Then
Skills2.AddItem ("Spear " & PlayerSkills.Spear)
End If
If TrueSkills.Staff = True Then
Skills2.AddItem ("Staff " & PlayerSkills.Staff)
End If
If TrueSkills.Stealth = True Then
Skills2.AddItem ("Stealth " & PlayerSkills.Stealth)
End If
If TrueSkills.Sword = True Then
Skills2.AddItem ("Sword " & PlayerSkills.Sword)
End If
If TrueSkills.Force = True Then
Skills2.AddItem ("The Force " & PlayerSkills.TheForce)
End If
If TrueSkills.unarmed = True Then
Skills2.AddItem ("Unarmed Attack " & PlayerSkills.UnarmedAttack)
End If
End Sub
Private Sub CheckSkills()

For I = I + 1 To Skills.ListCount

If TOOLS.Skills.List(I) = "Axe 6" Then
TrueSkills.Axe = True
End If
If TOOLS.Skills.List(I) = "BlackSmith 20" Then
TrueSkills.Blacksmith = True
End If
If TOOLS.Skills.List(I) = "Bombs 16" Then
TrueSkills.Bombs = True
End If
If TOOLS.Skills.List(I) = "Bow 20" Then
TrueSkills.Bow = True
End If
If TOOLS.Skills.List(I) = "Cast Spell 20" Then
TrueSkills.CastSpell = True
End If
If TOOLS.Skills.List(I) = "Dagger 8" Then
TrueSkills.Dagger = True
End If
If TOOLS.Skills.List(I) = "Fail 10" Then
TrueSkills.Fail = True
End If
If TOOLS.Skills.List(I) = "The Force 30" Then
TrueSkills.Force = True
End If
If TOOLS.Skills.List(I) = "Mace 6" Then
TrueSkills.Mace = True
End If
If TOOLS.Skills.List(I) = "Magic Defense 24" Then
TrueSkills.MagicDefense = True
End If
If TOOLS.Skills.List(I) = "Melee Defense 20" Then
TrueSkills.MeleeDefense = True
End If
If TOOLS.Skills.List(I) = "Miner 18" Then
TrueSkills.Miner = True
End If
If TOOLS.Skills.List(I) = "Scan 4" Then
TrueSkills.Scan = True
End If
If TOOLS.Skills.List(I) = "Spear 10" Then
TrueSkills.Spear = True
End If
If TOOLS.Skills.List(I) = "Staff 12" Then
TrueSkills.Staff = True
End If
If TOOLS.Skills.List(I) = "Stealth 20" Then
TrueSkills.Stealth = True
End If
If TOOLS.Skills.List(I) = "Sword 18" Then
TrueSkills.Sword = True
End If
If TOOLS.Skills.List(I) = "Unarmed Attack 20" Then
TrueSkills.unarmed = True
End If
Next
End Sub

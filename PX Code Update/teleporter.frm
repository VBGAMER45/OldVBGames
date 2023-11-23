VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form teleporter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warp Map! Click on a place to warp"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "teleporter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdWarp 
      Caption         =   "Warp"
      Height          =   495
      Left            =   5880
      TabIndex        =   155
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtmapname 
      Height          =   285
      Left            =   1800
      TabIndex        =   153
      Text            =   "mapname.map"
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   960
      OleObjectBlob   =   "teleporter.frx":030A
      TabIndex        =   131
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "DungeonWarp"
      Height          =   5175
      Left            =   360
      TabIndex        =   125
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox Picture101 
         Height          =   735
         Left            =   2640
         ScaleHeight     =   675
         ScaleWidth      =   1035
         TabIndex        =   126
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Image Zup 
         Height          =   330
         Left            =   2760
         Picture         =   "teleporter.frx":0370
         Top             =   3480
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image ZDown 
         Height          =   330
         Left            =   2160
         Picture         =   "teleporter.frx":0756
         Top             =   3480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ZLeft 
         Height          =   345
         Left            =   3360
         Picture         =   "teleporter.frx":084B
         Top             =   3360
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image ZRight 
         Height          =   345
         Left            =   4080
         Picture         =   "teleporter.frx":0C45
         Top             =   3600
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image Image27 
         Height          =   330
         Left            =   1920
         Picture         =   "teleporter.frx":1040
         Top             =   4440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image Image28 
         Height          =   330
         Left            =   1200
         Picture         =   "teleporter.frx":143A
         Top             =   4320
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image Image29 
         Height          =   480
         Left            =   600
         Picture         =   "teleporter.frx":1836
         Top             =   4320
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image Image30 
         Height          =   480
         Left            =   2520
         Picture         =   "teleporter.frx":1C4B
         Top             =   4320
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label24 
         Caption         =   "Click on the box!"
         Height          =   375
         Left            =   2520
         TabIndex        =   129
         Top             =   1560
         Width           =   1815
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
      Height          =   255
      Left            =   6360
      OleObjectBlob   =   "teleporter.frx":204F
      TabIndex        =   152
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "teleporter.frx":20B1
      TabIndex        =   151
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
      Height          =   255
      Left            =   5160
      OleObjectBlob   =   "teleporter.frx":2111
      TabIndex        =   150
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "teleporter.frx":2171
      TabIndex        =   149
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   3960
      OleObjectBlob   =   "teleporter.frx":21D1
      TabIndex        =   148
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   255
      Left            =   3240
      OleObjectBlob   =   "teleporter.frx":2231
      TabIndex        =   147
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   255
      Left            =   2640
      OleObjectBlob   =   "teleporter.frx":2291
      TabIndex        =   146
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   2040
      OleObjectBlob   =   "teleporter.frx":22F1
      TabIndex        =   145
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   1560
      OleObjectBlob   =   "teleporter.frx":2351
      TabIndex        =   144
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "teleporter.frx":23B1
      TabIndex        =   143
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":2411
      TabIndex        =   142
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":2473
      TabIndex        =   141
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":24D3
      TabIndex        =   140
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":2533
      TabIndex        =   139
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":2593
      TabIndex        =   138
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":25F3
      TabIndex        =   137
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":2653
      TabIndex        =   136
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":26B3
      TabIndex        =   135
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":2713
      TabIndex        =   134
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "teleporter.frx":2773
      TabIndex        =   133
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "teleporter.frx":27D3
      TabIndex        =   132
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   615
      Left            =   5640
      OleObjectBlob   =   "teleporter.frx":2839
      TabIndex        =   130
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   4560
      OleObjectBlob   =   "teleporter.frx":28BF
      Top             =   5400
   End
   Begin VB.Timer TimerNocheat 
      Interval        =   2
      Left            =   6480
      Top             =   5760
   End
   Begin VB.PictureBox Picture102 
      Height          =   615
      Left            =   4560
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   127
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox ypos 
      Height          =   285
      Left            =   960
      MaxLength       =   2
      TabIndex        =   122
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox xpos 
      Height          =   285
      Left            =   0
      MaxLength       =   2
      TabIndex        =   121
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture100 
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   118
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture99 
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   117
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture98 
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   116
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture97 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   115
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture96 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   114
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture95 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   113
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture94 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   112
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture93 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   111
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture92 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   110
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture91 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   109
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture90 
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   108
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture89 
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   107
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture88 
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   106
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture87 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   105
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture86 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   104
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture85 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   103
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture84 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   102
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture83 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   101
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture82 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   100
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox Picture81 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   99
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture80 
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   96
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture79 
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   95
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture78 
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   94
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture77 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   93
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture76 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   92
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture75 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   91
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture74 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   90
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture73 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   89
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture72 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   88
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture71 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   87
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture70 
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   86
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture69 
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   85
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture68 
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   84
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture67 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   83
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture66 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   82
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture65 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   81
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture64 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   80
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture63 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   79
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture62 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   78
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture61 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   77
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture60 
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   74
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture59 
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   73
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture58 
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   72
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture57 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   71
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture56 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   70
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture55 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   69
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture54 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   68
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture53 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   67
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture52 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   66
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture51 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   65
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture50 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   64
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture49 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   63
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture48 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   62
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture47 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   61
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture46 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   60
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture45 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   59
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture44 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   58
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture43 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   57
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture42 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   56
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture41 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   55
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture40 
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   52
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture39 
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   51
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture38 
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   50
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture37 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   49
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture36 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   48
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture35 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   47
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture34 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   46
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture33 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   45
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture32 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   44
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture31 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   43
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture30 
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   42
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture29 
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   41
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture28 
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   40
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture27 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   39
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture26 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   38
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture25 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   37
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture24 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   36
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture23 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   35
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture22 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture21 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture20 
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture19 
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture18 
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture17 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture10 
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture9 
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox Picture8 
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3120
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblmapname 
      Caption         =   "Mapname"
      Height          =   255
      Left            =   2040
      TabIndex        =   154
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Horse Ranch and Farm"
      Height          =   495
      Left            =   5640
      TabIndex        =   128
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Xpos"
      Height          =   255
      Left            =   480
      TabIndex        =   124
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "Ypos"
      Height          =   255
      Left            =   960
      TabIndex        =   123
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   "10"
      Height          =   255
      Left            =   6360
      TabIndex        =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label19 
      Caption         =   "9"
      Height          =   255
      Left            =   5640
      TabIndex        =   119
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label18 
      Caption         =   "8"
      Height          =   255
      Left            =   5160
      TabIndex        =   98
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label17 
      Caption         =   "7"
      Height          =   255
      Left            =   4440
      TabIndex        =   97
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "6"
      Height          =   255
      Left            =   3960
      TabIndex        =   76
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "5"
      Height          =   255
      Left            =   3240
      TabIndex        =   75
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "4"
      Height          =   255
      Left            =   2760
      TabIndex        =   54
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "3"
      Height          =   255
      Left            =   2040
      TabIndex        =   53
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "2"
      Height          =   255
      Left            =   1560
      TabIndex        =   32
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "10"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "9"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "8"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "7"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "6"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "5"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "4"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "3"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "2"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "teleporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form holds all of zeldas movements

Private Sub CmdCancel_Click()
If Form1.Direction = "up" Then
Form1.Image1.Top = Form1.Image1.Top + 32
Form1.lblTileKind = "1"
Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
If Form1.Direction = "down" Then
Form1.Image1.Top = Form1.Image1.Top + 32
Form1.lblTileKind = "1"
Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
If Form1.Direction = "left" Then
Form1.Image1.Top = Form1.Image1.Top + 32
Form1.lblTileKind = "1"
Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
If Form1.Direction = "right" Then
Form1.Image1.Top = Form1.Image1.Top + 32
Form1.lblTileKind = "1"
Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
'adjust player's name
    Form1.lblPlayerName(0).Top = Form1.Image1.Top - 20
    Form1.lblPlayerName(0).Left = Form1.Image1.Left

Unload Me

End Sub

Private Sub CmdWarp_Click()
On Error GoTo nofile
Form1.MediaPlayer2.Filename = App.Path & "/warp.wav"
Open App.Path & "\maps\" & txtmapname.Text For Input As #1: Close
Open App.Path & "\maps\" & txtmapname.Text For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = txtmapname.Text
Form1.MapName = txtmapname.Text
Form1.RedrawMap
Call RPG.FastXYWarp(xpos.Text, ypos.Text)
Unload Me
nofile:
Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim response As String
If KeyCode = vbKeyP Then
If MainPlayer.Admin = True Then Exit Sub
response = InputBox("Enter a code", "Code!")
If response = "ranch" Then
Picture102.Visible = True
Label23.Visible = True

End If

End If
End Sub

Private Sub Form_Load()
RPG.ForceTextBoxNumeric xpos, True
RPG.ForceTextBoxNumeric ypos, True
If MainPlayer.MapName = "dungeon1.map" Then
Frame1.Visible = True
End If
If MainPlayer.Level >= 50 Then
Picture102.Visible = True
Label23.Visible = True
End If

If Bosses(0).Boss1Defeated = True Then
Picture10.BackColor = vbGreen
Picture10.ToolTipText = "Boss Defeated!"
End If
If MainPlayer.Admin = True Then
xpos.Visible = True
ypos.Visible = True
txtmapname.Visible = True
CmdWarp.Visible = True
lblmapname.Visible = True
End If
If SkinFilename = "" Then Exit Sub
  Skn.LoadSkin SkinFilename ' Loads another skin into Skin component
  Skn.ApplySkin Me.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Stuff.CenterPerson
 Call ImplodeForm(Me, 2, 500, 1)
 Form1.Refresh
 TOOLS.Refresh
 
End Sub

Private Sub Picture1_Click()
Form1.MediaPlayer2.Filename = App.Path & "/warp.wav"
  Open App.Path & "\maps\1.1.map" For Input As #1: Close
Open App.Path & "\maps\1.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "1.1.map"
Form1.MapName = "1.1.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\1.1.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\1.1.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:
Form1.EraseSpecialMap
Unload Me
Exit Sub

Unload Me
End Sub

Private Sub Picture10_Click()
  Open App.Path & "\maps\10.1.map" For Input As #1: Close
Open App.Path & "\maps\10.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "10.1.map"
Form1.MapName = "10.1.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\10.1.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\10.1.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:

Form1.EraseSpecialMap
Unload Me
Exit Sub
End Sub

Private Sub Picture100_Click()
  Open App.Path & "\maps\10.10.map" For Input As #1: Close
Open App.Path & "\maps\10.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "10.10.map"
Form1.MapName = "10.10.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\10.10.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\10.10.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:

Form1.EraseSpecialMap
Unload Me
Exit Sub
End Sub

Private Sub Picture101_Click()
If MainPlayer.MapName = "dungeon1.map" Then
Open App.Path & "\maps\b1.map" For Input As #1: Close
Open App.Path & "\maps\b1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\tiles2.bmp")

MainPlayer.MapName = "b1.map"
Form1.MapName = "b1.map"
Form1.RedrawMap
Form1.HScroll1.Value = 0
Form1.VScroll1.Value = 0
Form1.Image1.Top = 208
Form1.Image1.Left = 232
Unload Me

End If
End Sub

Private Sub Picture102_Click()
  Open App.Path & "\maps\ranch.map" For Input As #1: Close
Open App.Path & "\maps\ranch.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "ranch.map"
Form1.MapName = "ranch.map"
Form1.RedrawMap
Form1.HScroll1.Value = 45
Form1.VScroll1.Value = 29
Form1.Image1.Top = 208
Form1.Image1.Left = 232
Unload Me
End Sub

Private Sub Picture11_Click()
Form1.MediaPlayer2.Filename = App.Path & "/warp.wav"
  Open App.Path & "\maps\1.2.map" For Input As #1: Close
Open App.Path & "\maps\1.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "1.2.map"
Form1.MapName = "1.2.map"
Form1.RedrawMap


On Error GoTo Nofile9
Open App.Path & "\maps\1.2.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\1.2.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:

Form1.EraseSpecialMap
Unload Me
Exit Sub


End Sub

Private Sub Picture12_Click()
  Open App.Path & "\maps\2.2.map" For Input As #1: Close
Open App.Path & "\maps\2.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "2.2.map"
Form1.MapName = "2.2.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\2.2.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\2.2.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:

Form1.EraseSpecialMap
Unload Me
Exit Sub
End Sub

Private Sub Picture13_Click()
  Open App.Path & "\maps\3.2.map" For Input As #1: Close
Open App.Path & "\maps\3.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "3.2.map"
Form1.MapName = "3.2.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\3.2.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\3.2.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:

Form1.EraseSpecialMap
Unload Me
Exit Sub

End Sub

Private Sub Picture14_Click()
  Open App.Path & "\maps\4.2.map" For Input As #1: Close
Open App.Path & "\maps\4.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "4.2.map"
Form1.MapName = "4.2.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\4.2.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\4.2.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:
Form1.EraseSpecialMap
Unload Me
Exit Sub
End Sub

Private Sub Picture15_Click()
Open App.Path & "\maps\5.2.map" For Input As #1: Close
Open App.Path & "\maps\5.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "5.2.map"
Form1.MapName = "5.2.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\5.2.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\5.2.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:
Form1.EraseSpecialMap
Unload Me
Exit Sub
End Sub

Private Sub Picture16_Click()
  Open App.Path & "\maps\6.2.map" For Input As #1: Close
Open App.Path & "\maps\6.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "6.2.map"
Form1.MapName = "6.2.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\6.2.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\6.2.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:
Form1.EraseSpecialMap
Unload Me
Exit Sub
End Sub

Private Sub Picture17_Click()
  Open App.Path & "\maps\7.2.map" For Input As #1: Close
Open App.Path & "\maps\7.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "7.2.map"
Form1.MapName = "7.2.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture18_Click()
  Open App.Path & "\maps\8.2.map" For Input As #1: Close
Open App.Path & "\maps\8.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "8.2.map"
Form1.MapName = "8.2.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture19_Click()
  Open App.Path & "\maps\9.2.map" For Input As #1: Close
Open App.Path & "\maps\9.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "9.2.map"
Form1.MapName = "9.2.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture2_Click()
  Open App.Path & "\maps\2.1.map" For Input As #1: Close
Open App.Path & "\maps\2.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "2.1.map"
Form1.MapName = "2.1.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture20_Click()
  Open App.Path & "\maps\10.2.map" For Input As #1: Close
Open App.Path & "\maps\10.2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "10.2.map"
Form1.MapName = "10.2.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture21_Click()
Form1.MediaPlayer2.Filename = App.Path & "/warp.wav"
  Open App.Path & "\maps\1.3.map" For Input As #1: Close
Open App.Path & "\maps\1.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "1.3.map"
Form1.MapName = "1.3.map"
Form1.RedrawMap


Unload Me
End Sub

Private Sub Picture22_Click()
  Open App.Path & "\maps\2.3.map" For Input As #1: Close
Open App.Path & "\maps\2.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "2.3.map"
Form1.MapName = "2.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture23_Click()
  Open App.Path & "\maps\3.3.map" For Input As #1: Close
Open App.Path & "\maps\3.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "3.3.map"
Form1.MapName = "3.3.map"
Form1.RedrawMap
Unload Me
End Sub


Private Sub Picture24_Click()
  Open App.Path & "\maps\4.3.map" For Input As #1: Close
Open App.Path & "\maps\4.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "4.3.map"
Form1.MapName = "4.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture25_Click()
  Open App.Path & "\maps\5.3.map" For Input As #1: Close
Open App.Path & "\maps\5.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "5.3.map"
Form1.MapName = "5.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture26_Click()
  Open App.Path & "\maps\6.3.map" For Input As #1: Close
Open App.Path & "\maps\6.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "6.3.map"
Form1.MapName = "6.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture27_Click()
  Open App.Path & "\maps\7.3.map" For Input As #1: Close
Open App.Path & "\maps\7.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "7.3.map"
Form1.MapName = "7.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture28_Click()
  Open App.Path & "\maps\8.3.map" For Input As #1: Close
Open App.Path & "\maps\8.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "8.3.map"
Form1.MapName = "8.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture29_Click()
  Open App.Path & "\maps\9.3.map" For Input As #1: Close
Open App.Path & "\maps\9.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "9.3.map"
Form1.MapName = "9.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture3_Click()
  Open App.Path & "\maps\3.1.map" For Input As #1: Close
Open App.Path & "\maps\3.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "3.1.map"
Form1.MapName = "3.1.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture30_Click()
  Open App.Path & "\maps\10.3.map" For Input As #1: Close
Open App.Path & "\maps\10.3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "10.3.map"
Form1.MapName = "10.3.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture31_Click()
  Open App.Path & "\maps\1.4.map" For Input As #1: Close
Open App.Path & "\maps\1.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "1.4.map"
Form1.MapName = "1.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture32_Click()
  Open App.Path & "\maps\2.4.map" For Input As #1: Close
Open App.Path & "\maps\2.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "2.4.map"
Form1.MapName = "2.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture33_Click()
  Open App.Path & "\maps\3.4.map" For Input As #1: Close
Open App.Path & "\maps\3.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "3.4.map"
Form1.MapName = "3.4.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture34_Click()
  Open App.Path & "\maps\4.4.map" For Input As #1: Close
Open App.Path & "\maps\4.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "4.4.map"
Form1.MapName = "4.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture35_Click()
  Open App.Path & "\maps\5.4.map" For Input As #1: Close
Open App.Path & "\maps\5.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "5.4.map"
Form1.MapName = "5.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture36_Click()
  Open App.Path & "\maps\6.4.map" For Input As #1: Close
Open App.Path & "\maps\6.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "6.4.map"
Form1.MapName = "6.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture37_Click()
  Open App.Path & "\maps\7.4.map" For Input As #1: Close
Open App.Path & "\maps\7.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "7.4.map"
Form1.MapName = "7.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture38_Click()
  Open App.Path & "\maps\8.4.map" For Input As #1: Close
Open App.Path & "\maps\8.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "8.4.map"
Form1.MapName = "8.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture39_Click()
  Open App.Path & "\maps\9.4.map" For Input As #1: Close
Open App.Path & "\maps\9.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "9.4.map"
Form1.MapName = "9.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture4_Click()
  Open App.Path & "\maps\4.1.map" For Input As #1: Close
Open App.Path & "\maps\4.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "4.1.map"
Form1.MapName = "4.1.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture40_Click()
  Open App.Path & "\maps\10.4.map" For Input As #1: Close
Open App.Path & "\maps\10.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "10.4.map"
Form1.MapName = "10.4.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture41_Click()
  Open App.Path & "\maps\1.5.map" For Input As #1: Close
Open App.Path & "\maps\1.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "1.5.map"
Form1.MapName = "1.5.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture42_Click()
  Open App.Path & "\maps\2.5.map" For Input As #1: Close
Open App.Path & "\maps\2.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "2.5.map"
Form1.MapName = "2.5.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture43_Click()
  Open App.Path & "\maps\3.5.map" For Input As #1: Close
Open App.Path & "\maps\3.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "3.5.map"
Form1.MapName = "3.5.map"
Form1.RedrawMap
Unload Me
End Sub


Private Sub Picture44_Click()
   Open App.Path & "\maps\4.5.map" For Input As #1: Close
Open App.Path & "\maps\4.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "4.5.map"
Form1.MapName = "4.5.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture45_Click()
  Open App.Path & "\maps\newbie" For Input As #1: Close
Open App.Path & "\maps\newbie" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "newbie"
Form1.MapName = "newbie"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture46_Click()
  Open App.Path & "\maps\6.5.map" For Input As #1: Close
Open App.Path & "\maps\6.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "6.5.map"
Form1.MapName = "6.5.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture47_Click()
 Open App.Path & "\maps\7.5.map" For Input As #1: Close
Open App.Path & "\maps\7.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "7.5.map"
Form1.MapName = "7.5.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture48_Click()
 Open App.Path & "\maps\8.5.map" For Input As #1: Close
Open App.Path & "\maps\8.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "8.5.map"
Form1.MapName = "8.5.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture49_Click()
 Open App.Path & "\maps\9.5.map" For Input As #1: Close
Open App.Path & "\maps\9.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "9.5.map"
Form1.MapName = "9.5.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture5_Click()
  Open App.Path & "\maps\5.1.map" For Input As #1: Close
Open App.Path & "\maps\5.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "5.1.map"
Form1.MapName = "5.1.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture50_Click()
 Open App.Path & "\maps\10.5.map" For Input As #1: Close
Open App.Path & "\maps\10.5.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "10.5.map"
Form1.MapName = "10.5.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture51_Click()
  Open App.Path & "\maps\1.6.map" For Input As #1: Close
Open App.Path & "\maps\1.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "1.6.map"
Form1.MapName = "1.6.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture52_Click()
  Open App.Path & "\maps\2.6.map" For Input As #1: Close
Open App.Path & "\maps\2.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "2.6.map"
Form1.MapName = "2.6.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture53_Click()
  Open App.Path & "\maps\3.6.map" For Input As #1: Close
Open App.Path & "\maps\3.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "3.6.map"
Form1.MapName = "3.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture54_Click()
  Open App.Path & "\maps\4.6.map" For Input As #1: Close
Open App.Path & "\maps\4.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "4.6.map"
Form1.MapName = "4.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture55_Click()
  Open App.Path & "\maps\5.6.map" For Input As #1: Close
Open App.Path & "\maps\5.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "5.6.map"
Form1.MapName = "5.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture56_Click()
  Open App.Path & "\maps\6.6.map" For Input As #1: Close
Open App.Path & "\maps\6.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "6.6.map"
Form1.MapName = "6.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture57_Click()
  Open App.Path & "\maps\7.6.map" For Input As #1: Close
Open App.Path & "\maps\7.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "7.6.map"
Form1.MapName = "7.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture58_Click()
  Open App.Path & "\maps\8.6.map" For Input As #1: Close
Open App.Path & "\maps\8.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "8.6.map"
Form1.MapName = "8.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture59_Click()
  Open App.Path & "\maps\9.6.map" For Input As #1: Close
Open App.Path & "\maps\9.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "9.6.map"
Form1.MapName = "9.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture6_Click()
  Open App.Path & "\maps\6.1.map" For Input As #1: Close
Open App.Path & "\maps\6.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "6.1.map"
Form1.MapName = "6.1.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture60_Click()
  Open App.Path & "\maps\10.6.map" For Input As #1: Close
Open App.Path & "\maps\10.6.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "10.6.map"
Form1.MapName = "10.6.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture61_Click()
  Open App.Path & "\maps\1.7.map" For Input As #1: Close
Open App.Path & "\maps\1.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "1.7.map"
Form1.MapName = "1.7.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture62_Click()
  Open App.Path & "\maps\2.7.map" For Input As #1: Close
Open App.Path & "\maps\2.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "2.7.map"
Form1.MapName = "2.7.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture63_Click()
  Open App.Path & "\maps\3.7.map" For Input As #1: Close
Open App.Path & "\maps\3.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "3.7.map"
Form1.MapName = "3.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture64_Click()
  Open App.Path & "\maps\4.7.map" For Input As #1: Close
Open App.Path & "\maps\4.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "4.7.map"
Form1.MapName = "4.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture65_Click()
  Open App.Path & "\maps\5.7.map" For Input As #1: Close
Open App.Path & "\maps\5.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "5.7.map"
Form1.MapName = "5.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture66_Click()
  Open App.Path & "\maps\6.7.map" For Input As #1: Close
Open App.Path & "\maps\6.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "6.7.map"
Form1.MapName = "6.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture67_Click()
  Open App.Path & "\maps\7.7.map" For Input As #1: Close
Open App.Path & "\maps\7.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "7.7.map"
Form1.MapName = "7.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture68_Click()
  Open App.Path & "\maps\8.7.map" For Input As #1: Close
Open App.Path & "\maps\8.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "8.7.map"
Form1.MapName = "8.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture69_Click()
  Open App.Path & "\maps\9.7.map" For Input As #1: Close
Open App.Path & "\maps\9.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "9.7.map"
Form1.MapName = "9.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture7_Click()
  Open App.Path & "\maps\7.1.map" For Input As #1: Close
Open App.Path & "\maps\7.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "7.1.map"
Form1.MapName = "7.1.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture70_Click()
  Open App.Path & "\maps\10.7.map" For Input As #1: Close
Open App.Path & "\maps\10.7.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "10.7.map"
Form1.MapName = "10.7.map"
Form1.RedrawMap
Unload Me

End Sub

Private Sub Picture71_Click()
  Open App.Path & "\maps\1.8.map" For Input As #1: Close
Open App.Path & "\maps\1.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "1.8.map"
Form1.MapName = "1.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture72_Click()
  Open App.Path & "\maps\2.8.map" For Input As #1: Close
Open App.Path & "\maps\2.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "2.8.map"
Form1.MapName = "2.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture73_Click()
  Open App.Path & "\maps\3.8.map" For Input As #1: Close
Open App.Path & "\maps\3.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "3.8.map"
Form1.MapName = "3.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture74_Click()
  Open App.Path & "\maps\4.8.map" For Input As #1: Close
Open App.Path & "\maps\4.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "4.8.map"
Form1.MapName = "4.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture75_Click()
  Open App.Path & "\maps\5.8.map" For Input As #1: Close
Open App.Path & "\maps\5.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "5.8.map"
Form1.MapName = "5.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture76_Click()
  Open App.Path & "\maps\6.8.map" For Input As #1: Close
Open App.Path & "\maps\6.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "6.8.map"
Form1.MapName = "6.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture77_Click()
  Open App.Path & "\maps\7.8.map" For Input As #1: Close
Open App.Path & "\maps\7.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "7.8.map"
Form1.MapName = "7.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture78_Click()
  Open App.Path & "\maps\8.8.map" For Input As #1: Close
Open App.Path & "\maps\8.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "8.8.map"
Form1.MapName = "8.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture79_Click()
  Open App.Path & "\maps\9.8.map" For Input As #1: Close
Open App.Path & "\maps\9.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "9.8.map"
Form1.MapName = "9.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture8_Click()
  Open App.Path & "\maps\8.1.map" For Input As #1: Close
Open App.Path & "\maps\8.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "8.1.map"
Form1.MapName = "8.1.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture80_Click()
  Open App.Path & "\maps\10.8.map" For Input As #1: Close
Open App.Path & "\maps\10.8.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "10.8.map"
Form1.MapName = "10.8.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture81_Click()
  Open App.Path & "\maps\1.9.map" For Input As #1: Close
Open App.Path & "\maps\1.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "1.9.map"
Form1.MapName = "1.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture82_Click()
  Open App.Path & "\maps\2.9.map" For Input As #1: Close
Open App.Path & "\maps\2.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
'Form1.Money.ToolTipText = ""
MainPlayer.MapName = "2.9.map"
Form1.MapName = "2.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture83_Click()
  Open App.Path & "\maps\3.9.map" For Input As #1: Close
Open App.Path & "\maps\3.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "3.9.map"
Form1.MapName = "3.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture84_Click()
  Open App.Path & "\maps\4.9.map" For Input As #1: Close
Open App.Path & "\maps\4.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "4.9.map"
Form1.MapName = "4.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture85_Click()
  Open App.Path & "\maps\5.9.map" For Input As #1: Close
Open App.Path & "\maps\5.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "5.9.map"
Form1.MapName = "5.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture86_Click()
  Open App.Path & "\maps\6.9.map" For Input As #1: Close
Open App.Path & "\maps\6.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "6.9.map"
Form1.MapName = "6.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture87_Click()
  Open App.Path & "\maps\7.9.map" For Input As #1: Close
Open App.Path & "\maps\7.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "7.9.map"
Form1.MapName = "7.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture88_Click()
  Open App.Path & "\maps\8.9.map" For Input As #1: Close
Open App.Path & "\maps\8.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "3.9.map"
Form1.MapName = "8.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture89_Click()
  Open App.Path & "\maps\9.9.map" For Input As #1: Close
Open App.Path & "\maps\9.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "3.9.map"
Form1.MapName = "9.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture9_Click()
  Open App.Path & "\maps\9.1.map" For Input As #1: Close
Open App.Path & "\maps\9.1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
MainPlayer.MapName = "9.1.map"
Form1.MapName = "9.1.map"
Form1.RedrawMap
On Error GoTo Nofile9
Open App.Path & "\maps\9.1.map" & ".spcm" For Input As #1: Close
Open App.Path & "\maps\9.1.map" & ".spcm" For Binary Access Read Lock Read As #1
Form1.EraseSpecialMap
Get #1, , SpecialMap
Unload Me
Close
Nofile9:
Form1.EraseSpecialMap
Unload Me
Exit Sub
End Sub

Private Sub Picture90_Click()
  Open App.Path & "\maps\10.9.map" For Input As #1: Close
Open App.Path & "\maps\10.9.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "10.9.map"
Form1.MapName = "10.9.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture91_Click()
  Open App.Path & "\maps\1.10.map" For Input As #1: Close
Open App.Path & "\maps\1.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "1.10.map"
Form1.MapName = "1.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture92_Click()
  Open App.Path & "\maps\2.10.map" For Input As #1: Close
Open App.Path & "\maps\2.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "2.10.map"
Form1.MapName = "2.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture93_Click()
  Open App.Path & "\maps\3.10.map" For Input As #1: Close
Open App.Path & "\maps\3.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "3.10.map"
Form1.MapName = "3.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture94_Click()
  Open App.Path & "\maps\4.10.map" For Input As #1: Close
Open App.Path & "\maps\4.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "4.10.map"
Form1.MapName = "4.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture95_Click()
  Open App.Path & "\maps\5.10.map" For Input As #1: Close
Open App.Path & "\maps\5.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "5.10.map"
Form1.MapName = "5.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture96_Click()
  Open App.Path & "\maps\6.10.map" For Input As #1: Close
Open App.Path & "\maps\6.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "6.10.map"
Form1.MapName = "6.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture97_Click()
  Open App.Path & "\maps\7.10.map" For Input As #1: Close
Open App.Path & "\maps\7.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "7.10.map"
Form1.MapName = "7.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture98_Click()
  Open App.Path & "\maps\8.10.map" For Input As #1: Close
Open App.Path & "\maps\8.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "8.10.map"
Form1.MapName = "8.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub Picture99_Click()
  Open App.Path & "\maps\9.10.map" For Input As #1: Close
Open App.Path & "\maps\9.10.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "2.10.map"
Form1.MapName = "9.10.map"
Form1.RedrawMap
Unload Me
End Sub

Private Sub TimerNocheat_Timer()
If MainPlayer.Level < 50 Then
Picture102.Visible = False
Label23.Visible = False
End If
If SkinFilename = "" Then
SkinLabel1.Visible = False
SkinLabel2.Visible = False
SkinLabel3.Visible = False
SkinLabel4.Visible = False
SkinLabel5.Visible = False
SkinLabel6.Visible = False
SkinLabel7.Visible = False
SkinLabel8.Visible = False
SkinLabel9.Visible = False
SkinLabel10.Visible = False
SkinLabel11.Visible = False
SkinLabel13.Visible = False
SkinLabel12.Visible = False
SkinLabel14.Visible = False
SkinLabel15.Visible = False
SkinLabel16.Visible = False
SkinLabel17.Visible = False
SkinLabel18.Visible = False
SkinLabel19.Visible = False
SkinLabel20.Visible = False
SkinLabel21.Visible = False
SkinLabel22.Visible = False
SkinLabel23.Visible = False

Else
'SkinLabel1.Visible = True
'SkinLabel2.Visible = True
'SkinLabel3.Visible = True
SkinLabel4.Visible = True
SkinLabel5.Visible = True
SkinLabel6.Visible = True
SkinLabel7.Visible = True
SkinLabel8.Visible = True
SkinLabel9.Visible = True
SkinLabel10.Visible = True
SkinLabel11.Visible = True
SkinLabel13.Visible = True
SkinLabel12.Visible = True
SkinLabel14.Visible = True
SkinLabel15.Visible = True
SkinLabel16.Visible = True
SkinLabel17.Visible = True
SkinLabel18.Visible = True
SkinLabel19.Visible = True
SkinLabel20.Visible = True
SkinLabel21.Visible = True
SkinLabel22.Visible = True
SkinLabel23.Visible = True

End If

End Sub

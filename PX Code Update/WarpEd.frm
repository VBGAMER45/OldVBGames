VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PX Warp Editor"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtsign 
      Height          =   375
      Left            =   840
      TabIndex        =   23
      Text            =   "Sign Text"
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox txtmapname 
      Height          =   285
      Left            =   480
      TabIndex        =   20
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox xpos 
      Height          =   285
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   19
      Text            =   "0"
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox ypos 
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   18
      Text            =   "0"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Make a Sign"
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   6480
      Width           =   2415
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Warp To"
      Height          =   3255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   4095
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtYcord 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Text            =   "0"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtXCord 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtTiles 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "tiles.bmp"
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtmap2 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Left                       Top"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "X                               Y"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "TilesName"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "MapName"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "MapName"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "x                            Y"
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "x                            Y"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "MapName"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  F = FreeFile
Open (App.Path & "\scripts\" & txtMap.Text & txtX.Text & txtY.Text & ".warp") For Binary Access Write Lock Write As F
    Put #F, , txtmap2.Text & vbCrLf
    Put #F, , txtYcord.Text & vbCrLf
    Put #F, , txtXCord.Text & vbCrLf
  '  Put #F, , Text4.Text & vbCrLf
   ' Put #F, , Text5.Text & vbCrLf
   ' Put #F, , txtTiles.Text & vbCrLf
Close F
MsgBox "WARP Created", vbExclamation

End Sub

Private Sub Command2_Click()
  F = FreeFile
Open (App.Path & "\scripts\" & txtmapname.Text & xpos.Text & ypos.Text & ".sign") For Binary Access Write Lock Write As F
    Put #F, , txtsign.Text & vbCrLf
Close F
MsgBox "Sign Created", vbExclamation
End Sub


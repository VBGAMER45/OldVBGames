VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Old Script Window"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "Script.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRandomMap 
      Caption         =   "Random Map Creator"
      Height          =   3615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton CmdMap 
         Caption         =   "Create Map"
         Height          =   375
         Left            =   1680
         TabIndex        =   26
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Script.frx":030A
         Left            =   720
         List            =   "Script.frx":033E
         TabIndex        =   25
         Text            =   "forest"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtFlowers 
         Height          =   285
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   22
         Text            =   "0"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtTrees 
         Height          =   285
         Left            =   480
         MaxLength       =   5
         TabIndex        =   21
         Text            =   "0"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Flowers, Desert Trees"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Trees, Sand Dunes"
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.Frame FramePost 
      Caption         =   "Post"
      Height          =   3615
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label PostContent 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   2475
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   5145
      End
      Begin VB.Label postedby 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Posted By "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   5895
      End
   End
   Begin VB.TextBox Part11 
      Height          =   285
      Left            =   3120
      TabIndex        =   16
      Top             =   4920
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5400
      Picture         =   "Script.frx":03D1
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   3600
      Width           =   540
   End
   Begin VB.TextBox Part10 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Part9 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Part8 
      Height          =   285
      Left            =   480
      TabIndex        =   12
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Part3 
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Part7 
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Part6 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Part5 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Part4 
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Part2 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Part1 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   -600
      Width           =   735
   End
   Begin VB.FileListBox FileItem 
      Height          =   870
      Left            =   3000
      Pattern         =   "*.item*"
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.FileListBox FileMonster 
      Height          =   870
      Left            =   4560
      Pattern         =   "*.bad*"
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox FileSpell 
      Height          =   1065
      Left            =   5160
      Pattern         =   "*.spel*"
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.FileListBox FileNPc 
      Height          =   1065
      Left            =   4800
      Pattern         =   "*.npc*"
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Water As Boolean

Private Sub CmdMap_Click()
If Combo1.Text = "ocean" And Combo1.Text = "island" Then
Water = True
Else
Water = False
End If
Call ModRndMap.RandomMapGenerator(Combo1.Text, txtTrees.Text, Water, txtFlowers.Text, True)
Form1.RedrawMap

End Sub

'i used to have a scripting system but
'got rid of it this form is used
'for the message board.


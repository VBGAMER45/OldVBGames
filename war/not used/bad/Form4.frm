VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Dimploacy"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form4"
   ScaleHeight     =   4680
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   2400
      TabIndex        =   24
      Top             =   480
      Width           =   3015
      Begin VB.OptionButton Option7 
         Caption         =   "Option7"
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   2400
      TabIndex        =   23
      Top             =   1080
      Width           =   3015
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   195
         Left            =   2400
         TabIndex        =   27
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   2400
      TabIndex        =   22
      Top             =   2520
      Width           =   3015
      Begin VB.OptionButton Option11 
         Caption         =   "Option11"
         Height          =   195
         Left            =   2400
         TabIndex        =   33
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Option9"
         Height          =   195
         Left            =   1440
         TabIndex        =   32
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option8"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2400
      TabIndex        =   21
      Top             =   3240
      Width           =   3015
      Begin VB.OptionButton Option14 
         Caption         =   "Option14"
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Option13"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Option12"
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2400
      TabIndex        =   17
      Top             =   1800
      Width           =   3015
      Begin VB.OptionButton Option10 
         Caption         =   "Option10"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Option16"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2040
      TabIndex        =   12
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Offer allies or enemy money and click check box"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
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
      Left            =   600
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Ally"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Neteual"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Enemy"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
'Randomize
Label5.Caption = "greek"
'Randomize
End Sub


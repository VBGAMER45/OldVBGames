VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Communicate"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option5 
      Caption         =   "Option5"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Put Text in the box"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Player"
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "5"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Call TEXT1
Unload Me

End Sub

Private Sub Option5_Click()

If Option5.Value = True Then


Call CHECKED
End If
End Sub

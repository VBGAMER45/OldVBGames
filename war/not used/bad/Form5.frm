VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "News"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form5"
   ScaleHeight     =   3345
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2400
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label year 
      Caption         =   "0"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Timer1_Timer()
'If Timer1.Interval = 300 Then
'year = year + 10
Call TIMER
year = year + 10
'End If
End Sub


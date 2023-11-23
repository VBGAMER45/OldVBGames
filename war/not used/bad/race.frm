VERSION 5.00
Begin VB.Form race 
   Caption         =   "Race"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label6 
      Caption         =   "Enemy6"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Enemy3"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Enemy2"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Enemy5"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Enemy4"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Human"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   6720
      X2              =   0
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   1080
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   855
      Left            =   3240
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   5280
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   960
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   3120
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   5280
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "race"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Form9.Show vbModal, Me
Unload Me

End Sub

Private Sub Image2_Click()
Form9.Show vbModal, Me
Unload Me


End Sub

Private Sub Image3_Click()
 Form9.Show vbModal, Me
 Unload Me
 
End Sub

Private Sub Image4_Click()
Form9.Show vbModal, Me
Unload Me

End Sub

Private Sub Image5_Click()
Form9.Show vbModal, Me
Unload Me

End Sub

Private Sub Image6_Click()
Form9.Show vbModal, Me
Unload Me

End Sub


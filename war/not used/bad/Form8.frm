VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Settings"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form8"
   ScaleHeight     =   3690
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2775
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2775
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

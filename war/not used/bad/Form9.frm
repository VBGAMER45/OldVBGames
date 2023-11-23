VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Map"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1560
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Dim form2 As New form2
form2.Show vbModal, Me

End Sub

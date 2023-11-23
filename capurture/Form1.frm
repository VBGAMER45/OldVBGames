VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000040&
   Caption         =   "Capature the flag"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   855
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "4 people"
      Height          =   975
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3 people"
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2 People"
      Height          =   855
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   960
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   3480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response As String

Private Sub Command1_Click()
Unload Me

Form2.Show
Form2.Image4.Visible = False
Form2.Image4.Enabled = False
Form2.Image5.Visible = False
Form2.Image5.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me

Form2.Show
Form2.Image5.Visible = False
Form2.Image5.Enabled = False

End Sub

Private Sub Command3_Click()
Unload Me

Form2.Show
End Sub

Private Sub Command4_Click()
response = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit!!!")

If response = vbYes Then


End
End If
End Sub

Private Sub Image1_Click()
MsgBox "Zelda Game website at http://vbgames6.homestead.com/index.html", vbExclamation
End Sub

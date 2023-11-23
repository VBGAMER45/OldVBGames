VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1920
      Top             =   5880
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   2520
      Top             =   6120
   End
   Begin VB.CommandButton cmdGetMachineID 
      Caption         =   "Get IP"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   5280
   End
   Begin VB.Image Image7 
      Height          =   360
      Left            =   2160
      Picture         =   "Form1.frx":0000
      Top             =   2160
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   2160
      Picture         =   "Form1.frx":0822
      Top             =   2160
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   4200
      Picture         =   "Form1.frx":1044
      Top             =   7080
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   5520
      Picture         =   "Form1.frx":161A
      Top             =   6960
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   3240
      Picture         =   "Form1.frx":1D64
      Top             =   6960
      Width           =   525
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2160
      Picture         =   "Form1.frx":2ABA
      Top             =   6960
      Width           =   600
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Expect This design to get better in the furture!!!!!!!!!"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   5715
      Left            =   0
      Picture         =   "Form1.frx":381C
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Machine Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "    IP Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Programming Languages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   3360
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "VB Games!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   -120
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Jonathan Valentin 2000"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "VB Tips And Tricks!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   6720
      TabIndex        =   6
      Top             =   -360
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Animated Gifs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   5760
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   -120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Java Games"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      Top             =   -120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Java Script"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Baisc Games, Code and More! Active X Site!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "VB Code!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   -120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetMachineID_Click()

   Text1 = GetIPHostName()
   Text2 = GetIPAddress()
   
  
End Sub

Private Sub Label2_Click()
frmBrowser.Show
frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/code.html"
frmBrowser.brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub Timer1_Timer()
Label2.Top = Label2.Top + 20
Label4.Top = Label2.Top + 20
Label5.Top = Label2.Top + 20
Label7.Top = Label2.Top + 20

Label9.Top = Label2.Top + 20
If Label2.Top >= 1080 Then
Timer1.Enabled = False

End If

End Sub

Private Sub Timer2_Timer()


Image6.Visible = True

Image7.Visible = False
Timer3.Enabled = True
Timer2.Enabled = False


End Sub

Private Sub Timer3_Timer()
Image7.Visible = True

Image6.Visible = False
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

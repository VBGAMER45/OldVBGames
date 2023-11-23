VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00000000&
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ScaleHeight     =   7065
   ScaleWidth      =   9600
   Begin VB.Timer Timer2 
      Interval        =   550
      Left            =   2160
      Top             =   5760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   550
      Left            =   1560
      Top             =   5520
   End
   Begin VB.CommandButton cmdGetMachineID 
      Caption         =   "Get IP"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2640
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
      Left            =   3360
      TabIndex        =   11
      Top             =   2640
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
      Left            =   3360
      TabIndex        =   10
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   2280
      Top             =   4920
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Links!"
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
      Left            =   7440
      TabIndex        =   17
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic Chat!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   2400
      Picture         =   "UserControl1.ctx":0000
      Top             =   1800
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   360
      Left            =   2400
      Picture         =   "UserControl1.ctx":0822
      Top             =   1800
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Expect This design to get better in the furture!!!!!!!!!"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   4560
      Picture         =   "UserControl1.ctx":1044
      Top             =   6480
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   5880
      Picture         =   "UserControl1.ctx":161A
      Top             =   6360
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   3600
      Picture         =   "UserControl1.ctx":1D64
      Top             =   6360
      Width           =   525
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2520
      Picture         =   "UserControl1.ctx":2ABA
      Top             =   6360
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   5715
      Left            =   120
      Picture         =   "UserControl1.ctx":381C
      Top             =   1320
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
      Left            =   3360
      TabIndex        =   14
      Top             =   2280
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
      Left            =   3360
      TabIndex        =   13
      Top             =   3120
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
      Left            =   3480
      TabIndex        =   9
      Top             =   3960
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
      Left            =   5160
      TabIndex        =   8
      Top             =   -120
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Jonathan Valentin 2000"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   6000
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
      Left            =   6840
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
      Left            =   5880
      TabIndex        =   5
      Top             =   3960
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
      Left            =   840
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
      Left            =   1920
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
      Left            =   1680
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
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
      Left            =   3720
      TabIndex        =   1
      Top             =   -120
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub cmdGetMachineID_Click()

   Text1 = GetIPHostName()
   Text2 = GetIPAddress()

End Sub

Private Sub Label10_Click()

frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/otherlang.html"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label12_Click()
frmDPlayDevices.Show vbModal, Me


End Sub

Private Sub Label13_Click()


frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/links.html"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label2_Click()


frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/code.html"

 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label3_Click()


frmBrowser.cboAddress.Text = "http://www.homestead.com/visualbasiczone/java.html"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label4_Click()

frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/java.html"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label5_Click()


frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/files/cheats.htm"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label6_Click()

frmBrowser.cboAddress.Text = "http://www.homestead.com/visualbasiczone/gifs.html"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label7_Click()

frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/tips.html"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
End Sub

Private Sub Label8_Click()
MsgBox "Last Edited September 8,2000", vbExclamation



End Sub

Private Sub Label9_Click()

frmBrowser.cboAddress.Text = "http://www.homestead.com/vbgames6/game.html"
 frmBrowser.brwWebBrowser.Navigate frmBrowser.cboAddress.Text
frmBrowser.Show vbModal, Me
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

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form11 
   Caption         =   "Multiplayer"
   ClientHeight    =   5640
   ClientLeft      =   945
   ClientTop       =   1500
   ClientWidth     =   10770
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":030A
   ScaleHeight     =   5640
   ScaleWidth      =   10770
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   9960
      Top             =   4200
   End
   Begin VB.Frame Frame7 
      Height          =   375
      Left            =   8400
      TabIndex        =   61
      Top             =   4920
      Width           =   1695
      Begin VB.CheckBox Check1 
         Caption         =   "Host"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   9960
      Top             =   3720
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   5280
      TabIndex        =   56
      Text            =   "Map"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      Height          =   375
      Left            =   7320
      TabIndex        =   54
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Chat"
      Height          =   375
      Left            =   0
      TabIndex        =   53
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   51
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9840
      Top             =   3240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HangUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   48
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   47
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "call"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   46
      Top             =   4920
      Width           =   855
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   3240
      TabIndex        =   44
      Text            =   "RACE"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   3240
      TabIndex        =   43
      Text            =   "RACE"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   3240
      TabIndex        =   42
      Text            =   "RACE"
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3240
      TabIndex        =   41
      Text            =   "RACE"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3240
      TabIndex        =   40
      Text            =   "RACE"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form11.frx":53050
      Left            =   3240
      List            =   "Form11.frx":53052
      TabIndex        =   39
      Text            =   "RACE"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   7080
      TabIndex        =   33
      Top             =   4320
      Width           =   2655
      Begin VB.OptionButton Option24 
         BackColor       =   &H00000000&
         Caption         =   "Option24"
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         ToolTipText     =   "TEAM 4"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option23 
         BackColor       =   &H00000000&
         Caption         =   "Option23"
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         ToolTipText     =   "TEAM 3"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option22 
         BackColor       =   &H00000000&
         Caption         =   "Option22"
         Height          =   255
         Left            =   840
         TabIndex        =   35
         ToolTipText     =   "TEAM 2"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option21 
         BackColor       =   &H00000000&
         Caption         =   "Option21"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "TEAM 1"
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   7080
      TabIndex        =   24
      Top             =   3840
      Width           =   2655
      Begin VB.OptionButton Option20 
         BackColor       =   &H00000000&
         Caption         =   "Option20"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "TEAM 1"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option19 
         BackColor       =   &H00000000&
         Caption         =   "Option19"
         Height          =   255
         Left            =   840
         TabIndex        =   31
         ToolTipText     =   "TEAM 2"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option18 
         BackColor       =   &H00000000&
         Caption         =   "Option18"
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         ToolTipText     =   " TEAM 3"
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option17 
         BackColor       =   &H00000000&
         Caption         =   "Option17"
         Height          =   375
         Left            =   2280
         TabIndex        =   29
         ToolTipText     =   "TEAM 4"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   7080
      TabIndex        =   23
      Top             =   3360
      Width           =   2655
      Begin VB.OptionButton Option16 
         BackColor       =   &H00000000&
         Caption         =   "Option16"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         ToolTipText     =   "TEAM 4"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option15 
         BackColor       =   &H00000000&
         Caption         =   "Option15"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         ToolTipText     =   " TEAM 3"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H00000000&
         Caption         =   "Option14"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         ToolTipText     =   "TEAM 2"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00000000&
         Caption         =   "Option13"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "TEAM 1"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Top             =   3000
      Width           =   2655
      Begin VB.OptionButton Option12 
         BackColor       =   &H00000000&
         Caption         =   "Option12"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "TEAM 1"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00000000&
         Caption         =   "Option11"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         ToolTipText     =   "TEAM 2"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00000000&
         Caption         =   "Option10"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         ToolTipText     =   " TEAM 3"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00000000&
         Caption         =   "Option9"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         ToolTipText     =   "TEAM 4"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   7080
      TabIndex        =   13
      Top             =   2520
      Width           =   2655
      Begin VB.OptionButton Option8 
         BackColor       =   &H00000000&
         Caption         =   "Option8"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "TEAM 1"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00000000&
         Caption         =   "Option7"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         ToolTipText     =   "TEAM 2"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00000000&
         Caption         =   "Option6"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         ToolTipText     =   " TEAM 3"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00000000&
         Caption         =   "Option4"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "TEAM 4"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   2160
      Width           =   2655
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Option3"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "TEAM 1"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Option2"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         ToolTipText     =   "TEAM 2"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         ToolTipText     =   " TEAM 3"
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00000000&
         Caption         =   "Option5"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         ToolTipText     =   "TEAM 4"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LAUNCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Launch"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   6120
      Picture         =   "Form11.frx":53054
      ToolTipText     =   "Game is Locked"
      Top             =   4320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5520
      Picture         =   "Form11.frx":5335E
      ToolTipText     =   "Game is Unlocked!"
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   60
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   59
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   58
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   57
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   55
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   4920
      Picture         =   "Form11.frx":53668
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   52
      Top             =   5280
      Width           =   10815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   50
      ToolTipText     =   "Status"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1200
      TabIndex        =   49
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   45
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "RACE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   38
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "4"
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
      Left            =   9480
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "3"
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
      Left            =   8640
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "2"
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
      Left            =   7920
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "1"
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
      Left            =   7200
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Teams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "    WELCOME TO MULTIPLAYER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   9855
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   3600
      Picture         =   "Form11.frx":5800A
      Top             =   2760
      Width           =   960
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim team1 As String
Dim team2 As String
Dim team3 As String
Dim team4 As String
Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long


Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
    '
     Const RAS95_MaxEntryName = 256
     Const RAS95_MaxDeviceType = 16
     Const RAS95_MaxDeviceName = 32
    '


Private Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
    End Type
    '


Private Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
    End Type
Private Function IsConnected() As Boolean


    '
    Dim TRasCon(255) As RASCONN95
    Dim lg As Long
    Dim lpcon As Long
    Dim RetVal As Long
    Dim Tstatus As RASCONNSTATUS95
    '
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
    '
    RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)


    If RetVal <> 0 Then
        MsgBox "ERROR"
        Exit Function
    End If


    '
    Tstatus.dwSize = 160
    RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)


    If Tstatus.RasConnState = &H2000 Then
        IsConnected = True
    Else
        IsConnected = False
    End If


End Function

Private Sub Command5_Click()
Unload Me

End Sub

Private Sub Command6_Click()
'Label12.Caption = InputBox("Chat,enter stuff")
Form61.Show vbModal, Me

End Sub

Private Sub Command7_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
frmAbout.Show vbModal, Me
End Sub

Private Sub Image3_Click()
Image4.Visible = True
Image3.Visible = False
End Sub

Private Sub Image4_Click()
Image3.Visible = True
Image4.Visible = False
End Sub

Private Sub Timer1_Timer()
Dim IsConnectedOrNot
IsConnectedOrNot = IsConnected()
Select Case IsConnectedOrNot
    Case True  'If connected
        Label11.Caption = "Status: You are connected"
    Case False  'if not connected
        Label11.Caption = "Status: You are not connected"
End Select
End Sub


Private Sub Command1_Click()
'Call Module1.load
If Combo7.Text = "Random Map" Then



form2.Show vbModal, Me
End If
End Sub

Private Sub Command2_Click()
'InputBox ("Enter There Phone Number:", "phone")
InputBox "Enter phone number,Phone"
'MaskedEdit1.Mask = "()- "
'MaskedEdit1.Text = "(   )   -    "
''''strnumber$ = InputBox
''''RequestID% = PDQTapi1.PlaceCall(CommControl, strnumber$)


'Copyright ® 1996 Progress Software Corp.
End Sub

Private Sub Command3_Click()
MsgBox "waiting for call"
Success% = PDQTapi1.WaitForCall(CommControl)

'Copyright ® 1996 Progress Software Corp.
End Sub

Private Sub Command4_Click()
MsgBox "are you sure you want to hang up", vbYesNo
If vbYes Then

requestID% = PDQTapi1.HangUp
Else
MsgBox "hanged up"
End If

'Copyright ® 1996 Progress Software Corp.
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Soviet"
    Combo1.AddItem "Germany"
    Combo1.AddItem "UNITED STATES"
    Combo1.AddItem "ALEIN"
    Combo1.AddItem "FORGOTTON"
    Combo1.AddItem "CPU"
      Combo2.AddItem "Soviet"
    Combo2.AddItem "Germany"
    Combo2.AddItem "UNITED STATES"
    Combo2.AddItem "ALEIN"
    Combo2.AddItem "FORGOTTON"
    Combo2.AddItem "CPU"
      Combo3.AddItem "Soviet"
    Combo3.AddItem "Germany"
    Combo3.AddItem "UNITED STATES"
    Combo3.AddItem "ALEIN"
    Combo3.AddItem "FORGOTTON"
    Combo3.AddItem "CPU"
      Combo4.AddItem "Soviet"
    Combo4.AddItem "germany"
    Combo4.AddItem "UNITED STATES"
    Combo4.AddItem "ALEIN"
    Combo4.AddItem "FORGOTTON"
    Combo4.AddItem "CPU"
      Combo5.AddItem "Soviet"
    Combo5.AddItem "Germany"
    Combo5.AddItem "UNITED STATES"
    Combo5.AddItem "ALEIN"
    Combo5.AddItem "FORGOTTON"
    Combo5.AddItem "CPU"
      Combo6.AddItem "Soviet"
    Combo6.AddItem "Germany"
    Combo6.AddItem "UNITED STATES"
    Combo6.AddItem "ALEIN"
    Combo6.AddItem "FORGOTTON"
    Combo6.AddItem "CPU"
    Combo7.AddItem "Random"
    Combo7.AddItem "Load Map"
    
''''''   PDQComm1.PortOpen = True
  ' ProgressBar1.Scrolling
End Sub


Private Sub Option2_Click()
If Option2.Value = True Then
team2 = True
End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
team1 = True
End If


End Sub

'Jonathan Valentin
Private Sub Timer2_Timer()
If Combo7.Text = "Load Map" Then

CommonDialog1.ShowOpen
Combo7.Text = "Load Map1"
'form2.Show vbModal, Me

End If

End Sub

Private Sub Timer3_Timer()
Label12.Caption " Start the Game Already"

End Sub

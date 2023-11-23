VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4200
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   7380
   Icon            =   "FRMSPL~1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerDraw 
      Interval        =   951
      Left            =   3600
      Top             =   2280
   End
   Begin VB.Frame Light2 
      Caption         =   "Lightning Bolt Color"
      Height          =   975
      Left            =   240
      TabIndex        =   33
      Top             =   4800
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton OptionColor 
         Caption         =   "Grey Scale Dark"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   41
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "Grey Scale Light"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "P"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   39
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "T"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   38
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "Y"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "R"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "G"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "B"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Light1 
      Caption         =   "Lightning Bolt Stats"
      Height          =   2415
      Left            =   240
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CheckBox ChkSplit 
         Caption         =   "Split Bolt"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox ChkCloud 
         Caption         =   "Show Cloud"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   2295
      End
      Begin VB.HScrollBar NickNum 
         Height          =   255
         Left            =   120
         Max             =   30
         TabIndex        =   25
         Top             =   240
         Value           =   30
         Width           =   2775
      End
      Begin VB.HScrollBar Speed 
         Height          =   255
         Left            =   120
         Max             =   299
         TabIndex        =   24
         Top             =   1200
         Value           =   149
         Width           =   2775
      End
      Begin VB.HScrollBar HeightNum 
         Height          =   255
         Left            =   120
         Max             =   200
         TabIndex        =   23
         Top             =   480
         Value           =   200
         Width           =   2775
      End
      Begin VB.HScrollBar Velocity 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   22
         Top             =   720
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar Velocity 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   20
         Min             =   2
         TabIndex        =   21
         Top             =   960
         Value           =   20
         Width           =   2775
      End
      Begin VB.CheckBox ChkStart 
         Caption         =   "Always Start Same Location"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height: 190"
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblNicks 
         Caption         =   "Nicks: 20"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblspeed 
         Caption         =   "Speed: 150"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblVel 
         Caption         =   "V Min: 5"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblVel 
         Caption         =   "V Max: 15"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdMaped 
      Caption         =   "&MapEditor Offline"
      Height          =   615
      Left            =   6000
      TabIndex        =   14
      Top             =   480
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   10
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.CommandButton CmdSecond 
         Caption         =   "Second Connect"
         Height          =   615
         Left            =   3360
         TabIndex        =   42
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton CmdList 
         Caption         =   "&Credits"
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   2520
         Width           =   1215
      End
      Begin VB.PictureBox LogoPic 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   240
         Picture         =   "FRMSPL~1.frx":000C
         ScaleHeight     =   2415
         ScaleWidth      =   2415
         TabIndex        =   18
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton CmdSinglePlayer 
         Caption         =   "&Single Player Game"
         Height          =   975
         Left            =   5760
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   6240
         TabIndex        =   13
         Top             =   3600
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2880
         Top             =   2160
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   3840
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "C&onnect"
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox ServerIP 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Text            =   "MainServer"
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Check out the message board for the for the ip.  If the main server is down."
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Enter the ServerIP"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Normal Client"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5340
         TabIndex        =   7
         Top             =   2640
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Jonathan Valentin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3360
         TabIndex        =   5
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Image imgLogo2 
         Height          =   2385
         Left            =   480
         Picture         =   "FRMSPL~1.frx":1130E
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 2001-2003"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "VisualBasicZone.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   3
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   " Project X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   2985
      End
   End
   Begin VB.Label LbLStatus 
      Alignment       =   2  'Center
      Caption         =   "Checking for Update"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   5295
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You have permission to use the code in anyway shape or form
'in your own game but if you could add my name in the credit that would be nice


'My Plans for 2003
'Redo the entire store code
'I plan on fixing the store code
'Improving Internet Play

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'The following is for the tiny lightning on the earth

Private Type LightningBolt
    Center As Long
    Inner1 As Long
    Inner2 As Long
    Inner3 As Long
    Inner4 As Long
    Outer1 As Long
    Outer2 As Long
    Outer3 As Long
    Outer4 As Long
    Nicks As Long
    VelocityMin As Long
    VelocityMax As Long
    Height As Long
    SplitBolt As Boolean
    SameStart As Boolean
    ShowCloud As Boolean
End Type
Dim Bolt As LightningBolt

Private Function DrawGlobe(Pic As PictureBox)
Dim SM, I, LX, LY, SX, SY, LX2, LY2, SX2, SY2 As Integer
Dim Alter As Boolean
Pic.Cls

If Bolt.SameStart = True Then
    SM = Pic.ScaleWidth \ 2
Else
    SM = Int(Rnd * (Pic.ScaleWidth \ 2) + Pic.ScaleWidth \ 4)
End If
Pic.Forecolor = Bolt.Center
SX = SMSub
If Bolt.ShowCloud = True Then
    SY = 10
Else
    SY = 0
    SetPixel Pic.hDC, SM, 0, Bolt.Center
End If
I = Int(Rnd * 2)
If I = 0 Then
    Alter = False
    LX = SX + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
Else
    Alter = True
    LX = SX - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
End If
LY = Bolt.Height / Bolt.Nicks
For I = 0 To Bolt.Nicks
    Pic.Forecolor = Bolt.Center
    Pic.Line (SX, SY)-(LX, LY)
    Pic.Forecolor = Bolt.Inner1
    Pic.Line (SX - 1, SY)-((SX + LX) / 2 - 1, (SY + LY) / 2)
    Pic.Line (SX + 1, SY)-((SX + LX) / 2 + 1, (SY + LY) / 2)
    Pic.Forecolor = Bolt.Inner2
    Pic.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX - 1, LY)
    Pic.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX + 1, LY)
    Pic.Forecolor = Bolt.Outer1
    Pic.Line (SX - 2, SY)-((SX + LX) / 2 - 2, (SY + LY) / 2)
    Pic.Line (SX + 2, SY)-((SX + LX) / 2 + 2, (SY + LY) / 2)
    Pic.Forecolor = Bolt.Outer2
    Pic.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX - 2, LY)
    Pic.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX + 2, LY)
    If I >= Round((Bolt.Nicks / 2), 0) And Bolt.SplitBolt = True Then
        Pic.Forecolor = Bolt.Center
        Pic.Line (SX, SY)-(LX, LY)
        Pic.Forecolor = Bolt.Inner1
        Pic.Line (SX2 - 1, SY2)-((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)
        Pic.Line (SX2 + 1, SY2)-((SX2 + LX2) / 2 + 1, (SY2 + LY2) / 2)
        Pic.Forecolor = Bolt.Inner2
        Pic.Line ((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)-(LX2 - 1, LY2)
        Pic.Line ((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)-(LX2 + 1, LY2)
        Pic.Forecolor = Bolt.Outer1
        Pic.Line (SX2 - 2, SY2)-((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)
        Pic.Line (SX2 + 2, SY2)-((SX2 + LX2) / 2 + 2, (SY2 + LY2) / 2)
        Pic.Forecolor = Bolt.Outer2
        Pic.Line ((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)-(LX2 - 2, LY2)
        Pic.Line ((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)-(LX2 + 2, LY2)
    End If
    If I = Bolt.Nicks Then
        'Do Tail
        If Alter = True Then
            '-
            SX = LX
            SY = LY
            LX = LX - 1
            LY = LY + 2
        Else
            '+
            SX = LX
            SY = LY
            LX = LX + 1
            LY = LY + 2
        End If
        Pic.Forecolor = Bolt.Inner1
        Pic.Line (SX, SY)-(LX, LY)
        Pic.Forecolor = Bolt.Inner3
        Pic.Line (SX - 1, SY)-((SX + LX) / 2 - 1, (SY + LY) / 2)
        Pic.Line (SX + 1, SY)-((SX + LX) / 2 + 1, (SY + LY) / 2)
        Pic.Forecolor = Bolt.Inner4
        Pic.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX - 1, LY)
        Pic.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX + 1, LY)
        Pic.Forecolor = Bolt.Outer3
        Pic.Line (SX - 2, SY)-((SX + LX) / 2 - 2, (SY + LY) / 2)
        Pic.Line (SX + 2, SY)-((SX + LX) / 2 + 2, (SY + LY) / 2)
        Pic.Forecolor = Bolt.Outer4
        Pic.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX - 2, LY)
        Pic.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX + 2, LY)
        GoTo RefreshPic
    End If
    SX = LX
    SY = LY
    Alter = Int(Rnd * 2)
    If Alter = True Then
        Alter = False
        LX = LX + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
    Else
        Alter = True
        LX = LX - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
    End If
    LY = LY + Bolt.Height / Bolt.Nicks
    If I < Round((Bolt.Nicks / 2), 0) - 1 And Bolt.SplitBolt = True Then
        LX2 = SX
        LY2 = SY
    End If
    If I >= Round((Bolt.Nicks / 2), 0) - 1 And Bolt.SplitBolt = True Then
        SX2 = LX2
        SY2 = LY2
        Alter = Int(Rnd * 2)
        If Alter = True Then
            Alter = False
            LX2 = LX2 + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
        Else
            Alter = True
            LX2 = LX2 - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
        End If
        LY2 = LY2 + Bolt.Height / Bolt.Nicks
    End If
Next I
RefreshPic:
Pic.Refresh
End Function

Private Sub CmdExit_Click()
    Call ImplodeForm(Me, 2, 500, 1) 'displays cool unload form effect
End
End Sub

Private Sub CmdList_Click()
'Shows Credits
Call ShellExecute(Me.hwnd, vbNullString, App.Path & "\scripts\px credits.html", vbNullString, vbNullString, SW_SHOWNORMAL)


End Sub

Private Sub CmdMapEd_Click()
'sets options off for things not needed in map editor
OnlineGame = False
MapEd = True
TOOLS.CmdMapEd.ToolTipText = "Map Editor"
TOOLS.Command1.Visible = False
TOOLS.Command4.Visible = False
TOOLS.Command3.Visible = False
TOOLS.Command2.Visible = False
TOOLS.Frame6.Visible = True
TOOLS.Frame1.Visible = False
TOOLS.Frame2.Visible = False
TOOLS.Frame3.Visible = False
TOOLS.CmdGuard.Enabled = False
TOOLS.CmdAttack.Enabled = False
TOOLS.CmdScan.Enabled = False
TOOLS.CmdQuests.Enabled = False
TOOLS.CmdStealth.Enabled = False
TOOLS.NextLevel.Enabled = False
TOOLS.Show
Form1.Image1.Visible = False
Form1.Title.Left = 2999
Form1.Picture2.Visible = False
Form1.Meter.Visible = False
Form1.Title.Visible = False
Form1.Shape2.Visible = False
Form1.FrameStats.Visible = False
Form1.KeyPreview = False

Form1.txtDialog.Enabled = False
Form1.txtMessage.Enabled = False
Form1.cmdSend.Enabled = False
Form1.VScroll1.Visible = True
Form1.HScroll1.Visible = True

Form1.Image1.Visible = False
Form1.HScroll1.Value = "45"
Form1.VScroll1.Value = "23"
Form1.txtDialog.Text = ""
Unload Me

End Sub

Private Sub CmdSecond_Click()
'Connects without checking if server is up via the web
MainPlayer.ServerIP = ServerIP.Text
OnlineGame = True
Me.Hide

frmLogin.Show
End Sub

Private Sub CmdSinglePlayer_Click()
OnlineGame = False
Mutiplayer.CmdLogout.Enabled = False
MapEd = False
Mutiplayer.Show
Unload Me
End Sub

Private Sub Command1_Click()
If ServerIP.Text = "MainServer" Then
    Call MainServerConnect
Exit Sub
End If
MainPlayer.ServerIP = ServerIP.Text
Dim updatestring As String 'Holds the update information
Dim tInt As Integer 'Holds value of split
Dim WebAddress As String
Dim Connect As String
Me.Caption = "Connecting to Project X Online"
'Checks to see if there is game.
     WebAddress = "http://" & ServerIP.Text & "/connect.txt"
     updatestring = Inet1.OpenURL(WebAddress, icString)

tInt = InStr(1, updatestring, vbCrLf)
Connect = updatestring

If Connect = "You are Connected!" Then
    OnlineGame = True
    Me.Caption = Connect
    MainPlayer.ServerIP = ServerIP.Text
    Command1.Enabled = False
    MapEd = False
    Success% = FileExists%(App.Path & "\update\update.txt")
    If Success% = True Then
        Kill (App.Path & "\update\update.txt")
    End If
'show the download form and check for updates
  frmDownload.DownloadFile "http://" & ServerIP.Text & "/update.txt", _
                         App.Path & "\update\update.txt"
  frmDownload.Caption = "Checking for updates"
   Me.Hide
Else
    MsgBox "You are not connected! Check to make sure the ip is correct!", vbExclamation
    Me.Caption = "Not Connected!"

End If

End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

Randomize
'Set Center
Bolt.Center = RGB(207, 207, 233)
'Set Inner
Bolt.Inner1 = RGB(176, 176, 207)
Bolt.Inner2 = RGB(176, 176, 192)
Bolt.Inner3 = RGB(160, 160, 176)
Bolt.Inner4 = RGB(145, 145, 176)
'Set Outer
Bolt.Outer1 = RGB(90, 90, 110)
Bolt.Outer2 = RGB(80, 80, 110)
Bolt.Outer3 = RGB(65, 65, 80)
Bolt.Outer4 = RGB(50, 50, 70)
'Set Properties
Bolt.Nicks = 40 'NickNum.Value
Bolt.VelocityMin = 1
Bolt.VelocityMax = 40
Bolt.Height = 1300
Bolt.SameStart = True
Bolt.ShowCloud = True
Bolt.SplitBolt = True


On Error Resume Next
'Register a control
RegComCtl32
End Sub





Private Sub Label6_Click()
'Message Board for Project X Online
Call Shell("start.exe " & "http://x.incyberspace.net")
End Sub

Private Sub lblCompany_Click()
Call Shell("start.exe " & "http://www.visualbasiczone.com")
End Sub

Private Sub ServerIP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Command1_Click

End If
End Sub

Private Sub Timer1_Timer()
'Tells the user if is still trying to connect to the game
If Inet1.StillExecuting = True Then
    Me.Caption = "Excuting"
    Command1.Enabled = False
End If
If Inet1.StillExecuting = False Then
    Me.Caption = ""
    Command1.Enabled = True
End If
End Sub

Function FileExists%(fname$)
 On Local Error Resume Next
 
 Dim ff%
 ff% = FreeFile
 Open fname$ For Input As ff%
 If Err Then
  FileExists% = False
 Else
  FileExists% = True
 End If
 
 Close ff%
End Function

Private Sub TimerDraw_Timer()
'Draws the lightining on the world picture
DrawGlobe LogoPic
LogoPic.Refresh

End Sub

Private Sub MainServerConnect()
'Connects to the main server of the game!
Dim updatestring As String 'Holds the update information
Dim tInt As Integer 'Holds value of split
Dim WebAddress As String
Dim Connect As String
Me.Caption = "Connecting to Project X Online"
'Gets the game ip
     WebAddress = "http://x.incyberspace.net/ip.txt"
     updatestring = Inet1.OpenURL(WebAddress, icString)
tInt = InStr(1, updatestring, vbCrLf)
Connect = updatestring
If Connect = "" Then
    MsgBox "You are not connected! Check to make sure the ip is correct!", vbExclamation
    Me.Caption = "Not Connected!"
Else
   OnlineGame = True
   Me.Caption = "Connecting"
   MainPlayer.ServerIP = Connect
   Command1.Enabled = False
   MapEd = False
   Success% = FileExists%(App.Path & "\update\update.txt")
   If Success% = True Then
    Kill (App.Path & "\update\update.txt")
   End If
   'Main site for Project X checks if server is up
   frmDownload.DownloadFile "http://x.incyberspace.net/Update/update.text", _
                         App.Path & "\update\update.txt"
   frmDownload.Caption = "Checking for updates"
   Me.Hide


End If
End Sub


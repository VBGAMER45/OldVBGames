VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form Mutiplayer 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project X Online"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "&Play as Link"
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton CmdRandomHue 
      Caption         =   "Random Hue"
      Height          =   375
      Left            =   6240
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   1680
      OleObjectBlob   =   "Main.frx":030A
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   1440
      OleObjectBlob   =   "Main.frx":0388
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton CmdRemoveSkin 
      Caption         =   "Remove Skin"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton CmdOpenSkin 
      Caption         =   "Open Skin"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   2520
      OleObjectBlob   =   "Main.frx":0418
      Top             =   3600
   End
   Begin VB.TextBox txtCharSave 
      Height          =   405
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet InetCharList 
      Left            =   3000
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "What's New"
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   3615
      Begin VB.TextBox Text1 
         Height          =   3015
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "What's New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image BackImage2 
         Height          =   3495
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2400
      Width           =   1695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1920
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLogout 
      Caption         =   "&Logout"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton CmdWorld 
      Caption         =   "&Enter the World!"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1200
      Top             =   3120
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New Charecter"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete Charecter"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Exit the Game!"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Main.frx":061A
      Left            =   1680
      List            =   "Main.frx":061C
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox currrentusers 
      Height          =   1035
      ItemData        =   "Main.frx":061E
      Left            =   1680
      List            =   "Main.frx":0620
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1440
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Charecters!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current People Logged on!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Image BackImage 
      Height          =   4335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Mutiplayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response As String
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
Private Type OpenFilename
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    iFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim Filename As String


Private Function ShowFileDialog() As String
    Dim ofn As OpenFilename
    
    ofn.lpstrInitialDir = App.Path & "\skins\"
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hwnd
    ofn.lpstrFilter = "Skin files (*.skn)" & Chr$(0) & "*.skn" & Chr$(0) & Chr(0) & Chr(0)
    ofn.lpstrFile = String(256, 0)
    ofn.nMaxFile = 255
    ofn.lpstrTitle = "Open Skin"
    ofn.Flags = &H800000 + &H1000 + &H8 + &H4
    ofn.lpstrDefExt = "skn" + Chr(0)
    GetOpenFileName ofn
    If Mid(ofn.lpstrFile, 1, 1) <> Chr(0) Then ShowFileDialog = ofn.lpstrFile
End Function


Private Sub CmdDelete_Click()
response = MsgBox("Are You sure you want to delete " & List1.Text, vbYesNo, "Delete Charecter?")
If response = vbYes Then

If OnlineGame = True Then
'sends delete command to the server
 If frmLogin.Winsock1.State = sckConnected Then
  frmLogin.Winsock1.SendData "del" & "," & MainPlayer.AccountName & "," & List1.Text
    List1.RemoveItem List1.ListIndex
    Else
End If
End If
If OnlineGame = False Then
List1.RemoveItem List1.ListIndex

'OfflineSaves
Call SaveList(List1, App.Path & "\OfflineSaves\Charecters.txt")

End If
End If
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdLogout_Click()
If OnlineGame = False Then Exit Sub
'remove user from current users
Unload Me
frmLogin.Text1 = ""

frmLogin.txtPassword = ""
frmLogin.Show
End Sub

Private Sub CmdNew_Click()

Form5.Show
End Sub

Private Sub CmdOpenSkin_Click()
    Dim Filename As String
    Filename = ShowFileDialog
    If Filename <> "" Then
        SkinFilename = Filename
        Skn.LoadSkin Filename ' Loads another skin into Skin component
       Call stopmessage
        Skn.ApplySkin Me.hwnd ' Applies the skin to this window and its child controls
        Call stopmessage
        'Saves the skin name
        SaveSetting "Project X", "Startup", "Skin", Filename
        

    End If
End Sub

Private Sub CmdRandomHue_Click()
Dim img As ISkinImages
    Skn.GetImages.SetHue Rnd * 360, False
    
End Sub

Private Sub CmdRemoveSkin_Click()
  Skn.RemoveSkin ' Removes the skin from all the skinned windows
  SkinFilename = ""
  
End Sub

Private Sub CmdWorld_Click()

ZeldaMode = False
Timer1.Enabled = False
If OnlineGame = True Then

 If frmLogin.Winsock1.State = sckConnected Then
  frmLogin.Winsock1.SendData "pac" & "," & MainPlayer.AccountName & "," & List1.Text & "," & Winsock1.LocalIP
    
End If
End If


Call SaveList(List1, App.Path & "\OfflineSaves\" & MainPlayer.AccountName & "Charecters.txt")
MainPlayer.PlayerName = List1.Text

TOOLS.LoadCharecterInfo
TOOLS.Show


End Sub



Private Sub stopmessage()
On Error GoTo nofile
AppActivate "Unregistered"
SendKeys "{ENTER}"

nofile:
Exit Sub
End Sub


Private Sub Command5_Click()
Dim RupeeFormat As String
Timer1.Enabled = False
If OnlineGame = False Then
    ZeldaMode = True
    MainPlayer.PlayerName = "Link"
    Form1.Image1.Picture = LoadPicture(App.Path & "\monsters\link.gif")
    Form1.WindowState = 2
    TOOLS.Show
    Form1.Image1.Picture = LoadPicture(App.Path & "\monsters\link.gif")
    TOOLS.Hide
    Form1.WindowState = 2
    Form1.Title.Visible = False
    Form1.Caption = "The Legend of Zelda, Link's Destiny"
    Form1.FrameStats.Left = Form1.FrameStats.Left + 300
    Form1.cmdSend.Top = Form1.cmdSend.Top + 25
    Form1.txtMessage.Top = Form1.txtMessage.Top + 25
    Form1.txtDialog.Top = Form1.txtDialog.Top + 25
    Form1.txtMessage.Width = Form1.Width
    Form1.txtDialog.Width = 800
    RupeeFormat = Format(MainPlayer.Money, "000000")
    
    Form1.lblMoney.Caption = "Rupees " & RupeeFormat
    
    Form1.txtDialog.Text = Form1.txtDialog.Text & "Welcome to Link's Destiny Please try the the training tower next to the store" & vbCrLf
    Form1.txtDialog.Text = Form1.txtDialog.Text & "This game plays differently then Project X, so be sure to try both modes." & vbCrLf
    Form1.txtDialog.Text = Form1.txtDialog.Text & "To read a sign go up to it and press T when you are in front of it by one tile." & vbCrLf
Else
    MsgBox "Sorry you can only play Zelda in offline mode!"
End If

End Sub

Private Sub Form_Load()
Dim WebAddress As String
Dim updatestring As Long
Dim Filename As String
Dim fname As String
BackImage.Picture = LoadPicture(App.Path & "\wood.bmp")
If OnlineGame = True Then
Winsock1.RemotePort = "5774"
Winsock1.RemoteHost = MainPlayer.ServerIP
End If
  F = FreeFile


On Error GoTo nofile2
'Loads the What's New
Text1.Text = ""
Open App.Path & "\new.txt" For Input As F
Text1.Text = Input(LOF(F), F)
Close F
fname = GetSetting("Project X", "Startup", "Skin")

If fname <> "" Then
SkinFilename = fname
Skn.LoadSkin fname ' Loads another skin into Skin component
Call stopmessage
Skn.ApplySkin Me.hwnd
End If

WebAddress = "http://" & MainPlayer.ServerIP & "/users.txt"

Text2.Text = Inet1.OpenURL(WebAddress, icString)
If OnlineGame = False Then

'App.Path & "\OfflineSaves\
Call LoadList(List1, App.Path & "\OfflineSaves\Charecters.txt")
Else
If MainPlayer.MakeFolder = True Then
Else
'Loads player character list
 WebAddress = "http://" & MainPlayer.ServerIP & "/accounts/" & MainPlayer.AccountName & "/charecters.txt"
txtCharSave.Text = InetCharList.OpenURL(WebAddress, icString)



End If
End If
Mutiplayer.Text2.Text = Mutiplayer.Text2.Text & frmLogin.txtUserName.Text & vbCrLf


nofile2:
Exit Sub
End Sub



Private Sub Timer1_Timer()
If List1.Text = "" Then
CmdWorld.Enabled = False
CmdDelete.Enabled = False

Else
CmdWorld.Enabled = True
CmdDelete.Enabled = True
'Call loadmod
End If
If SkinFilename = "" Then
SkinLabel1.Visible = False
SkinLabel2.Visible = False
Else
SkinLabel1.Visible = True
SkinLabel2.Visible = True
Me.Refresh
End If
On Error GoTo nofile
AppActivate "Unregistered"
SendKeys "{ENTER}"

nofile:
Exit Sub

End Sub
Sub SaveList(Lst As ListBox, file As String)
'Call SaveList (List1,"C:\Windows\System\Saved.lst")
On Error GoTo error
Open file For Output As #1
For I = 0 To Lst.ListCount - 1
A$ = Lst.List(I)
Print #1, A$
Next
Close 1
Exit Sub
error:
X = MsgBox("There has been a error!", vbOKOnly, "Error")
End Sub
Sub LoadList(Lst As ListBox, file As String)
'Call LoadList (List1,"C:\Windows\System\Saved.lst")
On Error GoTo error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
Lst.AddItem A$
Loop
Close 1
Exit Sub
error:
Exit Sub
'x = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

Private Sub Winsock1_Connect()
Dim chartext As String
chartext = "newchar" & "," & MainPlayer.AccountName & "," & Form5.CharecterName.Text & "," & Form5.Picture1.Tag & "," & Form5.Health.Caption & "," & Form5.Stamina.Caption & "," & Form5.Mana.Caption & "," & Form5.Quickness.Caption & "," & Form5.Dex.Caption & "," & Form5.Strength.Caption & "," & Form5.Wisdom.Caption & "," & Form5.Intel.Caption
Winsock1.SendData chartext
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim incomdata As String
Winsock1.GetData incomdata
DataProcess incomdata
End Sub
Sub DataProcess(DataStr As String)
'Process Data From winsock control
'returns if the character was created
If Left(DataStr, 11) = "charcreated" Then

End If

If Left(DataStr, 9) = "charexist" Then
MsgBox "Your character was not created beacaue the character may already exist or you used improper characters in your character's name", vbExclamation
End If

End Sub

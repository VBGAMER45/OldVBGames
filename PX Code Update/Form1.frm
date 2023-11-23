VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   Caption         =   "New Charecter"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6465
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   2640
      OleObjectBlob   =   "Form1.frx":030A
      TabIndex        =   53
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   2640
      OleObjectBlob   =   "Form1.frx":038A
      TabIndex        =   52
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   3720
      OleObjectBlob   =   "Form1.frx":0404
      TabIndex        =   51
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   735
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0484
      TabIndex        =   50
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0592
      TabIndex        =   49
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinIntel 
      Height          =   255
      Left            =   4920
      OleObjectBlob   =   "Form1.frx":060C
      TabIndex        =   48
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinWisdom 
      Height          =   255
      Left            =   4920
      OleObjectBlob   =   "Form1.frx":0686
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinStrength 
      Height          =   255
      Left            =   3480
      OleObjectBlob   =   "Form1.frx":0700
      TabIndex        =   46
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinDex 
      Height          =   255
      Left            =   3480
      OleObjectBlob   =   "Form1.frx":077A
      TabIndex        =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinQuick 
      Height          =   255
      Left            =   3480
      OleObjectBlob   =   "Form1.frx":07F4
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinMana 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "Form1.frx":086E
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinStamina 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "Form1.frx":08DC
      TabIndex        =   42
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinHealth 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "Form1.frx":0950
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5880
      Top             =   1680
   End
   Begin VB.CommandButton Command14 
      Caption         =   "+"
      Height          =   255
      Left            =   6120
      TabIndex        =   40
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Height          =   255
      Left            =   6120
      TabIndex        =   39
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+"
      Height          =   255
      Left            =   4680
      TabIndex        =   38
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "+"
      Height          =   255
      Left            =   4680
      TabIndex        =   37
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "+"
      Height          =   255
      Left            =   4680
      TabIndex        =   36
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "+"
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "+"
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Apoints 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "100"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<"
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   2520
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   1425
      ItemData        =   "Form1.frx":09CA
      Left            =   3600
      List            =   "Form1.frx":09CC
      TabIndex        =   14
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Points 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "50"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":09CE
      Left            =   120
      List            =   "Form1.frx":0A17
      TabIndex        =   11
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton CmdDown 
      BackColor       =   &H00000040&
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton CmdUp 
      BackColor       =   &H00000040&
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox CharecterName 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   0
      Tag             =   "1"
      Top             =   120
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   3000
      OleObjectBlob   =   "Form1.frx":0B20
      Top             =   3720
   End
   Begin VB.Label Intel 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5760
      TabIndex        =   33
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Wisdom 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   32
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Strength 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Dex 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Quickness 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Mana 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   28
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Stamina 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Health 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Attribute Points:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   5880
      Picture         =   "Form1.frx":0D22
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   6240
      Picture         =   "Form1.frx":13E4
      Top             =   4200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   5280
      Picture         =   "Form1.frx":1926
      Top             =   4200
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   5640
      Picture         =   "Form1.frx":1F68
      Top             =   4200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4920
      Picture         =   "Form1.frx":2482
      Top             =   4200
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":2AC4
      Top             =   4200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Intelligence:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Wisdom: "
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Dexiterity: "
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Quickness:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Strength: "
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Skills Researched"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Skills Points:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mana:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stamina:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Skills:  I suggest you  go down the whole list of skills then chose the skills you want."
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Charecter Name"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response As String

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdDown_Click()
If Picture1.Tag < 18 Then
Picture1.Tag = Picture1.Tag + 1
End If
If Picture1.Tag = 2 Then
Picture1.Picture = Image2.Picture
End If
If Picture1.Tag = 3 Then
Picture1.Picture = Image3.Picture
End If
If Picture1.Tag = 4 Then
Picture1.Picture = Image4.Picture
End If
If Picture1.Tag = 5 Then
Picture1.Picture = Image5.Picture
End If
If Picture1.Tag = 6 Then
Picture1.Picture = Image6.Picture
End If
If Picture1.Tag = 7 Then
Picture1.Picture = Pics.Deer.Picture
End If
If Picture1.Tag = 8 Then
Picture1.Picture = Pics.Chicken.Picture
End If
If Picture1.Tag = 9 Then
Picture1.Picture = Pics.Leperchaun.Picture
End If
If Picture1.Tag = 10 Then
Picture1.Picture = Pics.King.Picture
End If
If Picture1.Tag = 11 Then
Picture1.Picture = Pics.RockMan.Picture
End If
If Picture1.Tag = 12 Then
Picture1.Picture = Pics.Mutant.Picture
End If
If Picture1.Tag = 13 Then
Picture1.Picture = Pics.Lion.Picture
End If
If Picture1.Tag = 14 Then
Picture1.Picture = Pics.LostMage.Picture
End If
If Picture1.Tag = 15 Then
Picture1.Picture = Pics.MagicTiger.Picture
End If
If Picture1.Tag = 16 Then
Picture1.Picture = Pics.Undead.Picture
End If
If Picture1.Tag = 17 Then
Picture1.Picture = Pics.Pirate.Picture
End If
End Sub

Private Sub CmdUp_Click()
If Picture1.Tag > 1 Then
Picture1.Tag = Picture1.Tag - 1
End If
If Picture1.Tag = 1 Then
Picture1.Picture = Image1.Picture
End If
If Picture1.Tag = 2 Then
Picture1.Picture = Image2.Picture
End If
If Picture1.Tag = 3 Then
Picture1.Picture = Image3.Picture
End If
If Picture1.Tag = 4 Then
Picture1.Picture = Image4.Picture
End If
If Picture1.Tag = 5 Then
Picture1.Picture = Image5.Picture
End If
If Picture1.Tag = 6 Then
Picture1.Picture = Image6.Picture
End If
If Picture1.Tag = 7 Then
Picture1.Picture = Pics.Deer.Picture
End If
If Picture1.Tag = 8 Then
Picture1.Picture = Pics.Chicken.Picture
End If
If Picture1.Tag = 9 Then
Picture1.Picture = Pics.Leperchaun.Picture
End If
If Picture1.Tag = 10 Then
Picture1.Picture = Pics.King.Picture
End If
If Picture1.Tag = 11 Then
Picture1.Picture = Pics.RockMan.Picture
End If
If Picture1.Tag = 12 Then
Picture1.Picture = Pics.Mutant.Picture
End If
If Picture1.Tag = 13 Then
Picture1.Picture = Pics.Lion.Picture
End If
If Picture1.Tag = 14 Then
Picture1.Picture = Pics.LostMage.Picture
End If
If Picture1.Tag = 15 Then
Picture1.Picture = Pics.MagicTiger.Picture
End If
If Picture1.Tag = 16 Then
Picture1.Picture = Pics.Undead.Picture
End If
If Picture1.Tag = 17 Then
Picture1.Picture = Pics.Pirate.Picture
End If
End Sub

Private Sub Command1_Click()
Dim chartext As String
Dim Loopc As Integer, Charascii As Integer
  For Loopc = 1 To Len(CharecterName.Text)
    Charascii = Asc(Mid$(CharecterName.Text, Loopc, 1))
    If Server.LegalCharacter(Charascii) = False Then
        MsgBox ("Invalid Username because of use of specail characters")
       Exit Sub
    
    End If
    
Next Loopc

response = MsgBox("Remeber to use only letters and numbers as your name.  If you need to CHANGE it press YES", vbYesNo, "Fix Name?")
If response = vbYes Then
Exit Sub
Exit Sub
End If

CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")

If Health.Caption > 110 Or Mana.Caption > 110 Or Stamina.Caption > 120 Or Wisdom.Caption > 110 Or Intel.Caption > 110 _
Or Strength.Caption > 110 Or Dex.Caption > 110 Or Quickness.Caption > 110 Then
MsgBox "Cheaters Never Prosper!", vbExclamation, "Cheater!!!!"
Exit Sub
End If
If CharecterName.Text = "" Then
MsgBox "Please enter a charecter Name!", vbExclamation, "New Charecter!"
Else

If Points.Text = 0 Then
Mutiplayer.List1.AddItem CharecterName
'Create Player Code
'## server stuff

CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
    If OnlineGame = True Then
    Mutiplayer.txtCharSave.Text = CharecterName.Text
    chartext = "newchar" & "," & MainPlayer.AccountName & "," & CharecterName.Text & "," & Picture1.Tag & "," & Health.Caption & "," & Stamina.Caption & "," & Mana.Caption & "," & Quickness.Caption & "," & Dex.Caption & "," & Strength.Caption & "," & Wisdom.Caption & "," & Intel.Caption
    'Mutiplayer.Winsock1.SendData chartext
   
    If frmLogin.Winsock1.State = sckConnected Then
    frmLogin.Winsock1.SendData chartext
    frmLogin.Winsock1.SendData GetSkills
    'Mutiplayer.Winsock1.SendData chartext
     '''  Unload Me
       ''Exit Sub
    Else
    frmLogin.Winsock1.Connect
  ' Mutiplayer.Winsock1.Connect
    'Mutiplayer.Winsock1.SendData chartext
       '' Unload Me
       ''Exit Sub
    End If
    End If
'## end server stuff


SaveList List1, App.Path & "\OfflineSaves\" & CharecterName.Text & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & CharecterName.Text & "skills.txt"
Dim outputfile As Integer
MainPlayer.Admin = False
outputfile = 1
Open App.Path & "\OfflineSaves\" & CharecterName.Text & ".char" For Output As #outputfile
Write #outputfile, 1, MainPlayer.Admin, 0, 1000, "newbie", 50, 50
Write #outputfile, 10, 10, Int(Picture1.Tag), 0, Int(Health.Caption), Int(Stamina.Caption), Int(Mana.Caption)
Write #outputfile, Int(Health.Caption), Int(Stamina.Caption), Int(Mana.Caption), Int(Strength.Caption), Int(Quickness.Caption), Int(Dex.Caption), Int(Wisdom.Caption)
Write #outputfile, Int(Intel.Caption), 500, 0, 0, 0, 100, 0, 0
Write #outputfile, 0, 0, 0, 0, 0, 0, 0, 0, 0
Write #outputfile, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
Write #outputfile, 0, 0, 0, 0, 0, 0
Close #outputfile


Unload Me

Else

response = MsgBox("You have not used up all your points!  Do you still want to create a new charecter?", vbYesNo, "New Charecter?")
If response = vbYes Then
CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
Mutiplayer.List1.AddItem CharecterName
'Create Player Code
'## server stuff

CharecterName.Text = modGlobals.sReplace(CharecterName.Text, "@", "Cheater")
    If OnlineGame = True Then
    Mutiplayer.txtCharSave.Text = CharecterName.Text
    chartext = "newchar" & "," & MainPlayer.AccountName & "," & CharecterName.Text & "," & Picture1.Tag & "," & Health.Caption & "," & Stamina.Caption & "," & Mana.Caption & "," & Quickness.Caption & "," & Dex.Caption & "," & Strength.Caption & "," & Wisdom.Caption & "," & Intel.Caption
    If frmLogin.Winsock1.State = sckConnected Then
   ' frmLogin.Winsock1.SendData chartext
  frmLogin.Winsock1.SendData chartext
   frmLogin.Winsock1.SendData GetSkills
   '' Unload Me
      ''  Exit Sub
    Else
     'frmLogin.Winsock1.SendData chartext
    '    frmLogin.Winsock1.Connect
   '' frmLogin.Winsock1.Connect
   '' Unload Me
        Exit Sub
'Mutiplayer.Winsock1.SendData chartext
    End If
    End If

SaveList List1, App.Path & "\OfflineSaves\" & CharecterName.Text & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & CharecterName.Text & "skills.txt"
outputfile = 1
Open App.Path & "\OfflineSaves\" & CharecterName.Text & ".char" For Output As #outputfile
Write #outputfile, 1, MainPlayer.Admin, 0, 1000, "newbie", 50, 50
Write #outputfile, 10, 10, Int(Picture1.Tag), 0, Int(Health.Caption), Int(Stamina.Caption), Int(Mana.Caption)
Write #outputfile, Int(Health.Caption), Int(Stamina.Caption), Int(Mana.Caption), Int(Strength.Caption), Int(Quickness.Caption), Int(Dex.Caption), Int(Wisdom.Caption)
Write #outputfile, Int(Intel.Caption), 500, 0, 0, 0, 100, 0, 0
Write #outputfile, 0, 0, 0, 0, 0, 0, 0, 0, 0
Write #outputfile, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
Write #outputfile, 0, 0, 0, 0, 0, 0
Close #outputfile

Unload Me

End If

End If
End If

End Sub

Private Sub Command10_Click()
If Apoints.Text > 0 Then
Apoints.Text = Apoints.Text - 1
Quickness.Caption = Quickness.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels
End Sub

Private Sub Command11_Click()
If Apoints.Text > 0 Then
Apoints.Text = Apoints.Text - 1
Dex.Caption = Dex.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels
End Sub

Private Sub Command12_Click()
If Apoints.Text > 0 Then
Apoints.Text = Apoints.Text - 1
Strength.Caption = Strength.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels
End Sub

Private Sub Command13_Click()
If Apoints.Text > 0 Then
Apoints.Text = Apoints.Text - 1
Wisdom.Caption = Wisdom.Caption + 1
Mana.Caption = Mana.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels

End Sub

Private Sub Command14_Click()
If Apoints.Text > 0 Then
Apoints.Text = Apoints.Text - 1
Intel.Caption = Intel.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels
End Sub

Private Sub Command5_Click()
If List1.Text = "" Then

Else
If List1.Text = "Melee Defense 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Run 10" Then
If Points.Text >= 10 Then
Points.Text = Points.Text - 10

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If


If List1.Text = "Cast Spell 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Swim 4" Then
If Points.Text >= 4 Then
Points.Text = Points.Text - 4

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Scan 4" Then
If Points.Text >= 4 Then
Points.Text = Points.Text - 4

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If

If List1.Text = "Stealth 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Sword 18" Then
If Points.Text >= 18 Then
Points.Text = Points.Text - 18

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Dagger 8" Then
If Points.Text >= 8 Then
Points.Text = Points.Text - 8

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Fail 10" Then
If Points.Text >= 10 Then
Points.Text = Points.Text - 10

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Staff 12" Then
If Points.Text >= 12 Then
Points.Text = Points.Text - 12

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Spear 10" Then
If Points.Text >= 10 Then
Points.Text = Points.Text - 10

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "The Force 30" Then
If Points.Text >= 30 Then
Points.Text = Points.Text - 30

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Unarmed Attack 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Bow 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Bombs 16" Then
If Points.Text >= 16 Then
Points.Text = Points.Text - 16

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "BlackSmith 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Miner 18" Then
If Points.Text >= 18 Then
Points.Text = Points.Text - 18

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Hunter 14" Then
If Points.Text >= 14 Then
Points.Text = Points.Text - 14

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Farmer 22" Then
If Points.Text >= 22 Then
Points.Text = Points.Text - 22

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Builder 16" Then
If Points.Text >= 16 Then
Points.Text = Points.Text - 16

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Magic Defense 24" Then
If Points.Text >= 24 Then
Points.Text = Points.Text - 24

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Mace 6" Then
If Points.Text >= 6 Then
Points.Text = Points.Text - 6

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If
If List1.Text = "Axe 6" Then
If Points.Text >= 6 Then
Points.Text = Points.Text - 6

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
End If
End If


End If

End Sub

Private Sub Command6_Click()
If List2.Text = "" Then

Else

If List2.Text = "Melee Defense 20" Then

Points.Text = Points.Text + 20

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex

End If
If List2.Text = "Run 10" Then

Points.Text = Points.Text + 10

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex

End If

If List2.Text = "Cast Spell 20" Then

Points.Text = Points.Text + 20

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Swim 4" Then

Points.Text = Points.Text + 4

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex

End If


If List2.Text = "Scan 4" Then

Points.Text = Points.Text + 4

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
'''''
If List2.Text = "Stealth 20" Then
'If Points.Text >= 20 Then
Points.Text = Points.Text + 20

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Sword 18" Then
'If Points.Text >= 18 Then
Points.Text = Points.Text + 18

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Dagger 8" Then
'If Points.Text >= 8 Then
Points.Text = Points.Text + 8

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Fail 10" Then
'If Points.Text >= 10 Then
Points.Text = Points.Text + 10

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Staff 12" Then
'If Points.Text >= 12 Then
Points.Text = Points.Text + 12

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Spear 10" Then
'If Points.Text >= 10 Then
Points.Text = Points.Text + 10

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "The Force 30" Then
'If Points.Text >= 30 Then
Points.Text = Points.Text + 30

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Unarmed Attack 20" Then
'If Points.Text >= 20 Then
Points.Text = Points.Text + 20

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
' If
End If
If List2.Text = "Bow 20" Then
'If Points.Text >= 20 Then
Points.Text = Points.Text + 20

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Bombs 16" Then
'If Points.Text >= 16 Then
Points.Text = Points.Text + 16

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "BlackSmith 20" Then
'If Points.Text >= 20 Then
Points.Text = Points.Text + 20

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Miner 18" Then
'If Points.Text >= 18 Then
Points.Text = Points.Text + 18

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Hunter 14" Then
'If Points.Text >= 14 Then
Points.Text = Points.Text + 14
List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Farmer 22" Then
'If Points.Text >= 22 Then
Points.Text = Points.Text + 22
List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex
'End If
End If
If List2.Text = "Builder 16" Then
'If Points.Text >= 16 Then
Points.Text = Points.Text + 16

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex

End If

If List2.Text = "Magic Defense 24" Then

Points.Text = Points.Text + 24

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex

End If
If List2.Text = "Mace 6" Then

Points.Text = Points.Text + 6

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex

End If
If List2.Text = "Axe 6" Then

Points.Text = Points.Text + 6

List1.AddItem List2.Text
List2.RemoveItem List2.ListIndex

End If


End If
End Sub

Private Sub Command7_Click()
If Apoints.Text > 0 Then

Apoints.Text = Apoints.Text - 1
Health.Caption = Health.Caption + 1
Stamina.Caption = Stamina.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels
End Sub

Private Sub Command8_Click()
If Apoints.Text > 0 Then
Apoints.Text = Apoints.Text - 1
Stamina.Caption = Stamina.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels
End Sub

Private Sub Command9_Click()
If Apoints.Text > 0 Then
Apoints.Text = Apoints.Text - 1
Mana.Caption = Mana.Caption + 1
Else
MsgBox "Not Enough attribute points!", vbExclamation, "No points left!"

End If
Call Refreshlabels
End Sub

Private Sub Form_Load()

Picture1.Picture = Image1.Picture
If SkinFilename = "" Then Exit Sub
  Skn.LoadSkin SkinFilename ' Loads another skin into Skin component
  Skn.ApplySkin Me.hwnd
Call Refreshlabels


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

Private Sub Form_Unload(Cancel As Integer)
 Call ImplodeForm(Me, 2, 500, 1)
End Sub

Private Sub Timer1_Timer()
If Points.Text > 50 Then
Points.Text = 0
End If
If Apoints.Text > 100 Then
Apoints.Text = 0
End If


If SkinFilename = "" Then
SkinHealth.Visible = False
SkinStamina.Visible = False
SkinMana.Visible = False
SkinDex.Visible = False
SkinQuick.Visible = False
SkinWisdom.Visible = False
SkinStrength.Visible = False
SkinIntel.Visible = False
SkinLabel1.Visible = False
SkinLabel2.Visible = False
SkinLabel3.Visible = False
SkinLabel4.Visible = False
SkinLabel5.Visible = False
Else
SkinHealth.Visible = True
SkinStamina.Visible = True
SkinMana.Visible = True
SkinDex.Visible = True
SkinQuick.Visible = True
SkinWisdom.Visible = True
SkinStrength.Visible = True
SkinIntel.Visible = True
SkinLabel1.Visible = True
SkinLabel2.Visible = True
SkinLabel3.Visible = True
SkinLabel4.Visible = True
SkinLabel5.Visible = True

End If
End Sub
Private Sub Refreshlabels()
SkinHealth.Caption = "Health: " & Health.Caption
SkinStamina.Caption = "Stamina: " & Stamina.Caption
SkinMana.Caption = "Mana: " & Mana.Caption
SkinDex.Caption = "Dexiterity: " & Dex.Caption
SkinQuick.Caption = "Quickness: " & Quickness.Caption
SkinStrength.Caption = "Strength: " & Strength.Caption
SkinWisdom.Caption = "Wisdom: " & Wisdom.Caption
SkinIntel.Caption = "Intelligence: " & Intel.Caption

End Sub
Private Function GetSkills() As String
'gets the player skills to send over winsock
GetSkills = GetSkills & List2.ListCount & "," & MainPlayer.AccountName & "," & CharecterName.Text
I = 0
For I = I + 1 To List2.ListCount
GetSkills = GetSkills & "," & List2.List(I)
Next
End Function

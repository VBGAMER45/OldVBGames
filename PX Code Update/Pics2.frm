VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Pics 
   Caption         =   "Pics"
   ClientHeight    =   4530
   ClientLeft      =   1845
   ClientTop       =   1590
   ClientWidth     =   5595
   LinkTopic       =   "Form3"
   ScaleHeight     =   4530
   ScaleWidth      =   5595
   Begin VB.Frame FrameMsg 
      Caption         =   "Message Board"
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton CmdClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   4680
         TabIndex        =   18
         Top             =   3240
         Width           =   855
      End
      Begin VB.ListBox Posts 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   2865
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   4995
      End
      Begin VB.TextBox Subject 
         Height          =   285
         Left            =   360
         TabIndex        =   15
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "Post"
         Height          =   375
         Left            =   4680
         TabIndex        =   14
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Messagetxt 
         Height          =   285
         Left            =   360
         TabIndex        =   13
         Top             =   3960
         Width           =   4095
      End
      Begin VB.Image BackImage 
         Height          =   5895
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   1335
      End
   End
   Begin VB.Frame FramePC 
      Caption         =   "Computer System"
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Frame Web 
         Caption         =   "Web Browser"
         Height          =   4455
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   255
            Left            =   4440
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   2895
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   5175
            ExtentX         =   9128
            ExtentY         =   5106
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clos&e"
            Height          =   375
            Left            =   1680
            TabIndex        =   8
            Top             =   3960
            Width           =   1935
         End
         Begin VB.TextBox txtPlace 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label Label1 
            Caption         =   "Address"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Logout"
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send Command"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtsend 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtpc 
         Height          =   2295
         Left            =   600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "Pics2.frx":0000
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Image Star 
      Height          =   585
      Left            =   3240
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image DeadChicken 
      Height          =   270
      Left            =   360
      Picture         =   "Pics2.frx":001A
      Top             =   2400
      Width           =   330
   End
   Begin VB.Image Dragon 
      Height          =   480
      Left            =   2280
      Picture         =   "Pics2.frx":03D4
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Queen 
      Height          =   465
      Left            =   2280
      Picture         =   "Pics2.frx":08AE
      Top             =   120
      Width           =   405
   End
   Begin VB.Image Priest 
      Height          =   480
      Left            =   2760
      Picture         =   "Pics2.frx":0DA7
      Top             =   120
      Width           =   435
   End
   Begin VB.Image Leperchaun 
      Height          =   465
      Left            =   2280
      Picture         =   "Pics2.frx":12A3
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image King 
      Height          =   480
      Left            =   1800
      Picture         =   "Pics2.frx":17C0
      Top             =   120
      Width           =   420
   End
   Begin VB.Image MagicTiger 
      Height          =   465
      Left            =   1800
      Picture         =   "Pics2.frx":1CF0
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Pirate 
      Height          =   465
      Left            =   1200
      Picture         =   "Pics2.frx":21BF
      Top             =   2280
      Width           =   435
   End
   Begin VB.Image Undead 
      Height          =   465
      Left            =   4200
      Picture         =   "Pics2.frx":26B4
      Top             =   1680
      Width           =   450
   End
   Begin VB.Image Mutant 
      Height          =   480
      Left            =   3600
      Picture         =   "Pics2.frx":2BC1
      Top             =   1680
      Width           =   435
   End
   Begin VB.Image LostMage 
      Height          =   450
      Left            =   3000
      Picture         =   "Pics2.frx":30AC
      Top             =   1680
      Width           =   435
   End
   Begin VB.Image RockMan 
      Height          =   450
      Left            =   1800
      Picture         =   "Pics2.frx":3580
      Top             =   1680
      Width           =   465
   End
   Begin VB.Image Lion 
      Height          =   450
      Left            =   1320
      Picture         =   "Pics2.frx":3AB2
      Top             =   1680
      Width           =   435
   End
   Begin VB.Image Deer 
      Height          =   450
      Left            =   2400
      Picture         =   "Pics2.frx":3F43
      Top             =   1680
      Width           =   435
   End
   Begin VB.Image Chicken 
      Height          =   270
      Left            =   840
      Picture         =   "Pics2.frx":43A0
      Top             =   1680
      Width           =   330
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Pics2.frx":474B
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   0
      Picture         =   "Pics2.frx":4B74
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   0
      Picture         =   "Pics2.frx":4F5B
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "Pics2.frx":534A
      Top             =   1440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   480
      Picture         =   "Pics2.frx":5742
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   480
      Picture         =   "Pics2.frx":5B2C
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Pics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You may be wondering what this is. This is a minigame
'using computers can be found in 7.4.map behind a bookself hint hint
'talk to the wiseman in newbie for a password or cheat and look at the code

Dim CorrectPass As Boolean
Dim CorrectPass2 As Boolean
Dim Portol As String
'Below holds cords for teleporting
Dim Xcord As Integer, Ycord As Integer, Ecord As Integer

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub CmdGo_Click()
If txtPlace.Text = "www.legacy.com" Then
    WebBrowser1.Navigate ("file://" & App.Path & "/scripts/legacy.htm")
Else
    WebBrowser1.Navigate ("file://" & App.Path & "/scripts/404.htm")
End If
End Sub

Private Sub CmdPost_Click()
Dim Data As String
If Messagetxt.Text = "" Then
    MsgBox "Please Type a Message No commas!!!!!)"
Else
If Subject.Text = "" Then
    Subject.Text = "No Subject"
End If
Data = Subject.Text & "," & Messagetxt.Text & "," & MainPlayer.PlayerName
Call Form1.SendClientData("Post", Data)
MsgBox "Your Message has been posted! To see it close the message board and open it again!", vbInformation
Unload Me

End If
End Sub

Private Sub Command1_Click()
If txtsend.Text = "z123b4" Then ' And MainPlayer.MapName = "c1.map" Then
txtpc.Text = ""
txtpc.Text = txtpc.Text & "Welcome to the Teleportation system" & vbCrLf
txtpc.Text = txtpc.Text & "Please enter the number to select your choice" & vbCrLf
txtpc.Text = txtpc.Text & "1 Open Partol" & vbCrLf
txtpc.Text = txtpc.Text & "2 Database" & vbCrLf
txtpc.Text = txtpc.Text & "3 List of Commands" & vbCrLf
txtpc.Text = txtpc.Text & "4 About" & vbCrLf
CorrectPass = True


ElseIf txtsend.Text = "baron@legacy.com:pass=r2d2" Then ' And MainPlayer.MapName = "c1.map" Then
txtpc.Text = ""
txtpc.Text = txtpc.Text & "Welcome to the Legacy Email System" & vbCrLf
txtpc.Text = txtpc.Text & "Please enter the number to select your choice" & vbCrLf
txtpc.Text = txtpc.Text & "1 Read Mail" & vbCrLf
txtpc.Text = txtpc.Text & "2 Compose" & vbCrLf
txtpc.Text = txtpc.Text & "3 Logout" & vbCrLf
CorrectPass2 = True
ElseIf txtsend.Text = "learning" Then ' And MainPlayer.MapName = "c1.map" Then
    'Used for training map on newbie
    txtpc.Text = ""
    txtpc.Text = txtpc.Text & "Welcome to the Legacy ComputerSystem" & vbCrLf
    txtpc.Text = txtpc.Text & "Well done!  You have learned how to use computers in the game they play an important role." & vbCrLf
ElseIf txtsend.Text = "/web" Then
WebBrowser1.Navigate ("file://" & App.Path & "/scripts/legacy.htm")
Web.Visible = True
Else
'If CorrectPass = True Then
'Call TeleSystem
'Exit Sub
'End If
'If CorrectPass2 = True Then
'Exit Sub
'End If
If CorrectPass2 = False And CorrectPass = False Then
txtpc.Text = txtpc.Text & "Invaild password!" & vbCrLf
End If

End If
Call EmailSystem

Call TeleSystem

txtsend.Text = ""
End Sub
Private Sub ClosePartol()
Map(17, 10).TileX = 10656
Map(17, 10).Tiley = 64
Map(17, 10).Tiletype = 0
Map(18, 10).TileX = 10688
Map(18, 10).Tiley = 64
Map(18, 10).Tiletype = 0
Map(19, 10).TileX = 10720
Map(19, 10).Tiley = 64
Map(19, 10).Tiletype = 0
Map(17, 9).TileX = 10656
Map(17, 9).Tiley = 32
Map(17, 9).Tiletype = 0
Map(18, 9).TileX = 10688
Map(18, 9).Tiley = 32
Map(18, 9).Tiletype = 0
Map(19, 9).TileX = 10720
Map(19, 9).Tiley = 32
Map(19, 9).Tiletype = 0
Form1.RedrawMap
End Sub
Private Sub OpenPartol()
Map(17, 10).TileX = 10752
Map(17, 10).Tiley = 64
Map(17, 10).Tiletype = 1
Map(18, 10).TileX = 10784
Map(18, 10).Tiley = 64
Map(18, 10).Tiletype = 1
Map(19, 10).TileX = 10816
Map(19, 10).Tiley = 64
Map(19, 10).Tiletype = 1
Map(17, 9).TileX = 10752
Map(17, 9).Tiley = 32
Map(17, 9).Tiletype = 1
Map(18, 9).TileX = 10784
Map(18, 9).Tiley = 32
Map(18, 9).Tiletype = 1
Map(19, 9).TileX = 10816
Map(19, 9).Tiley = 32
Map(19, 9).Tiletype = 1
Form1.RedrawMap
End Sub
Private Sub Command2_Click()
CorrectPass = False
CorrectPass2 = False

Unload Me

End Sub

Private Sub Command3_Click()
If Pics.Caption = "Project X Online Credits" Then
Unload Me
End If
Web.Visible = False
End Sub

Private Sub Form_Load()
CorrectPass = False
txtpc.Text = ""
txtpc.Text = txtpc.Text & "Please enter the password" & vbCrLf
'BackImage.Picture = LoadPicture("msgboard.jpg")

End Sub

Private Sub Form_Unload(Cancel As Integer)
CorrectPass = False
CorrectPass2 = False

End Sub

Private Sub Posts_Click()
Dim postnumber As Integer
I = -0
'I'f Posts.ListIndex > -1 Then
For I = I + 1 To Posts.ListCount
If Posts.Text = Messageboard(I).Subject Then


Form4.Caption = Messageboard(I).Subject
Form4.FramePost.Visible = True
  Form4.Show
    postnumber = Pics.Posts.ListIndex + 1
    Form4.PostContent = Messageboard(I).Post
    Form4.postedby = "Posted By " & Messageboard(I).Author
'End If
End If
Next I
End Sub

Private Sub txtPlace_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
CmdGo_Click

End If
End Sub

Private Sub txtsend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Command1_Click

End If
End Sub
Private Sub TeleSystem()
If txtsend.Text = "4" And CorrectPass = True Then
    txtpc.Text = txtpc.Text & "I was developed by Dr. Baron in the town of Thanatos" & vbCrLf
End If
If txtsend.Text = "3" And CorrectPass = True Then
    txtpc.Text = txtpc.Text & "Commands" & vbCrLf
    txtpc.Text = txtpc.Text & "/reboot" & vbCrLf
    txtpc.Text = txtpc.Text & "/closeportal" & vbCrLf
    txtpc.Text = txtpc.Text & "/setportal" & vbCrLf
    txtpc.Text = txtpc.Text & "/openportal" & vbCrLf
End If
If txtsend.Text = "/reboot" And CorrectPass = True Then
    CorrectPass = False
    txtpc.Text = ""
    txtpc.Text = txtpc.Text & "Please enter the password" & vbCrLf

End If
If txtsend.Text = "/closeportal" And CorrectPass = True Then
    Call ClosePartol
End If
If txtsend.Text = "/setportal" And CorrectPass = True Then
'<x>127</x><y>0</y><e>0</e>

    txtpc.Text = ""
    txtsend.Tag = "set"
    txtpc.Text = txtpc.Text & "Please enter portol cordinates in this format <x>cord1here</x><y>cord2here</y><e>cord3here</e>" & vbCrLf
    txtsend.Text = ""
Exit Sub
End If
If txtsend.Text = "1" And CorrectPass = True Or txtsend.Text = "/openportal" And CorrectPass = True Then
Call OpenPartol
End If
If txtsend.Tag = "set" And CorrectPass = True Then
On Error Resume Next
'Portol = txtsend.Text
Xcord = Extract(txtsend.Text, "<x>", "</x>")
Ycord = Extract(txtsend.Text, "<y>", "</y>")
Ecord = Extract(txtsend.Text, "<e>", "</e>")
'Xcords holds home many trees
'Y cord holds how many flowers
If Xcord > 1000 Then
Xcord = 1000
End If
If Ycord > 1000 Then
Ycord = 1000
End If
If Ecord > 100 Then
Ecord = 1
Else
Ecord = 2

End If
txtpc.Text = txtpc.Text & "Warp Portal set to cordinates " & Xcord & "." & Ycord & "." & Ecord & vbCrLf

End If
If txtsend.Text = "2" And CorrectPass = True Then
txtpc.Text = txtpc.Text & "The database is  down right now please repair the computer in Thanatos in order to get the database running. " & vbCrLf

End If
txtsend.Tag = ""
End Sub
Private Sub EmailSystem()
If txtsend.Text = "/reboot" And CorrectPass2 = True Or txtsend.Text = "3" And CorrectPass2 = True Then
CorrectPass = False
CorrectPass2 = False
txtpc.Text = ""
txtpc.Text = txtpc.Text & "Please enter the password" & vbCrLf

End If

End Sub
Public Function Extract(ByVal TextIN As String, Optional StartTag As String = " ", Optional ByVal EndTag As String = " ", Optional ByVal CheckCase As Boolean) As String
    On Error GoTo LocalError
    ' Extracts Text from string using start and end "tags"
    'NB:    If EndTag is ommitted the entire string from:
    '       StartTag to EndOfString is returned...
    Dim lArray As Variant
    Extract = ""
    lArray = Split(TextIN, StartTag)
    If IsArray(lArray) Then
        Extract = lArray(1)
        lArray = Split(Extract, EndTag)
        If IsArray(lArray) Then
            Extract = lArray(0)
        Else
            Extract = ""
        End If
    End If
Exit Function
LocalError:
    Extract = ""
End Function
Public Sub TeleportPortal()
'Right now just generates a random map!
Dim Number As Integer
Dim MapType As String
Randomize

Number = Rnd * 10
Select Case Number
    
    Case 0:
    MapType = "forest"
    Case 1:
    MapType = "desert"
    Case 2:
    MapType = "snow"
    Case 3:
    MapType = "light forest"
    Case 4:
    MapType = "deadforest"
    Case 5:
    MapType = "town"
    If Xcord > 20 Then
    Xcord = 20
    End If
    
    Case 6:
    MapType = "deadforest"
    Case 7:
     MapType = "light forest"
    Case 8:
    MapType = "desert"
    Case 9:
    MapType = "snow"
    Case 10:
    MapType = "forest"
End Select
'If Ecord = 1 Then
'BitBlt Form1.Tiles.hdc, 0, 0, TOOLS.picDisplay.Width, TOOLS.picDisplay.Height, TOOLS.picDisplay.hdc, 0, 0, SRCINVERT
'End If
'Invert (Form1.Tiles)
Call ModRndMap.RandomMapGenerator(MapType, Xcord, False, Ycord, True)
MainPlayer.MapName = MapType
Form1.MapName = MapType
Stuff.CenterPerson


End Sub
Public Sub Invert(PicBox)
'Invert the image of a picturebox
Dim newVal As Integer, H As Long, W As Long, K As Integer
Dim C As Long
Dim opRed As Long, opBlue As Long, opGreen As Long

Form1.Tiles.ScaleMode = 3

For H = 0 To Form1.Tiles.ScaleHeight
    For W = 0 To Form1.Tiles.ScaleWidth

        C = GetPixel(Form1.Tiles.hDC, W, H)
        RGBfromLONG C
        opRed = 255 - rRed
        opGreen = 255 - rGreen
        opBlue = 255 - rBlue
        C = RGB(opRed, opGreen, opBlue)
        SetPixel Form1.Tiles.hDC, W, H, C
    Next W
Next H
End Sub


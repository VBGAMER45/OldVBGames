VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AddMonster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Extra Stuff"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "AddMonster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameSpecial 
      Caption         =   "Extra Map Data"
      Height          =   5775
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Frame FrameExtra 
         Caption         =   "Extra Map stuff"
         Height          =   2055
         Left            =   120
         TabIndex        =   72
         Top             =   3000
         Width           =   3495
         Begin VB.TextBox txtNotes 
            Height          =   285
            Left            =   240
            TabIndex        =   75
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtMapTiltle 
            Height          =   285
            Left            =   240
            TabIndex        =   74
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtTiles 
            Height          =   285
            Left            =   240
            TabIndex        =   73
            Text            =   "tiles.bmp"
            Top             =   1680
            Width           =   2775
         End
         Begin VB.Label lblNotes 
            Caption         =   "Notes:"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblMapTitle 
            Caption         =   "MapTitle:"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblTileset 
            Caption         =   "Tileset Filename:"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   1440
            Width           =   1935
         End
      End
      Begin VB.TextBox TxTtop 
         Height          =   285
         Left            =   600
         MaxLength       =   3
         TabIndex        =   71
         Text            =   "0"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox TxTLeft 
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   70
         Text            =   "0"
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   2160
         TabIndex        =   69
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Done"
         Height          =   495
         Left            =   240
         TabIndex        =   68
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox txtHscroll 
         Height          =   285
         Left            =   2330
         MaxLength       =   3
         TabIndex        =   63
         Text            =   "0"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtVscroll 
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   62
         Text            =   "0"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox TxtWarpMap 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   60
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AddMonster.frx":030A
         Left            =   1200
         List            =   "AddMonster.frx":0317
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Gold"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox Txtmine 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   56
         Text            =   "0"
         Top             =   3720
         Width           =   2295
      End
      Begin VB.TextBox txtSound 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   54
         Text            =   "none"
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtmusic 
         Height          =   285
         Left            =   1080
         TabIndex        =   52
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AddMonster.frx":032F
         Left            =   1440
         List            =   "AddMonster.frx":033F
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Warp"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Txtvender 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   49
         Text            =   "none"
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox Tiley 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   48
         Text            =   "0"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox TileX 
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         MaxLength       =   3
         TabIndex        =   46
         Text            =   "0"
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "AddMonster.frx":035D
         Left            =   1080
         List            =   "AddMonster.frx":036D
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "None"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtBright 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   40
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label26 
         Caption         =   "X"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Y"
         Height          =   255
         Left            =   960
         TabIndex        =   66
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Hscroll"
         Height          =   255
         Left            =   1800
         TabIndex        =   65
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Vscroll"
         Height          =   255
         Left            =   2740
         TabIndex        =   64
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Warp Map"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Mine Type"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Mine Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Sound Effect"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Music"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Tile Type"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "TileY"
         Height          =   255
         Left            =   1320
         TabIndex        =   47
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Tilex"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Below is special stuff for a Tile!"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label13 
         Caption         =   "Vender Item"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Weather"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Brightness"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animate Tile"
      Height          =   4815
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CmdAniTile 
         Caption         =   "Add Animated Tile"
         Height          =   495
         Left            =   720
         TabIndex        =   37
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox ypos 
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Text            =   "0"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox xpos 
         Height          =   285
         Left            =   480
         TabIndex        =   33
         Text            =   "0"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Tile1y 
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Tile1x 
         Height          =   285
         Left            =   480
         TabIndex        =   30
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Tile2y 
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Text            =   "0"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Tile2x 
         Height          =   285
         Left            =   480
         TabIndex        =   28
         Text            =   "0"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Tile 2 x,y cords on main tile set"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Map Cords"
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Tile 1 x,y cords on main tile set"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   3360
      Top             =   3120
   End
   Begin VB.CommandButton CmdItem 
      Caption         =   "Item"
      Height          =   375
      Left            =   2400
      TabIndex        =   20
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton CmdMonster 
      Caption         =   "Monster"
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton CmdNpc 
      Caption         =   "Npc"
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox y 
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox X 
      Height          =   285
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton CmdCancel2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Frame Item 
      Caption         =   "Items"
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   3615
      Begin VB.FileListBox File3 
         Height          =   1065
         Left            =   1440
         Pattern         =   "*.item*"
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         Height          =   855
         Left            =   360
         ScaleHeight     =   795
         ScaleWidth      =   915
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CmdAddItem 
         Caption         =   "Add Item To Map"
         Height          =   615
         Left            =   840
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "ItemName"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame frameMonster 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3615
      Begin VB.FileListBox File1 
         Height          =   1065
         Left            =   1920
         Pattern         =   "*.bad*"
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   360
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CmdAddMonster 
         Caption         =   "Add Monster To Map"
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   855
         Left            =   600
         ScaleHeight     =   795
         ScaleWidth      =   915
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Monster Name"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label txtname 
         Caption         =   "Name"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame NPc 
      Height          =   2895
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   3615
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "Npc Text Here"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.FileListBox File2 
         Height          =   1065
         Left            =   1680
         Pattern         =   "*.npc*"
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton CmdAddNpc 
         Caption         =   "Add Npc To Map"
         Height          =   615
         Left            =   960
         TabIndex        =   16
         Top             =   2160
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         Height          =   855
         Left            =   360
         ScaleHeight     =   795
         ScaleWidth      =   915
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "NpcText"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Npc Name"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      Caption         =   $"AddMonster.frx":038A
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "X (1-85)"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Y (1-61)"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "AddMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W As Integer
Dim v As Integer
Dim vv As Integer
Dim Ani As Integer
'what I would do this module is to fix open and writing
'monsters, npcs, and items
'I would also fix variable names and fix the bug
'which causes you to keep clicking until the monster,npc or item appears on the screen


Private Sub CmdAddItem_Click()
If File3.Filename = "" Then
MsgBox "Please select a item!"
Else
vv = vv + 1
    Call ItemInfo.Createitem(X.Text, y.Text, vv)
End If
End Sub

Private Sub CmdAddMonster_Click()

Dim I As Integer
If File1.Filename = "" Then
MsgBox "Please select a monster!"
Else
I = 1
For I = I + 1 To 250 ' Iterate through each element.
 
If Monster(I).PlayerName = "" Then

    Call Stuff.createMonster(X.Text, y.Text, I)
    Call Stuff.MovenewMonster2


        Exit Sub
        I = 250
    End If
Next
End If

End Sub

Private Sub CmdAddNpc_Click()
If File2.Filename = "" Then
MsgBox "Please select a npc"
Else
v = v + 1
    Call Stuff.createNpc(X.Text, y.Text, v)
End If
End Sub

Private Sub CmdAniTile_Click()
F = FreeFile
Ani = Ani + 1
 Success% = FileExists%(App.Path & "\scripts\" & MainPlayer.MapName & Ani & ".ani")
 
If Success% = True Then
  MsgBox "This file exists.Please the button until it creates it!", vbOKOnly, "File Exists"
Else

Open (App.Path & "\scripts\" & MainPlayer.MapName & Ani & ".ani") For Binary Access Write Lock Write As F
    Put #F, , Tile1x.Text & vbCrLf  'cordx1
    Put #F, , Tile1y.Text & vbCrLf  'cordy1
    Put #F, , Tile2x.Text & vbCrLf  'cordx2
    Put #F, , Tile2y.Text & vbCrLf   'cordy2
    Put #F, , xpos.Text & vbCrLf
    Put #F, , ypos.Text & vbCrLf
Close F
MsgBox "Animation File created!", vbInformation
 End If
End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdCancel2_Click()
Unload Me
End Sub

Private Sub CmdItem_Click()
NPc.Visible = False
frameMonster.Visible = False
Item.Visible = True
End Sub

Private Sub CmdMonster_Click()
NPc.Visible = False
frameMonster.Visible = True
Item.Visible = False
End Sub

Private Sub CmdNpc_Click()
NPc.Visible = True
frameMonster.Visible = False
Item.Visible = False
End Sub

Private Sub CmdOk_Click()

Dim Choice As Integer
Dim Choice2 As Integer

If txtBright.Text < -255 Or txtBright.Text > 255 Then
    MsgBox "Brightness can only be between -255 and 255", vbExclamation
    Exit Sub
End If


MapExtra.Brightness = txtBright.Text
MapExtra.Music = txtmusic.Text
MapExtra.Weather = Combo1.Text

MapExtra.Notes = txtNotes.Text
MapExtra.MapName = txtMapTiltle.Text
MapExtra.MapTiles = txtTiles.Text

MapExtra.HScroll = txtHscroll.Text
MapExtra.VScroll = txtVscroll.Text
MapExtra.xpos = TxTLeft.Text
MapExtra.ypos = TxTtop.Text

Unload Me
Exit Sub

'used to store more data more made map files to large!!
'special stuff doesn't work
SpecialMap(0, 0).Brightness = txtBright.Text
SpecialMap(0, 0).Weather = Combo1.Text
SpecialMap(0, 0).Music = txtmusic.Text
If Combo2.Text = "Warp" Then
    Choice = 1
End If
If Combo2.Text = "Vender" Then
    Choice = 3
End If
If Combo2.Text = "Mine" Then
    Choice = 2
End If
SpecialMap(TileX, Tiley).Choice = Choice
SpecialMap(TileX, Tiley).HScroll = txtHscroll.Text
SpecialMap(TileX, Tiley).VScroll = txtVscroll.Text
SpecialMap(TileX, Tiley).PlayerLeft = TxTLeft.Text
SpecialMap(TileX, Tiley).PlayerTop = TxTtop.Text
SpecialMap(TileX, Tiley).MapName = TxtWarpMap.Text
If Combo2.Text = "Gold" Then
Choice2 = 1
End If
If Combo2.Text = "Silver" Then
Choice2 = 2
End If
If Combo2.Text = "Iron" Then
Choice2 = 3
End If

SpecialMap(TileX, Tiley).MineType = Choice2
SpecialMap(TileX, Tiley).MineAmount = Txtmine.Text
SpecialMap(TileX, Tiley).SoundEffect = txtSound.Text
SpecialMap(TileX, Tiley).VenderItem = Txtvender.Text

Unload Me
End Sub




Private Sub File1_Click()
If File1.Filename = "" Then Exit Sub

  Open App.Path & "\monsters\" & File1.Filename For Input As #1
    Line Input #1, textline1 'level
    Line Input #1, textline2 'admin
    Line Input #1, textline3 'xp
    Line Input #1, textline4 'xp2nextlevel
    Line Input #1, textline5 'mapname
    Line Input #1, textline6 'xpos
    Line Input #1, textline7 'ypos
    Line Input #1, textline8 'vscroll
    Line Input #1, textline9 'hscroll
    Line Input #1, textline10 'graphic
    Line Input #1, textline11 'money
    Line Input #1, textline12 'life
    Line Input #1, textline13 'stamina
    Line Input #1, textline14 'mana
    Line Input #1, textline15 'maxlife
    Line Input #1, textline16 'max stamina
    Line Input #1, textline17 'max mana
    Line Input #1, textline18 'strength
    Line Input #1, textline19 'quickness
    Line Input #1, textline20 'dexiterity
    Close #1
    
    
    Picture1.Picture = LoadPicture(App.Path & "\monsters\" & textline1 & ".gif")


End Sub

Private Sub File2_Click()
If File2.Filename = "" Then Exit Sub

  Open App.Path & "\monsters\" & File2.Filename For Input As #1
    Line Input #1, textline1 'level
    Line Input #1, textline2 'admin
    Line Input #1, textline3 'xp
    Line Input #1, textline4 'xp2nextlevel
    Line Input #1, textline5 'mapname
    Line Input #1, textline6 'xpos
   
    Close #1
    
    
    Picture3.Picture = LoadPicture(App.Path & "\monsters\" & textline1 & ".gif")


End Sub

Private Sub File3_Click()
If File3.Filename = "" Then Exit Sub

  Open App.Path & "\monsters\" & File3.Filename For Input As #1
    Line Input #1, textline1 'level
    Line Input #1, textline2 'admin
    Line Input #1, textline3 'xp
    Line Input #1, textline4 'xp2nextlevel
    Line Input #1, textline5 'mapname
    Line Input #1, textline6 'xpos
   
    Close #1
    
    
    Picture2.Picture = LoadPicture(App.Path & "\monsters\" & textline1 & ".gif")

End Sub

Private Sub Form_Load()
'txtBright.Text = SpecialMap(0, 0).Brightness
'Combo1.Text = SpecialMap(0, 0).Weather
'txtmusic.Text = SpecialMap(0, 0).Music

'Load extra data into textboxes
txtBright.Text = MapExtra.Brightness
txtmusic.Text = MapExtra.Music
Combo1.Text = MapExtra.Weather

txtNotes.Text = MapExtra.Notes
txtMapTiltle.Text = MapExtra.MapName
txtTiles.Text = MapExtra.MapTiles

txtHscroll.Text = MapExtra.HScroll
txtVscroll.Text = MapExtra.VScroll
TxTLeft.Text = MapExtra.xpos
TxTtop.Text = MapExtra.ypos


Combo1.Locked = False
Combo2.Locked = False
RPG.ForceTextBoxNumeric X, True
RPG.ForceTextBoxNumeric y, True
RPG.ForceTextBoxNumeric xpos, True
RPG.ForceTextBoxNumeric ypos, True
RPG.ForceTextBoxNumeric Tile1x, True
RPG.ForceTextBoxNumeric Tile1y, True
RPG.ForceTextBoxNumeric Tile2x, True
RPG.ForceTextBoxNumeric Tile2y, True
RPG.ForceTextBoxNumeric TileX, True
RPG.ForceTextBoxNumeric Tiley, True
RPG.ForceTextBoxNumeric Txtmine, True
RPG.ForceTextBoxNumeric TxTtop, True
RPG.ForceTextBoxNumeric TxTLeft, True
RPG.ForceTextBoxNumeric txtHscroll, True
RPG.ForceTextBoxNumeric txtVscroll, True
RPG.ForceTextBoxNumeric txtBright, True
File1.Path = App.Path & "\monsters\"
File2.Path = App.Path & "\monsters\"
File3.Path = App.Path & "\monsters\"
Ani = 0
End Sub

Private Sub Timer1_Timer()
Stuff.MovenewMonster2
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


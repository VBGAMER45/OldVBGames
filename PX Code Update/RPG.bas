Attribute VB_Name = "RPG"

'Tile constants
Public Const CustomStore = 22 'Used to load a custom store script!
Public Const TrapTILE = 21
Public Const LoadWarp = 20 'Sets a tile as a warp point, use the warp editor to make warps and then edit the map using the warp tile
Public Const Destroyable = 19
Public Const ExitTile = 18
Public Const HorseRanch = 17
Public Const Seeds = 16
Public Const TradeinStore = 15 'Place were you sell your items
Public Const SupplyStore = 14
Public Const FoodStore = 13
Public Const Gold = 12
Public Const Tree = 11 'allows you cut down trees
Public Const WarpPortal = 10 'Shows the map to places to warp to
Public Const ArmorStore = 9
Public Const WeaponStore = 8
Public Const MagicStore = 7
Public Const DoorTile = 6     'door tile
Public Const HurtTile = 5     'you loose one health if you go on this tile
Public Const DieTile = 4    'you die if you go in this tile
Public Const Sign = 3        'sign tile
Public Const SWIMABLE = 2    'swimable tile
Public Const WALKABLE = 1   'this will be a walkable tile
Public Const NONWALKABLE = 0 'this will be a non-walkable tile

Global Warping As Boolean 'Basicly tells the game to redraw the map or not

Global ZeldaMode As Boolean 'if the game is fullscreen or not

Global Treex As Integer
Global Treey As Integer

Global OnlineGame As Boolean
Dim QuestNumber As Integer
Dim NumQuest As Integer
Global AnimatedTiles As Boolean
Global Targetx As Integer
Global Targety As Integer
Global Z As Integer
Public Type typeTileStructure
    lngTileType As Long 'The type of tile the tile is (walkable/non-walkable)
    lngtilekind As Long
    lngTileImgNumber As Long 'stores which tile picture number this tile is associated with
End Type

'API Call to get Cursor

Public Type PointAPI
    X As Long
    y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Global Pnt As PointAPI

'API to Blit a picture to the screen
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

'Map Stuff
Global HScroll As Integer, VScroll As Integer ' Tells the Tools Form the map editors scrollbar positions
Global mWidth As Integer, mHeight As Integer
Global CurX As Single, CurY As Single

Public Type MapData
    TileX As Integer
    Tiley As Integer
    Tiletype As Integer
    lngtilekind As String
End Type
'This type holds extra data for the map
'I was having problems with this and took it out but
Private Type SpecialData
    Choice As Integer 'choices 1=warp 2=mine 3=vender
    VenderItem As String
    HScroll As Integer
    VScroll As Integer
    MapName As String
    Weather As String
    Brightness As Integer
    Locked As Integer
    KeyNumber As Integer
    Mapx As Integer
    Mapy As Integer
    MapTiles As String
    Music As String
    SoundEffectPlayed As Boolean
    SoundEffect As String
    PlayerTop As Integer
    PlayerLeft As Integer
    MineType As Integer ' 1=gold 2=iron 3=silver
    MineAmount As Integer
    Amountleft As Integer
    Seeds As Integer
End Type


Public SpecialMap() As SpecialData
Public Map() As MapData 'Stores map info
Global MapLayer As Integer 'What layer are we on?
'Global WMove As String
Global TempTileX As Integer, TempTileY As Integer 'Which tile is currently selected
Global SelectedTool As String 'Which tool are we using?
Global RefreshMini As Boolean 'Refresh minimap when needed

Private Type ExtraData
    HScroll As Integer
    VScroll As Integer
    MapName As String
    Weather As String
    Brightness As Integer
    Mapx As Integer
    Mapy As Integer
    MapTiles As String
    Music As String
    ypos As Integer
    xpos As Integer
    Notes As String
End Type
Public MapExtra As ExtraData

'Monster Type
Public Type MonsterType
    MonsterIsQuest As Integer
    DoesMonsterAttack As Boolean
    MonsterAttack As Boolean
    MonsterCasts As Boolean
    MonsterHeals As Boolean
    MonsterSpeakText As String
    MonsterAttack2 As Integer
    ImageName As String 'Image control number
    InUse As Boolean
    PlayerVisible As Boolean
    Cheater As Boolean
    Level As Integer
    PlayerName As String
    Health As Integer
    MaxHealth As Integer
    Wisdom As Integer
    Intelligence As Integer
    Strength As Integer
    Dexterity As Integer
    Stamina As Integer
    MaxStamina As Integer
    Mana As Integer
    MaxMana As Integer
    Quickness As Integer
    CastSpell As Integer
    Graphic As Integer
    xpos As Integer
    ypos As Integer
    VScroll As Integer
    HScroll As Integer
    Stealth As Integer
    MapName As String
    Scan As Integer
    OrginalLocation As Boolean
    OrgX As Integer
    OrgY As Integer
    CurX As Integer
    CurY As Integer
    Armor As Integer
    Weapon As String
    Damage As Integer
    Speed As Integer
    MonsterMove As Boolean
    MeleeDefense As Integer
    MagicDefense As Integer
End Type
Public Type Npctype
ImageName As String
CurX As Integer
CurY As Integer
MapName As String
xpos As Integer
ypos As Integer
PlayerName As String
NpcItem As String
NpcText As String
NpcisFlag As Boolean
End Type

Const MAX_Monster = 250
Global Monster(MAX_Monster) As MonsterType
Global NPc(MAX_Monster) As Npctype

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Sleep
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'APIs:
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long) As Long

'##npc speech stuff
Public NpcMap As New clsNPC
Public NPCcount As Integer
Public NPCTalk(5) As New clsNPC
Public Type SpeechType
    Pitanje As String
    Odgovor(10) As String
    Redirect(10) As Integer
    nOdgovor As Integer
End Type
Public Speech(30) As SpeechType
Public TalkingTo As Integer, CSpeech As Integer
''end npc speech stuff
'Variables (these are needed for GetRGBs):
Dim Red As Long
Dim Green As Long
Dim Blue As Long

'
Dim TalkOnlyOnce As Integer

#If Win16 Then
    Type RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type
#Else
    Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
#End If

'User and GDI Functions for Explode/Implode to work

#If Win16 Then
    Declare Sub GetWindowRect Lib "User" (ByVal hwnd As Integer, lpRect As RECT)
    Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
    Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
    Declare Sub SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
    Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
    Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Declare Function SelectObject Lib "user32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

'****************************************************************
'*Author: Carl Slutter
'*
'*Description:
'*The higher the "Movement", the slower the window
'*"explosion".
'*
'*Creation Date: Thursday  23 January 1997  2:27 pm
'*Revision Date: Thursday  23 January 1997  2:27 pm
'*
'*Version Number: 1.00
'****************************************************************

Sub ExplodeForm(F As Form, Movement As Integer)
    Dim myRect As RECT
    Dim formWidth%, formHeight%, I%, X%, y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect F.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(F.BackColor)
    
    For I = 1 To Movement
        cx = formWidth * (I / Movement)
        cy = formHeight * (I / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, y, X + cx, y + cy
    Next I
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
    
End Sub


Public Sub ImplodeForm(F As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'****************************************************************
'*Author: Carl Slutter
'*
'*Description:
'*The larger the "Movement" value, the slower the "Implosion"
'*
'*Creation Date: Thursday  23 January 1997  2:42 pm
'*Revision Date: Thursday  23 January 1997  2:42 pm
'*
'*Version Number: 1.00
'****************************************************************
    
    Dim myRect As RECT
    Dim formWidth%, formHeight%, I%, X%, y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect F.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(F.BackColor)
    
        For I = Movement To 1 Step -1
        cx = formWidth * (I / Movement)
        cy = formHeight * (I / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, y, X + cx, y + cy
    Next I
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub



Public Sub ForceTextBoxNumeric(TextBox As TextBox, Optional Force As Boolean = True)
    Dim style As Long
    Const GWL_STYLE = (-16)
    Const ES_NUMBER = &H2000
    
    ' get current style
    style = GetWindowLong(TextBox.hwnd, GWL_STYLE)
    If Force Then
        style = style Or ES_NUMBER
    Else
        style = style And Not ES_NUMBER
    End If
    ' enforce new style
    SetWindowLong TextBox.hwnd, GWL_STYLE, style
End Sub


Public Function Snap(Cordinate As Variant, Dimension As Integer) As Integer
Snap = (Cordinate \ Dimension) * Dimension 'Small algorithm that snaps to grid sort of
End Function

Public Function Snap2(Cordinate As Variant, Dimension As Integer) As Integer
Snap2 = (Cordinate * Dimension) 'Small algorithm that takes a number and converts it into a pixel cord.
End Function

Public Sub SaveCharecter()
If MainPlayer.PlayerName = "" Then
Else
Dim outputfile As Integer
outputfile = 1
Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".char" For Output As #outputfile
Write #outputfile, MainPlayer.Level, MainPlayer.Admin, MainPlayer.XP, MainPlayer.Xp4nextLevel, MainPlayer.MapName, MainPlayer.xpos, MainPlayer.ypos
Write #outputfile, MainPlayer.VScroll, MainPlayer.HScroll, MainPlayer.Graphic, MainPlayer.Money, MainPlayer.Health, MainPlayer.Stamina, MainPlayer.Mana
Write #outputfile, MainPlayer.MaxHealth, MainPlayer.MaxStamina, MainPlayer.MaxMana, MainPlayer.Strength, MainPlayer.Quickness, MainPlayer.Dexterity, MainPlayer.Wisdom
Write #outputfile, MainPlayer.Intelligence, MainPlayer.Burden, MainPlayer.KILLS, MainPlayer.Deaths, MainPlayer.SkillPoints, MainPlayer.SkillsPercent, MainPlayer.SpendXp, MainPlayer.BankAccountMoney
Write #outputfile, MainPlayer.AttributePoints, PlayerSkills.Axe, PlayerSkills.Blacksmith, PlayerSkills.Bombs, PlayerSkills.Bow, PlayerSkills.Builder, PlayerSkills.CastSpell, PlayerSkills.Dagger, PlayerSkills.Fail
Write #outputfile, PlayerSkills.Farmer, PlayerSkills.Hunter, PlayerSkills.Lumberjack, PlayerSkills.Mace, PlayerSkills.MagicDefense, PlayerSkills.MeleeDefense, PlayerSkills.Miner, PlayerSkills.Run, PlayerSkills.Scan, PlayerSkills.Spear
Write #outputfile, PlayerSkills.Staff, PlayerSkills.Stealth, PlayerSkills.Swim, PlayerSkills.Sword, PlayerSkills.TheForce, PlayerSkills.UnarmedAttack
Close #outputfile

'Here is a more secure way to store data just
'uncommented and comment the above write staments
''Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".char" For Binary Access Write Lock Write As #2
''Put #2, , MainPlayer
''Close #2
''Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".char" For Binary Access Write Lock Write As #3
''Put #3, , PlayerSkills
''Close #3
Call ItemInfo.SavePlayerItems
Debug.Print "###Player Stats Saved###"
End If
End Sub

Sub WriteVar(file As String, Main As String, Var As String, Value As String)
On Error Resume Next

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

WritePrivateProfileString Main, Var, Value, file
    
End Sub
Sub WriteINT(file As String, Main As String, Var As String, Value As Variant)
Dim Value2 As String
On Error Resume Next
Value2 = Value
'*****************************************************************
'Writes intenger to a text file
'*****************************************************************

WritePrivateProfileString Main, Var, Value2, file
    
End Sub
Public Sub CheckTile()

If Map(MainPlayer.xpos, MainPlayer.ypos = MainPlayer.ypos + 1).Tiletype = 0 Then
Form1.NowALKDIRECTION.Caption = "Up"
End If
If Map(MainPlayer.xpos, MainPlayer.ypos = MainPlayer.ypos - 1).Tiletype = 0 Then
Form1.NowALKDIRECTION.Caption = "Down"
End If
If Map(Linkx = Linkx - 1, linky).Tiletype = 0 Then
Form1.NowALKDIRECTION.Caption = "Left"
End If
If Map(Linkx = Linkx + 1, linky).Tiletype = 0 Then
Form1.NowALKDIRECTION.Caption = "Right"
End If
End Sub


Public Sub CutDownTree()
Dim X As Single
Dim y As Single
Dim wood As Integer
I = 0
For I = I + 1 To MaxItems
    If PXItem(I).ItemName = "wood" Then
        wood = I
        I = MaxItems
    End If
Next I

'NOte 10368 is temptilex
'note 0 is temptiley
'change numbers to dead tree pos

If Map(Treex, Treey).Tiletype = Tree Or Map(Treex, Treey).TileX = 352 And Map(Treex, Treey).Tiley = 0 Then
BitBlt Form1.hDC, Snap(X, 32), Snap(y, 32), 32, 32, Form1.Tiles.hDC, 8032, 64, SRCCOPY
    
    'Save map info to the map array
    Map(Treex, Treey).TileX = 8032
    Map(Treex, Treey).Tiley = 64
    Map(Treex, Treey).Tiletype = NONWALKABLE
    Form1.RedrawMap
    PXItem(wood).Quanity = PXItem(wood).Quanity + 100
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You cut down tree for 100 wood" & vbNewLine
'make a save map sub
    MainPlayer.XP = MainPlayer.XP + 5
    MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 5
    TOOLS.LoadStats
    Form1.AddPlayerStats "Xp", 5
Else
    Form1.txtDialog.Text = Form1.txtDialog.Text & "Unable to cut down tree" & vbNewLine
Exit Sub
End If
End Sub
Public Sub PlantSeeds()
Dim X As Single
Dim y As Single
Dim Seeds2 As Integer

I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "Seeds" Then
Seeds2 = I
I = MaxItems
End If
Next I
'NOte 10368 is temptilex
'note 0 is temptiley
If Map(Treex, Treey).Tiletype = 1 Then
PXItem(Seeds2).Quanity = PXItem(Seeds2).Quanity - 1
BitBlt Form1.hDC, Snap(X, 32), Snap(y, 32), 32, 32, Form1.Tiles.hDC, 608, 0, SRCCOPY
    
    'Save map info to the map array
    Map(Treex, Treey).TileX = 608
    Map(Treex, Treey).Tiley = 0
    Map(Treex, Treey).Tiletype = Seeds
    Form1.RedrawMap
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You planted seeds" & vbNewLine
'make a save map sub
    MainPlayer.XP = MainPlayer.XP + 1
    MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 1
    TOOLS.LoadStats
    Form1.AddPlayerStats "Xp", 1
Else
    Form1.txtDialog.Text = Form1.txtDialog.Text & "Unable to plant seeds" & vbNewLine
Exit Sub
End If
End Sub
Public Sub HarvestSeeds(HarvestX As Integer, HarvestY As Integer)
Dim X As Single
Dim y As Single
Dim Apple As Integer, Strawberry As Integer, Watermelon As Integer, Carrot As Integer
Dim seedType As Integer
Randomize
seedType = (Rnd * 3)
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "Apple" Then
Apple = I
End If
If PXItem(I).ItemName = "Carrot" Then
Carrot = I
End If
If PXItem(I).ItemName = "Watermelon" Then
Watermelon = I
End If
If PXItem(I).ItemName = "Strawberry" Then
Strawberry = I
End If
Next I
'NOte 10368 is temptilex
'note 0 is temptiley

If Map(HarvestX, HarvestY).Tiletype = 16 Then
If seedType = 0 Then
PXItem(Apple).Quanity = PXItem(Apple).Quanity + 1
End If
If seedType = 1 Then
PXItem(Strawberry).Quanity = PXItem(Strawberry).Quanity + 1
End If
If seedType = 2 Then
PXItem(Watermelon).Quanity = PXItem(Watermelon).Quanity + 1

End If
If seedType = 3 Then
PXItem(Carrot).Quanity = PXItem(Carrot).Quanity + 1
End If

BitBlt Form1.hDC, Snap(X, 32), Snap(y, 32), 32, 32, Form1.Tiles.hDC, 640, 0, SRCCOPY
  
    'Save map info to the map array
    Map(HarvestX, HarvestY).TileX = 640
    Map(HarvestX, HarvestY).Tiley = 0

    Map(HarvestX, Treey).Tiletype = WALKABLE
    Form1.RedrawMap
 
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You harvest seeds" & vbNewLine
'make a save map sub FOR online play

    MainPlayer.XP = MainPlayer.XP + 2
    MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 2
    TOOLS.LoadStats
    Form1.AddPlayerStats "Xp", 2
Else
    Form1.txtDialog.Text = Form1.txtDialog.Text & "Unable to harvest seeds" & vbNewLine
    Exit Sub
End If
End Sub
Public Sub RandomMap()
'Loads a random map from 1.1.map to 10.10.map
Dim num1 As Integer
Dim Num2 As Integer
Randomize
num1 = (Rnd * 10)
Num2 = (Rnd * 10)

If num1 = 0 Then
num1 = 1

End If
If Num2 = 0 Then
Num2 = 1
End If
If num1 = 5 And Num2 = 5 Then
num1 = 5
Num2 = 6
End If
Open App.Path & "\maps\" & num1 & "." & Num2 & ".map" For Input As #1: Close
Open App.Path & "\maps\" & num1 & "." & Num2 & ".map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.RedrawMap
Form1.HScroll1.Value = "45"
Form1.VScroll1.Value = "23"

Form1.MapName.Caption = num1 & "." & Num2 & ".map"
MainPlayer.MapName = num1 & "." & Num2 & ".map"
End Sub
Public Sub SaveMonsterLocation()
Dim K As Integer
K = 0
For K = K + 1 To 250
    newmon = "MonsterImage" & K
    On Error GoTo nofile:
    F = FreeFile
    If Monster(K).PlayerName = "" Then
    Exit Sub
    Exit For
    End If
Open (App.Path & "\monsters\monster" & K & ".xy") For Binary Access Write Lock Write As F
    Put #F, , Monster(K).xpos & vbNewLine
    Put #F, , Monster(K).ypos & vbNewLine
    Put #F, , Monster(K).MapName & vbNewLine
    Put #F, , Monster(K).PlayerName & vbNewLine

Close F

If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
Debug.Print "###Monsters Saved###"
Exit Sub
nofile:
Debug.Print "!!!Monsters Not Saved!!!"
Exit Sub
End Sub

Public Sub Opendoors()
'open doors
'If Map(MainPlayer.xpos, MainPlayer.ypos - 1).TileType = 1 Then

If Map(MainPlayer.xpos, MainPlayer.ypos - 1).TileX = 11072 And Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiley = 64 Then
Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiletype = 0
Map(MainPlayer.xpos, MainPlayer.ypos - 1).TileX = 11104
Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiley = 64
Map(MainPlayer.xpos, MainPlayer.ypos - 2).Tiletype = 0
Map(MainPlayer.xpos, MainPlayer.ypos - 2).TileX = 11104
Map(MainPlayer.xpos, MainPlayer.ypos - 2).Tiley = 32
Form1.RedrawMap
Exit Sub
End If
'End If

'close doors
'If Map(MainPlayer.xpos, MainPlayer.ypos - 1).TileType = 0 Then
If Map(MainPlayer.xpos, MainPlayer.ypos - 1).TileX = 11104 And Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiley = 64 Then
Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiletype = 1
Map(MainPlayer.xpos, MainPlayer.ypos - 1).TileX = 11072
Map(MainPlayer.xpos, MainPlayer.ypos - 1).Tiley = 64
Map(MainPlayer.xpos, MainPlayer.ypos - 2).Tiletype = 1
Map(MainPlayer.xpos, MainPlayer.ypos - 2).TileX = 11072
Map(MainPlayer.xpos, MainPlayer.ypos - 2).Tiley = 32
Form1.RedrawMap
Exit Sub
End If
'End If


End Sub
Public Sub SaveNpcLocation()
Dim K As Integer
K = 0
For K = K + 1 To 250
   ' newmon = "NPcImage" & k
    On Error GoTo nofile:
    F = FreeFile
    If NPc(K).PlayerName = "" Then
        Exit Sub
        Exit For
    End If
    Open (App.Path & "\monsters\NPc" & K & ".xy") For Binary Access Write Lock Write As F
        Put #F, , NPc(K).xpos & vbNewLine
        Put #F, , NPc(K).ypos & vbNewLine
        Put #F, , NPc(K).MapName & vbNewLine
        Put #F, , NPc(K).PlayerName & vbNewLine
        Put #F, , NPc(K).NpcText & vbNewLine
    Close F

If NPc(K).PlayerName = "" Then
       Exit For
       End If
Next
Debug.Print "###Npc's Saved###"
Exit Sub
nofile:
Debug.Print "!!!Npc NOT Saved!!!"
Exit Sub
End Sub
Public Sub LoadMonsters()
Dim K As Integer

K = 0

For K = K + 1 To 250
    Open App.Path & "\monsters\monster" & K & ".xy" For Input As #1
        Line Input #1, Line1 'x
        Line Input #1, LINE2 'y
        Line Input #1, Line3 'map
        Line Input #1, Line4 'name
    Close #1

    On Error GoTo nofile:
    Static I As Integer


   If Monster(K).PlayerName = "" Then
   
    Open App.Path & "\monsters\" & Line4 & ".bad" For Input As #1
        Line Input #1, textline1 'name
        Line Input #1, textline2 'damage
        Line Input #1, textline3 'speed
        Line Input #1, textline4 'monstermove
        Line Input #1, textline5 'weapon
        Line Input #1, textline6 'armor
        Line Input #1, textline7 'strength
        Line Input #1, textline8 'level
        Line Input #1, textline9 'health
        Line Input #1, textline10 'stamina
        Line Input #1, textline11 'inteligence
        Line Input #1, textline12 'mana
        Line Input #1, textline13 'scan
        Line Input #1, textline14 'magicattack
        Line Input #1, textline15 'stealh
        Line Input #1, textline16 'magicdefense
        Line Input #1, textline17 'meleedefense
        Line Input #1, textline18 'xpos
        Line Input #1, textline19 'ypos
        Line Input #1, textline20 'monster attack
    Close #1

    Z = Z + 1
    Monster(K).PlayerName = textline1
    Monster(K).Damage = textline2
    Monster(K).Speed = textline3
    Monster(K).MonsterMove = textline4
    Monster(K).Weapon = textline5
    Monster(K).Armor = textline6
    Monster(K).Strength = textline7
    Monster(K).Level = textline8
    Monster(K).MaxHealth = textline9
    Monster(K).Health = Monster(K).MaxHealth
    Monster(K).MaxStamina = textline10
    Monster(K).Intelligence = textline11
    Monster(K).MaxMana = textline12
    Monster(K).Scan = textline13
    Monster(K).CastSpell = textline14
    Monster(K).Stealth = textline15
    Monster(K).MeleeDefense = textline17
    Monster(K).MagicDefense = textline16
    Monster(K).Mana = Monster(K).MaxMana
    Monster(K).Stamina = Monster(K).MaxStamina
    Monster(K).Armor = textline6
    Monster(K).MonsterAttack2 = textline20
    
' Form1.p21 (i)
     Monster(K).xpos = Line1
      Monster(K).ypos = LINE2
   Set P51 = Form1.Controls.Add("VB.Image", "MonsterImage" & Z)

   With P51
      .Enabled = True
      .Top = Line1
      .Left = LINE2
      .Height = "10"
      .Width = "10"
      .Picture = LoadPicture(App.Path & "\monsters\" & Monster(K).PlayerName & ".gif")
      .Visible = True
      .ToolTipText = Monster(K).PlayerName & " Level " & Monster(K).Level & " Health " & Monster(K).Health & "\" & Monster(K).MaxHealth
     ' Form1.Caption = .Name & "mon"
   End With
     
'   Form1.p21(z).Top = y
'   Form1.p21(z).Left = X
   Monster(K).ImageName = Z
   Monster(K).MapName = Line3
'   Monster(k).Curx = (Snap(Form1.p21(z).Left, 32) \ 32) + HScroll
 '  Monster(k).CurY = (Snap(Form1.p21(z).Top, 32) \ 32) + VScroll

   ' Form1.p21(z).Visible = True
  '  Call MovenewMonster(k, z)
    
    End If
    If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
Debug.Print "###Monsters Loaded###"
Exit Sub
nofile:
Debug.Print "!!!Monsters NOT Loaded!!! " & Err.Description & " " & Err.Source
Exit Sub
End Sub
Public Sub LoadNpcs()
Dim K As Integer
Dim C As Integer
K = 0
On Error GoTo nofile:
For K = K + 1 To 250
    Open App.Path & "\monsters\npc" & K & ".xy" For Input As #1
        Line Input #1, Line1 'x
        Line Input #1, LINE2 'y
        Line Input #1, Line3 'map
        Line Input #1, Line4 'name
        Line Input #1, Line5 ' text
    Close #1



Static I As Integer


   If NPc(K).PlayerName = "" Then
   
    Open App.Path & "\monsters\" & Line4 & ".npc" For Input As #1
        Line Input #1, textline1 'name
        Line Input #1, textline2 'damage
        Line Input #1, textline3 'speed
        Line Input #1, textline4 'monstermove
        Line Input #1, textline5 'weapon
        Line Input #1, textline6 'armor
    Close #1

    zz = zz + 1
      NPc(K).PlayerName = textline1
      NPc(K).NpcText = Line5
      NPc(K).NpcItem = textline3
      NPc(K).xpos = Line1
      NPc(K).ypos = LINE2
  ' Set P51 = Form1.Controls.Add("VB.Image", "NpcImage" & zz)
C = Form1.PICnpc.UBound + 1
Load Form1.PICnpc(C)
 '  With P51
      Form1.PICnpc(C).Enabled = True
      Form1.PICnpc(C).Top = Line1
      Form1.PICnpc(C).Left = LINE2
      Form1.PICnpc(C).Height = 10
      Form1.PICnpc(C).Width = 10
      Form1.PICnpc(C).Picture = LoadPicture(App.Path & "\monsters\" & NPc(K).PlayerName & ".gif")
      Form1.PICnpc(C).Visible = True
      Form1.PICnpc(C).ToolTipText = NPc(K).PlayerName

'   End With
     

   NPc(K).ImageName = C 'zz
   NPc(K).MapName = Line3

    
    End If
    
       If NPc(K).PlayerName = "" Then
       Exit For
       End If
Next
Debug.Print "###Npcs Loaded###"
Exit Sub
nofile:
Debug.Print "!!!Npcs NOT Loaded!!! " & Err.Description & " " & Err.Source
Exit Sub
End Sub
Public Sub TalkToNpc(K As Integer) 'Talk to npc
'displays that cool box

TalkOnlyOnce = TalkOnlyOnce + 1
If TalkOnlyOnce = 1 Then
Exit Sub
Else
Form1.TextScroll.Visible = True
Form1.Holder.Visible = True

Form1.PicTXT.Visible = True
NpcMap.LoadMap App.Path & "\scripts\" & K & ".snpc"
'Form1.txtDialog.Text = Form1.txtDialog.Text & NPc(k).PlayerName & " " & NPc(k).NpcText & vbCrLf
InitSpeech (0)
TalkOnlyOnce = 0
End If


End Sub
Public Sub SkillStatLevel(Skillname As String)

If Skillname = "All" Then
If PlayerSkills.MeleeDefense = 0 Then
PlayerSkills.MeeleNextLevel = 100
Else
PlayerSkills.MeeleNextLevel = PlayerSkills.MeleeDefense * 100
End If
If PlayerSkills.MagicDefense = 0 Then
PlayerSkills.MagicDefenseNextLevel = 100
Else
PlayerSkills.MagicDefenseNextLevel = PlayerSkills.MagicDefense * 100
End If
If PlayerSkills.CastSpell = 0 Then
PlayerSkills.CastSpellNextLevel = 100
Else
PlayerSkills.CastSpellNextLevel = PlayerSkills.CastSpell * 100
End If
If PlayerSkills.Stealth = 0 Then
PlayerSkills.StealthNextLevel = 100
Else
PlayerSkills.StealthNextLevel = PlayerSkills.Stealth * 100
End If
If PlayerSkills.Sword = 0 Then
PlayerSkills.SwordNextLvl = 100
Else
PlayerSkills.SwordNextLvl = PlayerSkills.Sword * 100
End If
If PlayerSkills.Axe = 0 Then
PlayerSkills.AxeNextLevel = 100
Else
PlayerSkills.AxeNextLevel = PlayerSkills.Axe * 100
End If
If PlayerSkills.Bombs = 0 Then
PlayerSkills.BombsNextLevel = 100
Else
PlayerSkills.BombsNextLevel = PlayerSkills.Bombs * 100
End If
If PlayerSkills.Bow = 0 Then
PlayerSkills.BowNextLevel = 100
Else
PlayerSkills.BowNextLevel = PlayerSkills.Bow * 100
End If
If PlayerSkills.Dagger = 0 Then
PlayerSkills.DaggerNxtLvl = 100
Else
PlayerSkills.DaggerNxtLvl = PlayerSkills.Dagger * 100
End If
If PlayerSkills.Fail = 0 Then
PlayerSkills.FailNextLvl = 100
Else
PlayerSkills.FailNextLvl = PlayerSkills.Fail * 100
End If
If PlayerSkills.Mace = 0 Then
PlayerSkills.MaceNextLevel = 100
Else
PlayerSkills.MaceNextLevel = PlayerSkills.Mace * 100
End If
If PlayerSkills.Spear = 0 Then
PlayerSkills.SpearNextLvl = 100
Else
PlayerSkills.SpearNextLvl = PlayerSkills.Spear * 100
End If
If PlayerSkills.Staff = 0 Then
PlayerSkills.StaffNxtLvl = 100
Else
PlayerSkills.StaffNxtLvl = PlayerSkills.Staff * 100
End If
If PlayerSkills.TheForce = 0 Then
PlayerSkills.theforceNextLvl = 100
Else
PlayerSkills.theforceNextLvl = PlayerSkills.TheForce * 100
End If
If PlayerSkills.UnarmedAttack = 0 Then
PlayerSkills.UnarmedAttackNextLevel = 100
Else
PlayerSkills.UnarmedAttackNextLevel = PlayerSkills.UnarmedAttack * 100
End If
End If
If Skillname = "melee" Then
If PlayerSkills.MeleeDefense = 0 Then
PlayerSkills.MeeleNextLevel = 100
Else
PlayerSkills.MeeleNextLevel = PlayerSkills.MeleeDefense * 100
'MsgBox "melee" & PlayerSkills.MeleeDefense & " nextlevel " & PlayerSkills.MeeleNextLevel
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your Melee Defense is now " & PlayerSkills.MeleeDefense & ". " & PlayerSkills.MeeleNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "magic" Then
If PlayerSkills.MagicDefense = 0 Then
PlayerSkills.MagicDefenseNextLevel = 100
Else
PlayerSkills.MagicDefenseNextLevel = PlayerSkills.MagicDefense * 100
'MsgBox "magic" & PlayerSkills.MagicDefense & " nextlevel " & PlayerSkills.MagicDefenseNextLevel
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your Magic Defense is now " & PlayerSkills.MagicDefense & ". " & PlayerSkills.MagicDefenseNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "castspell" Then
If PlayerSkills.CastSpell = 0 Then
PlayerSkills.CastSpellNextLevel = 100
Else
PlayerSkills.CastSpellNextLevel = PlayerSkills.CastSpell * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your castspell is now " & PlayerSkills.CastSpell & ". " & PlayerSkills.CastSpellNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "stealth" Then
If PlayerSkills.Stealth = 0 Then
PlayerSkills.StealthNextLevel = 100
Else
PlayerSkills.StealthNextLevel = PlayerSkills.Stealth * 100
'MsgBox "stealth" & PlayerSkills.Stealth & " nextlevel " & PlayerSkills.StealthNextLevel
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your stealth is now " & PlayerSkills.Stealth & ". " & PlayerSkills.StealthNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "sword" Then
If PlayerSkills.Sword = 0 Then
PlayerSkills.SwordNextLvl = 100
Else
PlayerSkills.SwordNextLvl = PlayerSkills.Sword * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your sword is now " & PlayerSkills.Sword & ". " & PlayerSkills.SwordNextLvl & " till next rasie." & vbNewLine
End If
End If
If Skillname = "axe" Then
If PlayerSkills.Axe = 0 Then
PlayerSkills.AxeNextLevel = 100
Else
PlayerSkills.AxeNextLevel = PlayerSkills.Axe * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your axe is now " & PlayerSkills.Axe & ". " & PlayerSkills.AxeNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "bombs" Then
If PlayerSkills.Bombs = 0 Then
PlayerSkills.BombsNextLevel = 100
Else
PlayerSkills.BombsNextLevel = PlayerSkills.Bombs * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your Bombs Skill is now " & PlayerSkills.Bombs & ". " & PlayerSkills.BombsNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "bow" Then
If PlayerSkills.Bow = 0 Then
PlayerSkills.BowNextLevel = 100
Else
PlayerSkills.BowNextLevel = PlayerSkills.Bow * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your bow is now " & PlayerSkills.Bow & ". " & PlayerSkills.BowNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "dagger" Then
If PlayerSkills.Dagger = 0 Then
PlayerSkills.DaggerNxtLvl = 100
Else
PlayerSkills.DaggerNxtLvl = PlayerSkills.Dagger * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your dagger is now " & PlayerSkills.Dagger & ". " & PlayerSkills.DaggerNxtLvl & " till next rasie." & vbNewLine
End If
End If
If Skillname = "fail" Then
If PlayerSkills.Fail = 0 Then
PlayerSkills.FailNextLvl = 100
Else
PlayerSkills.FailNextLvl = PlayerSkills.Fail * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your fail is now " & PlayerSkills.Fail & ". " & PlayerSkills.FailNextLvl & " till next rasie." & vbNewLine
End If
End If
If Skillname = "mace" Then
If PlayerSkills.Mace = 0 Then
PlayerSkills.MaceNextLevel = 100
Else
PlayerSkills.MaceNextLevel = PlayerSkills.Mace * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your mace is now " & PlayerSkills.Mace & ". " & PlayerSkills.MaceNextLevel & " till next rasie." & vbNewLine
End If
End If
If Skillname = "spear" Then
If PlayerSkills.Spear = 0 Then
PlayerSkills.SpearNextLvl = 100
Else
PlayerSkills.SpearNextLvl = PlayerSkills.Spear * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your spear is now " & PlayerSkills.Spear & ". " & PlayerSkills.SpearNextLvl & " till next rasie." & vbNewLine
End If
End If
If Skillname = "staff" Then
If PlayerSkills.Staff = 0 Then
PlayerSkills.StaffNxtLvl = 100
Else
PlayerSkills.StaffNxtLvl = PlayerSkills.Staff * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your staff is now " & PlayerSkills.Staff & ". " & PlayerSkills.StaffNxtLvl & " till next rasie." & vbNewLine
End If
End If
If Skillname = "force" Then
If PlayerSkills.TheForce = 0 Then
PlayerSkills.theforceNextLvl = 100
Else
PlayerSkills.theforceNextLvl = PlayerSkills.TheForce * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your The Force is now " & PlayerSkills.TheForce & ". " & PlayerSkills.theforceNextLvl & " till next rasie." & vbNewLine
End If
End If
If Skillname = "unarmedattack" Then
If PlayerSkills.UnarmedAttack = 0 Then
PlayerSkills.UnarmedAttackNextLevel = 100
Else
PlayerSkills.UnarmedAttackNextLevel = PlayerSkills.UnarmedAttack * 100
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your unarmed attack is now " & PlayerSkills.UnarmedAttack & ". " & PlayerSkills.UnarmedAttackNextLevel & " till next rasie." & vbNewLine
End If
End If
End Sub
Public Sub CheckIfStatLeveled()
If PlayerSkills.MeeleNextLevel <= 0 Then
PlayerSkills.MeleeDefense = PlayerSkills.MeleeDefense + 1
Call SkillStatLevel("melee")

End If
If PlayerSkills.MagicDefenseNextLevel <= 0 Then
PlayerSkills.MagicDefense = PlayerSkills.MagicDefense + 1
Call SkillStatLevel("magic")
End If
If PlayerSkills.CastSpellNextLevel <= 0 Then
PlayerSkills.CastSpell = PlayerSkills.CastSpell + 1
Call SkillStatLevel("castspell")
End If
If PlayerSkills.StealthNextLevel <= 0 Then
PlayerSkills.Stealth = PlayerSkills.Stealth + 1
Call SkillStatLevel("stealth")
End If
If PlayerSkills.AxeNextLevel <= 0 Then
PlayerSkills.Axe = PlayerSkills.Axe + 1
Call SkillStatLevel("axe")
End If
If PlayerSkills.BombsNextLevel <= 0 Then
PlayerSkills.Bombs = PlayerSkills.Bombs + 1
Call SkillStatLevel("bombs")
End If
If PlayerSkills.BowNextLevel <= 0 Then
PlayerSkills.Bow = PlayerSkills.Bow + 1
Call SkillStatLevel("bow")
End If
If PlayerSkills.DaggerNxtLvl <= 0 Then
PlayerSkills.Dagger = PlayerSkills.Dagger + 1
Call SkillStatLevel("dagger")
End If
If PlayerSkills.FailNextLvl <= 0 Then
PlayerSkills.Fail = PlayerSkills.Fail + 1
Call SkillStatLevel("fail")
End If
If PlayerSkills.MaceNextLevel <= 0 Then
PlayerSkills.Mace = PlayerSkills.Mace + 1
Call SkillStatLevel("mace")
End If
If PlayerSkills.SpearNextLvl <= 0 Then
PlayerSkills.SpearNextLvl = PlayerSkills.SpearNextLvl + 1
Call SkillStatLevel("spear")
End If
If PlayerSkills.StaffNxtLvl <= 0 Then
PlayerSkills.Staff = PlayerSkills.Staff + 1
Call SkillStatLevel("staff")
End If
If PlayerSkills.SwordNextLvl <= 0 Then
PlayerSkills.Sword = PlayerSkills.Sword + 1
Call SkillStatLevel("sword")
End If
If PlayerSkills.theforceNextLvl <= 0 Then
PlayerSkills.TheForce = PlayerSkills.TheForce + 1
Call SkillStatLevel("force")
End If
If PlayerSkills.UnarmedAttackNextLevel <= 0 Then
PlayerSkills.UnarmedAttack = PlayerSkills.UnarmedAttack + 1
Call SkillStatLevel("unarmedattack")
End If
End Sub
Public Sub Boss1Deafeated()
If MainPlayer.MapName = "newbie" Then Exit Sub
Form1.txtDialog.Text = Form1.txtDialog.Text & "You have defeated the Elk Warrior and have gained the Magic Book and some money" & vbNewLine
I = 0
For I = I + 1 To MaxItems
    If PXItem(I).ItemName = "MagicBook" Then
        PXItem(I).Quanity = PXItem(I).Quanity + 1
        I = MaxItems
    End If
Next I
MainPlayer.Money = MainPlayer.Money + 2000
TOOLS.LoadStats
Form1.Tiles.Picture = LoadPicture(App.Path & "\TILES.BMP")
Open App.Path & "\maps\newbie" For Input As #1: Close
Open App.Path & "\maps\newbie" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

Form1.MapName.Caption = "newbie"
Call Stuff.MovenewMonster2
Form1.ShowMonsters
Form1.HScroll1.Value = "45"
'
Form1.VScroll1.Value = "23"

MediaPlayer1.Filename = App.Path & "\" & "z5oot[2].mid"
Form1.RedrawMap
MainPlayer.MapName = "newbie"
End Sub
'this is one of the most important functions!
'this IS ONE OF MY BIGGEST ACCOMPLISHMENTS
'other types of quest still need to be added
Public Sub QuestGenerator()
Dim QuestMapX As Integer
Dim QuestMapY As Integer
Dim OthermapNumber As Integer
Randomize
OthermapNumber = Rnd * 10
QuestNumber = Rnd * 10
If QuestNumber = 0 Or QuestNumber = 1 Or QuestNumber = 2 Then
'defeat a monster
If QuestNumber = 0 Then
QuestNumber = 1
End If
If OthermapNumber = 0 Then
OthermapNumber = 1
End If

Form1.txtDialog.Text = Form1.txtDialog.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
'Add the quest to your quest page!
TOOLS.Text2.Text = TOOLS.Text2.Text & " " & vbCrLf
TOOLS.Text2.Text = TOOLS.Text2.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
'calls create monster
Stuff.CreateQuestMonster QuestNumber * 15, OthermapNumber * 4, 1, 1, QuestNumber, OthermapNumber
'saves your quest
 F = FreeFile
Open (App.Path & "\scripts\quest" & NumQuest & ".qst") For Binary Access Write Lock Write As F
Put #F, , QuestNumber & vbCrLf   'questnumber
Put #F, , OthermapNumber & vbCrLf 'othermapnumber
Put #F, , QuestMapX & vbCrLf 'questmapx
Put #F, , QuestMapY & vbCrLf  'questmapy
Close F
NumQuest = NumQuest + 1
End If
If QuestNumber = 3 Or QuestNumber = 4 Then
'wizard power
If OthermapNumber = 0 Then
OthermapNumber = 1
End If

Form1.txtDialog.Text = Form1.txtDialog.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
TOOLS.Text2.Text = TOOLS.Text2.Text & " " & vbCrLf
TOOLS.Text2.Text = TOOLS.Text2.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
'calls create monster
Stuff.CreateQuestMonster QuestNumber * 5, OthermapNumber * 4, 1, 1, QuestNumber, OthermapNumber
'saves your quest
 F = FreeFile
Open (App.Path & "\scripts\quest" & NumQuest & ".qst") For Binary Access Write Lock Write As F
Put #F, , QuestNumber & vbCrLf   'questnumber
Put #F, , OthermapNumber & vbCrLf 'othermapnumber
Put #F, , QuestMapX & vbCrLf 'questmapx
Put #F, , QuestMapY & vbCrLf  'questmapy
Close F
NumQuest = NumQuest + 1
End If
If QuestNumber = 5 Or QuestNumber = 6 Then
'world hunt
If OthermapNumber = 0 Then
OthermapNumber = 1
End If
Form1.txtDialog.Text = Form1.txtDialog.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
TOOLS.Text2.Text = TOOLS.Text2.Text & " " & vbCrLf
TOOLS.Text2.Text = TOOLS.Text2.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
'calls create monster
Stuff.CreateQuestMonster QuestNumber * 5, OthermapNumber * 4, 1, 1, QuestNumber, OthermapNumber
'saves your quest
 F = FreeFile
Open (App.Path & "\scripts\quest" & NumQuest & ".qst") For Binary Access Write Lock Write As F
Put #F, , QuestNumber & vbCrLf   'questnumber
Put #F, , OthermapNumber & vbCrLf 'othermapnumber
Put #F, , QuestMapX & vbCrLf 'questmapx
Put #F, , QuestMapY & vbCrLf  'questmapy
Close F
NumQuest = NumQuest + 1
End If
If QuestNumber = 7 Or QuestNumber = 8 Then
'treasure hunt
If OthermapNumber = 0 Then
OthermapNumber = 1
End If
Form1.txtDialog.Text = Form1.txtDialog.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
TOOLS.Text2.Text = TOOLS.Text2.Text & " " & vbCrLf
TOOLS.Text2.Text = TOOLS.Text2.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
'calls create monster
Stuff.CreateQuestMonster QuestNumber * 5, OthermapNumber * 4, 1, 1, QuestNumber, OthermapNumber
'saves your quest
 F = FreeFile
Open (App.Path & "\scripts\quest" & NumQuest & ".qst") For Binary Access Write Lock Write As F
Put #F, , QuestNumber & vbCrLf   'questnumber
Put #F, , OthermapNumber & vbCrLf 'othermapnumber
Put #F, , QuestMapX & vbCrLf 'questmapx
Put #F, , QuestMapY & vbCrLf  'questmapy
Close F
NumQuest = NumQuest + 1
End If
If QuestNumber = 9 Or QuestNumber = 10 Then
'talk to an npc
If OthermapNumber = 0 Then
OthermapNumber = 1
End If
Form1.txtDialog.Text = Form1.txtDialog.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
TOOLS.Text2.Text = TOOLS.Text2.Text & " " & vbCrLf
TOOLS.Text2.Text = TOOLS.Text2.Text & "In this quest you will have to defeat a monster and if you kill it you may be rewared nicely!  Your Monster Is Located in sector " & QuestNumber & "." & OthermapNumber & vbNewLine
'calls create monster
Stuff.CreateQuestMonster QuestNumber * 5, OthermapNumber * 4, 1, 1, QuestNumber, OthermapNumber
'saves your quest
 F = FreeFile
Open (App.Path & "\scripts\quest" & NumQuest & ".qst") For Binary Access Write Lock Write As F
Put #F, , QuestNumber & vbCrLf   'questnumber
Put #F, , OthermapNumber & vbCrLf 'othermapnumber
Put #F, , QuestMapX & vbCrLf 'questmapx
Put #F, , QuestMapY & vbCrLf  'questmapy
Close F
NumQuest = NumQuest + 1
End If
End Sub
Public Sub CheckIfYouFinishedQuest()

End Sub
Private Sub SpecialData(X As Integer, y As Integer, MapOption As String)

End Sub
Public Sub Miner(X As Integer, y As Integer)

If Map(X, y).Tiletype = Gold Then
    If SpecialMap(X, y).Choice = 2 Then
        If SpecialMap(X, y).MineAmount <= 0 Then
            BitBlt Form1.hDC, Snap(X, 32), Snap(y, 32), 32, 32, Form1.Tiles.hDC, 11904, 0, SRCCOPY
            'Save map info to the map array
            Map(X, y).TileX = 11904
            Map(X, y).Tiley = 0
            Map(X, y).Tiletype = WALKABLE
            Form1.RedrawMap
            Form1.txtDialog.Text = Form1.txtDialog.Text & "This location is now empty or resources" & vbNewLine
            'todo make a save map sub for online play
            'or send info to server that a person attempted to mine
            
            MainPlayer.XP = MainPlayer.XP + 5
            MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 5
            Form1.AddPlayerStats "Xp", 5
        Else
            SpecialMap(X, y).MineAmount = SpecialMap(X, y).MineAmount - 10
            Form1.txtDialog.Text = Form1.txtDialog.Text & "You got 10 gold" & vbNewLine
            MainPlayer.Money = MainPlayer.Money + 10
            TOOLS.LoadStats
End If
End If

    Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "Unable to mine that area" & vbNewLine
Exit Sub
End If
End Sub
Public Sub NpcFunctions(Command As String, Target As Integer)
'to do add the code for some of the functions
If Command = "/hail" Then
RPG.TalkToNpc (Target)
End If
If Command = "/trade" Then

End If
If Command = "/gossip" Then
On Error GoTo Nogossip
    Open App.Path & "\scripts\" & Target & ".goss" For Input As #1
        Line Input #1, Line1 'x
    Close #1
    Form1.txtDialog.Text = Form1.txtDialog.Text & Line1 & vbNewLine
Exit Sub
Nogossip:
Form1.txtDialog.Text = Form1.txtDialog.Text & "I do not like to gossip about other people" & vbNewLine
Exit Sub
End If

'banker functions
If Command = "/balance" Then

End If
If Command = "/deposit" Then

End If '
If Command = "/withdraw" Then

End If
'animal functions
If Command = "/tame" Then

End If
If Command = "/feed" Then

End If
If Command = "/follow" Then

End If
If Command = "/stop" Then

End If
If Command = "/free" Then

End If
If Command = "/animalattack" Then

End If
If Command = "/animalstats" Then

End If
End Sub
Public Sub FastXYWarp(X As Integer, y As Integer)
'This places a person on the x, y cordintes specifed
'it is very fast
Do Until MainPlayer.xpos = X And MainPlayer.ypos = y
MainPlayer.xpos = (Snap(Form1.Image1.Left, 32) \ 32) + HScroll
MainPlayer.ypos = (Snap(Form1.Image1.Top, 32) \ 32) + VScroll
Warping = True
    If MainPlayer.xpos > X Then
        If Form1.HScroll1.Value > 0 Then
            Form1.HScroll1.Value = Form1.HScroll1.Value - 1
        Else
            Form1.Image1.Left = Form1.Image1.Left - 32
        End If
    End If
    If MainPlayer.xpos < X Then
        If Form1.HScroll1.Value < Form1.HScroll1.Max Then
            Form1.HScroll1.Value = Form1.HScroll1.Value + 1
        Else
            Form1.Image1.Left = Form1.Image1.Left + 32
        End If
    End If
    If MainPlayer.ypos > y Then
        If Form1.VScroll1.Value > 0 Then
            Form1.VScroll1.Value = Form1.VScroll1.Value - 1
        Else
            Form1.Image1.Top = Form1.Image1.Top - 32
        End If
    End If
    If MainPlayer.ypos < y Then
        If Form1.VScroll1.Value < Form1.VScroll1.Max Then
            Form1.VScroll1.Value = Form1.VScroll1.Value + 1
        Else
            Form1.Image1.Top = Form1.Image1.Top + 32
        End If
    End If

Loop
Warping = False
Form1.RedrawMap
'Adjust the players caption label
Form1.lblplayername(0).Top = Form1.Image1.Top - 20
Form1.lblplayername(0).Left = Form1.Image1.Left
End Sub


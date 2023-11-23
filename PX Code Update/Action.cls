VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NOT USED yet purpose to make traps in dungeons
'This class controls events that happen in the game
'for example
'· Trap (Single) - affects just the character that triggered the trap
'· Trap (Multi) - affects the entire area
'· Teleport - teleports the character to another place on the map
'· Heal
'· Map Teleport - teleports the character to another map
'· Spawn Monster
'· Spawn Item
'· Lock - locks a door
'· Unlock - unlocks a door
'· Message - displays a message to the character
'· Player Start - the position the character starts when a map is loaded

'I am still working on this and is not yet used!


'In order for teleport to work you must give
'X cordinate
'Y cordinate
'Mapname

'How the prefixes work for  variables i = integer s = string b= boolean
'File name for Monsters and Items
Dim SFileName As String
'Tile Switch variables
Private iTileType As Integer
Private iTilex As Integer
Private iTiley As Integer
'Tile coords for action - can be different than tile triggering action
Private iX As Integer
Private iY As Integer
'Action Type - see ActionType enum
Private iActionType As Integer
'Specifies if the action is to be reset after triggering
Private bReset As Boolean
'Time before action is reset
Private iResetTime As Integer
'Delay from time trigger is pressed until action occurs
Private iDelay As Integer
'Should action be triggered
Private bTriggerAction As Boolean
'Message that is displayed when the action is triggered
Private sMessage As String
'Is the message shown to just everyone in the party
Private bMsgShowAll As Boolean
'Type of trap that is triggered
Private iTrapType As Integer
' monster that spawns when action is triggered
Private iSpawnIndex As Integer
'item that spawns when action is triggered
Private iItemIndex As Integer
'Map that character is teleported to
Private sLevelName As String
'Amount of hp the character gets healed
Private HealAmount As Integer
'Amount of stamina the character gets
Private ReviveAmount As Integer

Public Property Get X() As Integer
    X = iX
End Property

Public Property Let X(ByVal iData As Integer)
    iX = iData
End Property

Public Property Get y() As Integer
    y = iY
End Property

Public Property Let y(ByVal iData As Integer)
    iY = iData
End Property

Public Property Get ActionType() As Integer
    ActionType = iActionType
    MsgBox ActionType
End Property

Public Property Let ActionType(ByVal iData As Integer)
    iActionType = iData
End Property

Public Property Get Reset() As Boolean
    Reset = bReset
End Property

Public Property Let Reset(ByVal bData As Boolean)
    bReset = bData
End Property

Public Property Get ResetTime() As Integer
    ResetTime = iResetTime
End Property

Public Property Let ResetTime(ByVal iData As Integer)
    iResetTime = iData
End Property

Public Property Get Delay() As Integer
    Delay = iDelay
End Property

Public Property Let Delay(ByVal iData As Integer)
    iDelay = iData
End Property

Public Property Get TriggerAction() As Boolean
    TriggerAction = bTriggerAction
End Property

Public Property Let TriggerAction(ByVal bData As Boolean)
    bTriggerAction = bData
End Property

Public Property Get Message() As String
    Message = sMessage
Form1.txtDialog.Text = Form1.txtDialog.Text & Message & vbCrLf
   '' MsgBox "Get"
End Property

Public Property Let Message(ByVal sData As String)
    sMessage = sData
  '  MsgBox "Let"
End Property

Public Property Get MsgShowAll() As Boolean
    MsgShowAll = bMsgShowAll
End Property

Public Property Let MsgShowAll(ByVal bData As Boolean)
    bMsgShowAll = bData
End Property

Public Property Get TrapType() As Integer
    TrapType = iTrapType
End Property

Public Property Let TrapType(ByVal iData As Integer)
    iTrapType = iData
End Property
Public Property Get TileX() As Integer
    TileX = iTilex
End Property

Public Property Let TileX(ByVal iData As Integer)
    iTilex = iData
End Property
Public Property Get Tiley() As Integer
    Tiley = iTiley
End Property

Public Property Let Tiley(ByVal iData As Integer)
    iTiley = iData
End Property
Public Property Get Tiletype() As Integer
    Tiletype = iTileType
End Property

Public Property Let Tiletype(ByVal iData As Integer)
    iTileType = iData
End Property
Public Property Get SpawnIndex() As Integer
    SpawnIndex = iSpawnIndex
End Property

Public Property Let SpawnIndex(ByVal iData As Integer)
    iSpawnIndex = iData
End Property

Public Property Get LevelName() As String
    LevelName = sLevelName
End Property

Public Property Let LevelName(ByVal sData As String)
    sLevelName = sData
End Property
Public Property Get Filename() As String
    Filename = SFileName
End Property

Public Property Let Filename(ByVal sData As String)
    SFileName = sData
End Property
Public Property Get Heal() As Integer
    Heal = HealAmount
    If MainPlayer.Health + Heal > MainPlayer.MaxHealth Then
    MainPlayer.Health = MainPlayer.MaxHealth
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You are completly healed!" & vbNewLine
    Else
    MainPlayer.Health = MainPlayer.Health + Heal
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You are healed for " & Heal & vbNewLine
    End If
    
End Property

Public Property Let Heal(ByVal sData As Integer)
    HealAmount = sData
End Property
Private Sub SpawnMonster(X As Single, y As Single, MapName As String, MonsterFileName As String)
Dim I As Integer
I = -1
For I = I + 1 To 250

   If Monster(I).PlayerName = "" Then
   On Error GoTo nofile
    Open App.Path & "\monsters\" & MonsterFileName For Input As #1
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
    Monster(I).PlayerName = textline1
    Monster(I).Damage = textline2
    Monster(I).Speed = textline3
    Monster(I).MonsterMove = textline4
    Monster(I).Weapon = textline5
    Monster(I).Armor = textline6
    Monster(I).Strength = textline7
    Monster(I).Level = textline8
    Monster(I).MaxHealth = textline9
    Monster(I).Health = Monster(I).MaxHealth
    Monster(I).MaxStamina = textline10
    Monster(I).Intelligence = textline11
    Monster(I).MaxMana = textline12
    Monster(I).Scan = textline13
    Monster(I).CastSpell = textline14
    Monster(I).Stealth = textline15
    Monster(I).MeleeDefense = textline17
    Monster(I).MagicDefense = textline16
    Monster(I).Mana = Monster(I).MaxMana
    Monster(I).Stamina = Monster(I).MaxStamina
    Monster(I).Armor = textline6
    Monster(I).MonsterAttack2 = textline20
    Monster(I).xpos = X
    Monster(I).ypos = y
   Set P51 = Form1.Controls.Add("VB.Image", "MonsterImage" & I)

   With P51
      .Enabled = True
      .Top = y
      .Left = X
      .Height = 32
      .Width = 32
      .Picture = LoadPicture(App.Path & "\monsters\" & Monster(I).PlayerName & ".gif")
      .Visible = True
      .ToolTipText = Monster(I).PlayerName & " Level " & Monster(I).Level & " Health " & Monster(I).Health & "\" & Monster(I).MaxHealth
   End With
   Monster(K).ImageName = I
   Monster(K).MapName = MapName
nofile:
Exit Sub
End If
   I = 250
Next I
End Sub
Private Sub SpawnItem(X As Single, y As Single, MapName As String, ItemFileName As String)
K = -1
For K = K + 1 To 250
   If NewItem(K).ItemName = "" Then
   On Error GoTo nofile
    Open App.Path & "\monsters\" & ItemFileName For Input As #1
    Line Input #1, textline1 'name
    Line Input #1, textline2 'damage
    Line Input #1, textline3 'speed
    Line Input #1, textline4 'monstermove
    Line Input #1, textline5 'weapon
    Line Input #1, textline6 'armor
    Close #1
     NewItem(K).ItemName = textline1
     NewItem(K).xpos = y
     NewItem(K).ypos = X
C = Form1.PicItem.UBound + 1
Load Form1.PicItem(C)
   With Form1.PicItem(C)
      .Enabled = True
      .Top = y
      .Left = X
      .Height = 32
      .Width = 32
      .Picture = LoadPicture(App.Path & "\monsters\" & NewItem(K).ItemName & ".gif")
      .Visible = True
      .ToolTipText = NewItem(K).ItemName
   End With
   NewItem(K).ImageName = C
   NewItem(K).MapName = MapName
nofile:
Exit Sub
End If
   K = 250
Next K
End Sub
Private Sub WarpPlayer(X As Integer, y As Integer, MapName As String)

End Sub
Public Sub TileSwitch(X As Integer, y As Integer, TileX As Integer, Tiley As Integer, Tiletype As Integer)
    Map(X, y).TileX = TileX
    Map(X, y).Tiley = Tiley
    Map(X, y).Tiletype = Tiletype
End Sub
Sub HandleActions()

If iActionType = Action_Trap_Single Then

End If
If iActionType = ACTION_TRAP_MULTI Then

End If
If iActionType = ACTION_TELEPORT Then

End If
If iActionType = ACTION_LEVEL_CHANGE Then

End If
If iActionType = ACTION_SPAWN_MONSTER Then

End If
If iActionType = ACTION_SPAWN_ITEM Then

End If
If iActionType = ACTION_LOCK Then

End If
If iActionType = ACTION_UNLOCK Then

End If
If iActionType = ACTION_MESSAGE Then

End If
If iActionType = ACTION_PLAYER_START Then

End If


End Sub

Attribute VB_Name = "ItemInfo"
'What you need to do:
'edit the six line on each of the .item files
'and change it so it is the proper type like Weapon, Armor, or Item.  The default is weapon


Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Public Type Bitmap '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
    Public Const LR_CREATEDIBSECTION = &H2000
    Public Const LR_LOADFROMFILE = &H10


'exit tile stuff
Global ExitMap As String
Global ExitHscroll As Integer
Global ExitVscroll As Integer
Global ExitLeft As Integer
Global ExitTop As Integer
Global ExitTiles As String
Global ExitDirection As String


'this holds the items on the map
Public Type Itemtype
ImageName As String
CurX As Integer
CurY As Integer
MapName As String
xpos As Integer
ypos As Integer
ItemName As String ' * 30 '30 just used to save memory
Burden As Integer
'UseTimes As Integer 'How many times an item can be used
End Type
Global zzz As Integer
'this holds all the item quanitys
Type ItemNumType
ItemName As String
Burden As Integer
Skill As String * 25
Quanity As Integer
MinDamage As Integer
MaxDamage As Integer
Armor As Integer
MagicArmor As Integer
Value As Integer
Food As Integer
Speed As Integer
Decay As Integer
Health As Integer
Mana As Integer
Stamina As Integer
ItemPic As Integer
ItemKind As Integer '1 weapon 2 armor 3 Item 7 sheild 4 pants 5 shoes 6 helment
End Type

Public Const MaxItems = 200 'increase this to add more items to the game
Global PXItem(MaxItems) As ItemNumType
Global NewItem(999) As Itemtype

'Stuff for animating tiles
Public Type Animatetype
MapName As String
Tile1x As Integer
Tile1y As Integer
Tile2x As Integer
Tile2y As Integer
Tile1Visible As Boolean
xpos As Integer
ypos As Integer
End Type
Private Type BossType
Boss1Defeated As Boolean
Boss2Defeated As Boolean
End Type
Global AnimateTileNumber As Integer
Global Bosses(1) As BossType
Global AnimateTile(999) As Animatetype

Public Sub Createitem(X As Single, y As Single, K As Integer)

Static I As Integer


   If NewItem(K).ItemName = "" Then
   
    Open App.Path & "\monsters\" & AddMonster.File3.Filename For Input As #1
    Line Input #1, textline1 'name
    Line Input #1, textline2 'damage
    Line Input #1, textline3 'speed
    Line Input #1, textline4 'monstermove
    Line Input #1, textline5 'weapon
    Line Input #1, textline6 'armor

    Close #1

    zzz = zzz + 1
     NewItem(K).ItemName = textline1
     NewItem(K).xpos = AddMonster.X.Text
     NewItem(K).ypos = AddMonster.y.Text

C = Form1.PicItem.UBound + 1
Load Form1.PicItem(C)
   'With P51
   With Form1.PicItem(C)
      .Enabled = True
      .Top = AddMonster.y.Text
      .Left = AddMonster.X.Text
      .Height = AddMonster.Picture2.Height
      .Width = AddMonster.Picture2.Width
      .Picture = LoadPicture(App.Path & "\monsters\" & NewItem(K).ItemName & ".gif")
      .Visible = True
      .ToolTipText = NewItem(K).ItemName

   End With
     

   NewItem(K).ImageName = C 'zzz
   NewItem(K).MapName = MainPlayer.MapName

    
    End If
    
End Sub
Public Sub SaveItemLocation()
Dim K As Integer
K = 0
For K = K + 1 To 250
newmon = "ItemImage" & K
On Error GoTo nofile:
If NewItem(K).ItemName = "" Then
Exit Sub
Exit For
End If
 F = FreeFile
Open (App.Path & "\monsters\Item" & K & ".xy") For Binary Access Write Lock Write As F
    Put #F, , NewItem(K).xpos & vbCrLf   'level
    Put #F, , NewItem(K).ypos & vbCrLf
    Put #F, , NewItem(K).MapName & vbCrLf  'level
    Put #F, , NewItem(K).ItemName & vbCrLf    'level
    Put #F, , NewItem(K).Burden & vbCrLf
Close F

If NewItem(K).ItemName = "" Then
       Exit For
       End If
Next
Debug.Print "###Items Saved###"
Exit Sub
nofile:
Debug.Print "!!!Items Not Saved!!!"
Exit Sub
End Sub
Public Sub LoadItems()
Dim K As Integer

K = 0
On Error GoTo nofile:
For K = K + 1 To 250
Open App.Path & "\monsters\item" & K & ".xy" For Input As #1
Line Input #1, Line1 'x
Line Input #1, LINE2 'y
Line Input #1, Line3 'map
Line Input #1, Line4 'name
Line Input #1, Line5 ' text
 Close #1



Static I As Integer


   If NewItem(K).ItemName = "" Then
   
    Open App.Path & "\monsters\" & Line4 & ".item" For Input As #1
    Line Input #1, textline1 'name
    Line Input #1, textline2 'damage
    Line Input #1, textline3 'speed
    Line Input #1, textline4 'monstermove
    Line Input #1, textline5 'weapon
    Line Input #1, textline6 'armor

    Close #1

    zzz = zzz + 1
      NewItem(K).ItemName = textline1
      NewItem(K).xpos = Line1
      NewItem(K).ypos = LINE2
   'Set P51 = Form1.Controls.Add("VB.Image", "ItemImage" & zzz)
C = Form1.PicItem.UBound + 1
Load Form1.PicItem(C)
  ' With P51
  With Form1.PicItem(C)
      .Enabled = True
      .Top = Line1
      .Left = LINE2
      .Height = 10
      .Width = 10
      .Picture = LoadPicture(App.Path & "\monsters\" & NewItem(K).ItemName & ".gif")
      .Visible = True
      .ToolTipText = NewItem(K).ItemName

   End With
     

   NewItem(K).ImageName = C 'zzz
   NewItem(K).MapName = Line3

    
    End If
    
       If NewItem(K).ItemName = "" Then
       Exit For
       End If
Next
nofile:
Exit Sub
End Sub
Public Sub ShowItem()

Dim newitem2 As String
Dim K As Integer
Call ShowPlayerGuy
K = 0

For K = K + 1 To 250
newitem2 = "ItemImage" & NewItem(K).ImageName
On Error GoTo nofile:

 With Form1.PicItem(K)
If NewItem(K).MapName = MainPlayer.MapName And NewItem(K).ImageName = K Then
.Visible = True
Else
.Visible = False
End If

End With
If NewItem(K).ItemName = "" Then
       Exit For
       End If
Next
nofile:
Exit Sub
End Sub
Public Sub MovenewItem()
Dim newitem2 As String
Dim K As Integer
K = 0
Call MovenewPlayer

For K = K + 1 To 250
newitem2 = "ItemImage" & K
On Error GoTo nofile:
 ' Set P51 = Form1.Controls(newitem2)
'With P51
 With Form1.PicItem(K)
NewItem(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
 NewItem(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
If NewItem(K).CurX > NewItem(K).xpos Then
.Left = .Left - 32
End If
If NewItem(K).CurX < NewItem(K).xpos Then
.Left = .Left + 32

End If

If NewItem(K).CurY > NewItem(K).ypos Then
.Top = .Top - 32
End If
If NewItem(K).CurY < NewItem(K).ypos Then
.Top = .Top + 32

End If
End With
If NewItem(K).ItemName = "" Then
       Exit For
       End If
Next
nofile:

Exit Sub
End Sub
Public Sub WeaponInfo()
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = TOOLS.InfoName.Caption Then
        TOOLS.Damage.Caption = PXItem(I).MinDamage & "-" & PXItem(I).MaxDamage
        TOOLS.Damage.Caption = "Burden " & PXItem(I).Burden
        I = MaxItems
End If
Next I

End Sub
Public Sub ArmorInfo()
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = TOOLS.InfoName.Caption Then
        TOOLS.lblDamage.Caption = "Armor Level " & PXItem(I).Armor
        TOOLS.Damage.Caption = "Burden " & PXItem(I).Burden
        I = MaxItems
End If
Next I

End Sub
Public Sub SpellInfo()
I = -1
For I = I + 1 To MAXSpells
    If Spells(I).Spellname = TOOLS.InfoName.Caption Then
       TOOLS.lblDamage.Caption = "Damage:"
       TOOLS.Damage.Caption = Spells(I).MinDamage & "-" & Spells(I).MaxDamage
       TOOLS.lblspeed.Caption = "Cast Spell Skill: " & Spells(I).CSskill
       TOOLS.Speed.Caption = "Mana Needed: " & Spells(I).Mana
    End If
Next I
End Sub
Public Sub UseSpell()
'use spell code i have fixed this up a lot
Dim K As Integer
Dim attack As Integer
 Targetx = (Snap(Form1.TargetShape.Left, 32) \ 32) + HScroll
 Targety = (Snap(Form1.TargetShape.Top, 32) \ 32) + VScroll
I = -1
K = -1
For I = I + 1 To MAXSpells
    If Spells(I).Spellname = TOOLS.Spell.ToolTipText Then
    
        If PlayerSkill.CastSpell >= Spells(I).CSskill Then
            If MainPlayer.Mana >= Spells(I).Mana Then
            If Spells(I).MaxDamage > 0 Then
                For K = K + 1 To 250
                If Monster(K).CurX = Targetx And Monster(K).CurY = Targety Or Monster(K).CurX = Targetx And Monster(K).CurY + 1 = Targety Or Monster(K).CurX = Targetx And Monster(K).CurY - 1 = Targety Or Monster(K).CurX + 1 = Targetx And Monster(K).CurY = Targety Or Monster(K).CurX - 1 = Targetx And Monster(K).CurY = Targety Or Monster(K).CurX + 1 = Targetx And Monster(K).CurY + 1 = Targety Then
                If Monster(K).MagicDefense >= PlayerSkills.CastSpell Then
                Form1.lblAttack.Visible = True
                Form1.lblAttack.Caption = "Monster resists your spell!"
                MainPlayer.XP = MainPlayer.XP + Monster(K).Level
                MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
                MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
                Form1.TimerAttacklbl.Enabled = True
                PlayerSkills.CastSpellNextLevel = PlayerSkills.CastSpellNextLevel - Monster(K).Level
                Call RPG.CheckIfStatLeveled
                Exit Sub
                Else
                        'You attack monster
                        Randomize
                        attack = (Rnd * Spells(I).MaxDamage)
                            If attack < Spells(I).MinDamage Then
                                attack = Spells(I).MinDamage
                            End If
                Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
                Monster(K).Health = Monster(K).Health - attack
                Call Monsters.checkMonsterdeath
                Form1.lblAttack.Visible = True
                Form1.lblAttack = "Your spell does " & attack & " damage!"
                MainPlayer.XP = MainPlayer.XP + Monster(K).Level
                MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
                MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
                Form1.TimerAttacklbl.Enabled = True
                PlayerSkills.CastSpellNextLevel = PlayerSkills.CastSpellNextLevel - Monster(K).Level
                Call RPG.CheckIfStatLeveled
                Exit Sub
                End If
                End If
                Next K
                End If
                    Else
                    Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough mana to cast this spell!" & vbCrLf
                    End If
                        Else
                             Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough cast spell to cast this spell!" & vbCrLf
                             End If
            If MainPlayer.Health + Spells(I).AddHealth > MainPlayer.MaxHealth Then MainPlayer.Health = MainPlayer.MaxHealth
            If MainPlayer.Stamina + Spells(I).AddStamina > MainPlayer.MaxStamina Then MainPlayer.Stamina = MainPlayer.MaxStamina
            MainPlayer.Health = MainPlayer.Health + Spells(I).AddHealth
            MainPlayer.Mana = MainPlayer.Mana - Spells(I).Mana
            MainPlayer.XP = MainPlayer.XP + Spells(I).Mana
            MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Spells(I).Mana
            MainPlayer.SpendXp = MainPlayer.SpendXp + Spells(I).Mana
            PlayerSkills.CastSpellNextLevel = PlayerSkills.CastSpellNextLevel - Spells(I).Mana
            Call RPG.CheckIfStatLeveled
            Call TOOLS.LoadStats
            I = MAXSpells
        End If
    
Next I

End Sub
Public Sub UseWeapon()
'1=very fast
'2=fast
'3=aveage
'4 slow
'5 very slow
'6 or higher even slower
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = TOOLS.Pweapon.ToolTipText Then
        MainPlayer.CurrentWeaponSkill = PXItem(I).Skill
        MainPlayer.WeaponSpeed = PXItem(I).Speed
        MainPlayer.MinDamge = PXItem(I).MinDamage
        MainPlayer.MaxDamge = PXItem(I).MaxDamage
        I = MaxItems
End If
Next I

End Sub
Public Sub UseItem()
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = ItemName Then
    If PXItem(I).Quanity > 0 Then
        MainPlayer.Food = MainPlayer.Food + PXItem(I).Food
        MainPlayer.Health = MainPlayer.Health + PXItem(I).Health
        MainPlayer.Stamina = MainPlayer.Stamina + PXItem(I).Stamina
        MainPlayer.Mana = MainPlayer.Mana + PXItem(I).Mana
        PXItem(I).Quanity = PXItem(I).Quanity - 1
        I = MaxItems
    End If
End If
Next I
End Sub

Public Sub PickupItems(ItemX As Integer, ItemY As Integer)
Dim newmon As String
Dim K As Integer
Dim kk As Integer
Dim kkk As Integer
Dim newnewitem As String
Dim NameofItem As String
K = 0
For K = K + 1 To 250
'newmon = "ItemImage" & k
'newnewitem = "itemimage" & kk
On Error GoTo nofile:
If NewItem(K).xpos = ItemX And NewItem(K).ypos = ItemY Or NewItem(K).xpos = ItemX And NewItem(K).ypos = ItemY + 1 Or NewItem(K).xpos = ItemX + 1 And NewItem(K).ypos = ItemY Then
kk = NewItem(K).ImageName
NameofItem = NewItem(K).ItemName
newnewitem = "itemimage" & kk
'ItemNumbers(0).NameofItem = ItemNumbers(0).NameofItem + 1
'Form1.Caption = NewItem(k).ItemName
Call AddWeapons(NewItem(K).ItemName)

   'Set P51 = Form1.Controls("ItemImage" & kk)
 With Form1.PicItem(kk)
  ' With P51
        .ToolTipText = ""
        
      .Enabled = False
      .Visible = False
      .Picture = Nothing
      
   End With
'Else
'F.txtDialog.Text = Form1.txtDialog.Text & "Move closer to the item!" & vbCrLf
'orm1
End If
     If NewItem(K).ItemName = "" Then
       Exit For
       End If
Next
nofile:
'Form1.txtDialog.Text = Form1.txtDialog.Text & "Move closer to the item!" & vbCrLf

Exit Sub
End Sub
Private Sub AddWeapons(ItemName As String)
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = ItemName Then
    If PXItem(I).Quanity > 0 Then
        PXItem(I).Quanity = PXItem(I).Quanity + 1
        Call TOOLS.LoadStats
        I = MaxItems
    End If
End If
Next I
End Sub
Public Sub DropItem(ItemX As Integer, ItemY As Integer, ItemName As String)

'On Error GoTo nofile:
C = Form1.PicItem.UBound + 1
Load Form1.PicItem(C)
NewItem(C).ItemName = ItemName
NewItem(C).xpos = ItemX
NewItem(C).ypos = ItemY
With Form1.PicItem(C)
      .Enabled = True
      .Top = ItemX
      .Left = ItemY
      .Height = 10
      .Width = 10
      .Picture = LoadPicture(App.Path & "\monsters\" & NewItem(C).ItemName & ".gif")
      .Visible = True
      .ToolTipText = NewItem(C).ItemName

   End With
     
   NewItem(C).ImageName = C 'zzz
   NewItem(C).MapName = MainPlayer.MapName
    Call RemoveWeapons(ItemName)

'nofile:
'Exit Sub
End Sub


Private Sub RemoveWeapons(ItemName As String)
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = ItemName Then
    If PXItem(I).Quanity > 0 Then
        PXItem(I).Quanity = PXItem(I).Quanity - 1
        I = MaxItems
    End If
End If
Next I

If TOOLS.FrameEquip.Caption = "Items" Then
Call TOOLS.EquipList.Clear

Call ItemInfo.ItemsListing
End If

If TOOLS.FrameEquip.Caption = "Weapons" Then
Call TOOLS.EquipList.Clear

Call ItemInfo.EquipmentList
End If
If TOOLS.FrameEquip.Caption = "Armor" Then
Call TOOLS.EquipList.Clear

Call ItemInfo.ArmorList
End If

End Sub
Public Sub ShowPlayerGuy()

Dim newitem2 As String
Dim NewLabel As String
Dim K As Integer
K = 0
'newmon = "MonsterImage" & Monster(k).ImageName

For K = K + 1 To 250
newitem2 = "PlayerImage" & Player(K).ImageName
On Error GoTo nofile:
  Set P51 = Form1.Controls(newitem2)
With P51
If Player(K).MapName = MainPlayer.MapName Then
.Visible = True
Else
.Visible = False
End If

End With
NewLabel = "PlayerLabel" & Player(K).ImageName
Set LblPName = Form1.Controls(NewLabel)
With LblPName
If Player(K).MapName = MainPlayer.MapName Then
.Visible = True
Else
.Visible = False
End If
End With

If Player(K).PlayerName = "" Then
       Exit For
       End If
Next
nofile:
'Form1.Caption = Err.Description

Exit Sub
End Sub
'***************************************************************************************'
' Purpose:  Called when a new player has entered the online world          '                                                                        '
'***********************************************************************************'***'
Public Sub MovenewPlayer()
Dim newitem2 As String
Dim K As Integer
Dim Top1 As Integer
Dim Left1 As Integer
K = 0
Dim NewLabel As String
For K = K + 1 To 250
newitem2 = "PlayerImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newitem2)
With P51

Player(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Player(K).CurY = (Snap(.Top, 32) \ 32) + VScroll


If Player(K).CurX > Player(K).xpos Then
.Left = .Left - 32
End If
If Player(K).CurX < Player(K).xpos Then
.Left = .Left + 32

End If

If Player(K).CurY > Player(K).ypos Then
.Top = .Top - 32
End If
If Player(K).CurY < Player(K).ypos Then
.Top = .Top + 32

End If
Top1 = .Top
Left1 = .Left
End With
NewLabel = "PlayerLabel" & Player(K).ImageName
Set LblPName = Form1.Controls(NewLabel)
With LblPName
.Left = Left1
.Top = .Top + 16
End With

If Player(K).PlayerName = "" Then
       Exit For
       End If
Next
nofile:

Exit Sub
End Sub
Public Sub EquipmentList()
I = 0

For I = I + 1 To MaxItems
If PXItem(I).ItemKind = 1 Then
    If PXItem(I).Quanity > 0 Then
        TOOLS.EquipList.AddItem PXItem(I).ItemName
    End If
End If
Next I
End Sub
Public Sub ArmorList()
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemKind = 2 Then
    If PXItem(I).Quanity > 0 Then
        TOOLS.EquipList.AddItem PXItem(I).ItemName
    End If
End If
Next I
End Sub
Public Sub ItemsListing()
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemKind = 3 Then
    If PXItem(I).Quanity > 0 Then
        TOOLS.EquipList.AddItem PXItem(I).ItemName
    End If
End If
Next I
End Sub
Public Sub PXLoadItems()
'Loads all the items and sets the quantity to zero
TOOLS.ItemFile.Pattern = "*.item"
TOOLS.ItemFile.Path = App.Path & "\monsters\"
I = 0
For I = I + 1 To TOOLS.ItemFile.ListCount

If TOOLS.ItemFile.List(I) = "" Then
Else

    Open App.Path & "\monsters\" & TOOLS.ItemFile.List(I) For Input As #1
    Line Input #1, textline1 'name
    Line Input #1, textline2 'damage
    Line Input #1, textline3 'speed
    Line Input #1, textline4 'monstermove
    Line Input #1, textline5 'weapon
    Line Input #1, textline6 'Item Type
    Close #1
If textline6 = "Armor" Then
PXItem(I).ItemKind = 2
End If
If textline6 = "Weapon" Then
PXItem(I).ItemKind = 1
End If
If textline6 = "Item" Then
PXItem(I).ItemKind = 3
End If
PXItem(I).ItemName = textline1
PXItem(I).Speed = textline3
PXItem(I).Quanity = 0
PXItem(I).MaxDamage = textline2
End If
Next I
Debug.Print "###PX ITEMS Loaded###"
'Call LoadPlayerItems
Call LoadSpells
End Sub
Private Sub LoadSpells()
If MapEd = True Then Exit Sub

TOOLS.ItemFile.Pattern = "*.spel"
TOOLS.ItemFile.Path = App.Path & "\monsters\"

I = -1
For I = I + 1 To TOOLS.ItemFile.ListCount
If TOOLS.ItemFile.List(I) = "" Then
Else
'Debug.Print I
'Debug.Print TOOLS.ItemFile.List(I)
Open (App.Path & "\monsters\" & TOOLS.ItemFile.List(I)) For Input As #1
Input #1, Spells(I).Spellname, Spells(I).MinDamage, Spells(I).MaxDamage, Spells(I).CSskill, Spells(I).Duration
Input #1, Spells(I).AddHealth, Spells(I).AddStamina, Spells(I).AddMana, textline4, Spells(I).Graphic, Spells(I).Target, Spells(I).Posion, Spells(I).EffectArea
Close #1
'Debug.Print App.Path & "\monsters\" & TOOLS.ItemFile.List(I)

End If
Next I
Debug.Print "###Spells Loaded###"
End Sub
'**********************************************************************************************'
' Purpose:  Called when tools is loaded, opens a file and loads the data to the items array    '
'**********************************************************************************************'
Public Sub LoadPlayerItems()
'take out exit sub if you want to use it
Exit Sub

If MapEd = True Then Exit Sub
If OnlineGame = True Then Exit Sub

On Error GoTo nofile:
'This sub loads the player items from a file to an array
Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".itms" For Binary Access Read Lock Read As #2
Get #2, , PXItem
Close #2
Debug.Print "###Player Items Loaded###"
Exit Sub
nofile:
Debug.Print "!!!Player Items NOT Loaded!!!"
Exit Sub
End Sub
'***************************************************************************************'
' Purpose:  Called when form1 is unloaded. Saves the array of pxitem to a file          '
'***********************************************************************************'***'
Public Sub SavePlayerItems()
Open App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & ".itms" For Binary Access Write Lock Write As #2
Put #2, , PXItem
Close #2
Debug.Print "###Player Items Saved###"
End Sub
'***************************************************************************************'
' Purpose:  Called when change spell is clicked on tools.frm. Lists all spells          '                                                                        '
'***********************************************************************************'***'

Public Sub SpellsListing()
I = -1
For I = I + 1 To MAXSpells
    If Spells(I).Spellname = "" Then ' And Spells(i).Researched = True Then
        Else
        TOOLS.EquipList.AddItem Spells(I).Spellname
    End If
Next I

End Sub

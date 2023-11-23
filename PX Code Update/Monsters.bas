Attribute VB_Name = "Monsters"
Global MonsterNumber As Integer
Dim attackDamage As Integer
Dim zzz As Integer 'Holds the player number
Global MonsterAttacking As Boolean
Private Type PlayerType ' other players on the screen
    Admin As Boolean
    InUse As Boolean
    ImageName As Integer
    PlayerVisible As Boolean
    Cheater As Boolean
    Level As String
    CurX As Integer
    CurY As Integer
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
    Graphic As Integer
    xpos As Integer
    ypos As Integer
    VScroll As Integer
    HScroll As Integer
    Stealth As Boolean
    MapName As String
End Type

Private Type SkillsType
'Holds skill value
'and how much till next increase
    MeleeDefense As Integer
    Run As Integer
    CastSpell As Integer
    Swim As Integer
    Scan As Integer
    Stealth As Integer
    Sword As Integer
    SwordNextLvl As Integer
    Dagger As Integer
    DaggerNxtLvl As Integer
    Fail As Integer
    FailNextLvl As Integer
    Staff  As Integer
    StaffNxtLvl As Integer
    Spear As Integer
    SpearNextLvl As Integer
    TheForce As Integer
    theforceNextLvl As Integer
    UnarmedAttack As Integer
    Bow As Integer
    Bombs As Integer
    Blacksmith As Integer
    Miner As Integer
    Farmer As Integer
    Builder As Integer
    Hunter  As Integer
    MagicDefense As Integer
    Mace As Integer
    Axe As Integer
    AxeNextLevel As Integer
    MaceNextLevel As Integer
    BombsNextLevel As Integer
    BowNextLevel As Integer
    UnarmedAttackNextLevel As Integer
    Rancher As Integer
    Chemist As Integer
    Cook As Integer
    Seige As Integer
    Lumberjack As Integer
    SurvialSkill As Integer
    MeeleNextLevel As Integer
    MagicDefenseNextLevel As Integer
    CastSpellNextLevel As Integer
    StealthNextLevel As Integer
    
End Type
Const MAX_UNITS = 100
Dim zplayer As Integer
Global PlayerSkills As SkillsType
Global Player(MAX_UNITS) As PlayerType
Public Sub CreatePlayer(Mapz As String, X As Integer, y As Integer, PlayerName As String)
Dim K As Integer

K = 0

For K = K + 1 To 250


On Error GoTo nofile:

Static I As Integer


   If Player(K).PlayerName = "" Then
    

   zplayer = zplayer + 1
      Player(K).PlayerName = PlayerName
      Player(K).MapName = Mapz
      Player(K).xpos = X
      Player(K).ypos = y
   Set P51 = Form1.Controls.Add("VB.Image", "PlayerImage" & zplayer)

   With P51
      .Enabled = True
      .Top = y
      .Left = X
      .Height = 32
      .Width = 32
      .Picture = Pics.Image1.Picture
      .Visible = True
      .ToolTipText = Player(K).PlayerName

   End With
     

   Player(K).ImageName = zplayer
   '   Set LblPName = Form1.Controls.Add("VB.label", "PlayerLabel" & zplayer)

  ' With LblPName
      '.Caption = Player(k).PlayerName
      '.Enabled = True
      '.Top = Y
      '.Left = x
      '.Height = 32
     ' .Width = 32
    '  .Visible = True
      
   '   .AutoSize = True
  ' End With

    Exit For
    End If
    
       'If Player(k).PlayerName = "" Then
      'Exit For
      ' End If
Next
nofile:
Exit Sub
End Sub



Public Sub DamageforATTack()
If MainPlayer.xpos = Monster(MonsterNumber).CurX And MainPlayer.ypos = Monster(MonsterNumber).CurY Then

Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200

attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(MonsterNumber).Armor
If (attackDamage <= 0) Then
Damage = 0 'or 1
Monster(MonsterNumber).Stamina = Monster(MonsterNumber).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = "Chicken Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(MonsterNumber).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(MonsterNumber).Level

MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(MonsterNumber).Level

Form1.TimerAttacklbl.Enabled = True
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"

End If
If (attackDamage > 0) Then
If Monster(MonsterNumber).Health > 0 Then

MainPlayer.Stamina = MainPlayer.Stamina - 1

Monster(MonsterNumber).Stamina = Monster(MonsterNumber).Stamina - 1
Monster(MonsterNumber).Health = Monster(MonsterNumber).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage

Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True

Form1.lblAttack = "You attack " & Monster(MonsterNumber).PlayerName & " for " & attackDamage

Form1.TimerAttacklbl.Enabled = True
End If
End If
End If
Call TOOLS.LoadStats
'Call checkMonster
Call checkMonsterdeath
End Sub
Public Sub CheckDead()
If MainPlayer.Health <= 0 Then
    Monsters.MonsterReturn
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You have died!" & vbCrLf
    MainPlayer.Deaths = MainPlayer.Deaths + 1
    MainPlayer.SkillsPercent = MainPlayer.SkillsPercent - 5
    MainPlayer.MapName = "newbie"
    Form1.MapName = "newbie"
    Form1.MediaPlayer1.Filename = App.Path & "\" & "z5oot[2].mid"
Form1.Tiles.Picture = LoadPicture(App.Path & "\TILES.BMP")
   Open App.Path & "\maps\newbie" For Input As #1: Close
Open App.Path & "\maps\newbie" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

    Form1.RedrawMap
    Form1.HScroll1.Value = "45"
    Form1.VScroll1.Value = "23"
    MainPlayer.Health = MainPlayer.MaxHealth \ 2
    MainPlayer.Stamina = MainPlayer.MaxStamina \ 2
    MainPlayer.Mana = MainPlayer.MaxMana \ 2

    Monsters.MonsterReturn
I = 0
For I = I + 1 To 100
Monster(I).MonsterAttack = False
Next
Call TOOLS.LoadStats

End If



End Sub

Public Sub checkMonsterdeath()
Dim newmon As String
Dim K As Integer
K = 0

For K = K + 1 To 250
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51

If Monster(K).Health <= 0 Then
If .Tag = "dead" Then
Else
    MainPlayer.XP = MainPlayer.XP + Monster(K).Level * 10
    MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level * 10
    Form1.txtDialog.Text = Form1.txtDialog.Text & "You killed a " & Monster(K).PlayerName & vbCrLf
    MainPlayer.KILLS = MainPlayer.KILLS + 1
    Monster(K).Health = 0

.Tag = "dead"
Call Monsters.MonstersDrop(Monster(K).Level)

End If

'tells what picture to show when you kill the monster
If Monster(K).PlayerName = "Chicken" Then
.Picture = Pics.DeadChicken.Picture
Else
.Picture = LoadPicture(App.Path & "\Monsters\dead.gif")

End If
If Monster(K).MapName = "nowhere" Then
Else
Call MonQuest(K)
End If
End If
End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
nofile:

Exit Sub

End Sub
Public Sub attackMonster()
'core player attack code
Dim newmon As String
Dim K As Integer
K = 0

For K = K + 1 To 250
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
Call ItemInfo.UseWeapon
If MainPlayer.xpos = Monster(K).CurX And MainPlayer.ypos = Monster(K).CurY And .Visible = True Then
attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
attackDamage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
'+y
If MainPlayer.xpos + 1 = Monster(K).CurX And MainPlayer.ypos = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200

attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
Damage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
'-y
If MainPlayer.xpos - 1 = Monster(K).CurX And MainPlayer.ypos = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200
attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
attackDamage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
'x-
If MainPlayer.xpos = Monster(K).CurX And MainPlayer.ypos - 1 = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200

attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
Damage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Call RPG.CheckIfStatLeveled
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
'x+
If MainPlayer.xpos = Monster(K).CurX And MainPlayer.ypos + 1 = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200
attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
Damage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
'++
If MainPlayer.xpos + 1 = Monster(K).CurX And MainPlayer.ypos + 1 = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
Damage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If

If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Form1.TimerAttacklbl.Enabled = True
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
'-1+1
If MainPlayer.xpos - 1 = Monster(K).CurX And MainPlayer.ypos + 1 = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200
attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
attackDamage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
   .ToolTipText = Monster(K).PlayerName & " Level " & Monster(K).Level & " Health " & Monster(K).Health & "\" & Monster(K).MaxHealth
'+-
If MainPlayer.xpos + 1 = Monster(K).CurX And MainPlayer.ypos - 1 = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200
attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
attackDamage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then

MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath
   .ToolTipText = Monster(K).PlayerName & " Level " & Monster(K).Level & " Health " & Monster(K).Health & "\" & Monster(K).MaxHealth
'--
If MainPlayer.xpos - 1 = Monster(K).CurX And MainPlayer.ypos - 1 = Monster(K).CurY And .Visible = True Then
Call ItemInfo.UseWeapon
'attackDamage = 100 + Monster(MonsterNumber).Wisdom + Monster(MonsterNumber).Dexterity + Monster(MonsterNumber).Intelligence + Monster(MonsterNumber).Quickness + Monster(MonsterNumber).Strength - Monster(MonsterNumber).Armor - MainPlayer.Armor - MainPlayer.MeeleeDefense - MainPlayer.Dexterity - MainPlayer.Strength - MainPlayer.Quickness \ 2 - 5 ' \ 3200
attackDamage = (MainPlayer.MinDamge + (Rnd * (MainPlayer.MaxDamge - MainPlayer.MinDamge))) - Monster(K).Armor
If (attackDamage <= 0) Then
Damage = 0 'or 1
Monster(K).Stamina = Monster(K).Stamina - 1
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " Evades your attack"
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
MainPlayer.SpendXp = MainPlayer.SpendXp + Monster(K).Level
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
End If
If (attackDamage > 0) Then
If Monster(K).Health > 0 Then
MainPlayer.Stamina = MainPlayer.Stamina - 1
Monster(K).Stamina = Monster(K).Stamina - 1
Monster(K).Health = Monster(K).Health - attackDamage '+ MainPlayer.Armor
MainPlayer.XP = MainPlayer.XP + attackDamage
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - attackDamage
MainPlayer.SpendXp = MainPlayer.SpendXp + attackDamage
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Form1.lblAttack = "You attack " & Monster(K).PlayerName & " for " & attackDamage
Form1.TimerAttacklbl.Enabled = True
Call UpdateWeaponskill(Monster(K).Level)
Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
End If
End If
Call TOOLS.LoadStats
Call checkMonsterdeath

   .ToolTipText = Monster(K).PlayerName & " Level " & Monster(K).Level & " Health " & Monster(K).Health & "\" & Monster(K).MaxHealth
End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
nofile:
'MsgBox Err.Description & Err.Source
Exit Sub
End Sub
Public Sub MoveNpc()
'If An Npc has Script move it along it's path
'For making npc's move in pattern's
Dim GoOn As Boolean
Dim Firstk As Boolean
'Dim textline1 As Integer, textline2 As Integer, textline3 As Integer, textline4 As Integer
Dim K As Integer
Dim SlowDown As Integer
K = 0
Dim newtext As Integer
Dim newtext2 As Integer
Dim newtext3 As Integer
Dim newtext4 As Integer
'If Firstk = True Then
'Coolsub:
'k = 0

'End If


For K = K + 1 To 1
If GoOn = True Then
Newsub:
K = K + 1
End If
'SlowDown = SlowDown + 1
On Error GoTo nofile:
    Open App.Path & "\monsters\npc" & K & ".move" For Input As #1
   Line Input #1, textline1 'up
   Line Input #1, textline2 'down
   Line Input #1, textline3 'left
   Line Input #1, textline4 'right
   Line Input #1, textline5 'movement?

    Close #1
   ' MsgBox NPc(k).PlayerName
'If Left(textline1, 4) = "" Then
newtext = textline1 'Right(textline1, Len(textline1) - 4)
newtext2 = textline2
newtext3 = textline3
newtext4 = textline4
If newtext = 0 And newtext2 = 0 And newtext4 = 0 Then
GoOn = True
GoTo Newsub
End If
'End If

'If SlowDown = 0 Then
If newtext2 > 0 And textline5 = "down" Then
F = FreeFile
 NPc(K).ypos = NPc(K).ypos + newtext2
 
Open (App.Path & "\monsters\NPc" & K & ".move") For Binary Access Write Lock Write As F
    Put #F, , newtext & vbCrLf
    Put #F, , newtext2 & vbCrLf
    Put #F, , newtext3 & vbCrLf
    Put #F, , newtext4 & vbCrLf
    Put #F, , "right" & vbCrLf
Close F
Exit Sub
Exit For
End If
If newtext > 0 And textline5 = "up" Then
NPc(K).ypos = NPc(K).ypos - newtext
 'NPc(k).xpos = NPc(k).ypos - newtext
 F = FreeFile
Open (App.Path & "\monsters\NPc" & K & ".move") For Binary Access Write Lock Write As F
    Put #F, , newtext & vbCrLf
    Put #F, , newtext2 & vbCrLf
    Put #F, , newtext3 & vbCrLf
    Put #F, , newtext4 & vbCrLf
    Put #F, , "left" & vbCrLf
Close F
Exit Sub
Exit For
End If
'SlowDown = 0
'End If
'i 'f SlowDown = 0 Then

If newtext4 > 0 And textline5 = "right" Then
F = FreeFile

 NPc(K).xpos = NPc(K).xpos + newtext4
 
Open (App.Path & "\monsters\NPc" & K & ".move") For Binary Access Write Lock Write As F
    Put #F, , newtext & vbCrLf
    Put #F, , newtext2 & vbCrLf
    Put #F, , newtext3 & vbCrLf
    Put #F, , newtext4 & vbCrLf
    Put #F, , "up" & vbCrLf
Close F
Exit Sub
Exit For
End If
'SlowDown = 0
'End If
'If SlowDown > 15 Then
If newtext3 > 0 And textline5 = "left" Then
F = FreeFile
 NPc(K).xpos = NPc(K).xpos - newtext3
 
Open (App.Path & "\monsters\NPc" & K & ".move") For Binary Access Write Lock Write As F
    Put #F, , newtext & vbCrLf
    Put #F, , newtext2 & vbCrLf
    Put #F, , newtext3 & vbCrLf
    Put #F, , newtext4 & vbCrLf
    Put #F, , "down" & vbCrLf
Close F
Exit Sub
Exit For
End If
'SlowDown = 0
'End If
'If SlowDown = 0 Then

'SlowDown = 0
'End If
If NPc(K).PlayerName = "" Then
       Exit For
       End If
Next
nofile:
K = 0
'Form1.Caption = Err.Description
Exit Sub
End Sub
Public Sub MonsterAI() 'makes monster move plus attack
Dim DamageNumber As Integer
Dim K As Integer
Call MonsterWarp
If MapEd = True Then Exit Sub
'if the character can see the monster then the monster attacks
If Form1.Image1.Left >= Form1.AttackShape.Left And Form1.Image1.Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And Form1.Image1.Top >= Form1.AttackShape.Top And Form1.Image1.Top <= Form1.AttackShape.Top + Form1.AttackShape.Width Then
If MainPlayer.Stealth = True Then 'if your in stealth make sure the monster doesn't see you
Else
K = 0
For K = K + 1 To 250
If Monster(K).Health = 0 Then
 K = K + 1
End If
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51

If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width And .Visible = True And .Tag = "" Then

MonsterAttacking = True
Monster(K).MonsterAttack = True
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
'READ THIS FIRST
'Below this code stop monsters from
'walking over walls but there is no pathfinding
'code so player could just cast spells on them
If Monster(K).CurX > MainPlayer.xpos Then '
If Map(Monster(K).CurX - 1, Monster(K).CurY).Tiletype = 0 Or Map(Monster(K).CurX - 1, Monster(K).CurY).Tiletype = 11 Then
Call Pathfinding(K, "right")
Else

.Left = .Left - 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
End If
End If
If Monster(K).CurX < MainPlayer.xpos Then
If Map(Monster(K).CurX + 1, Monster(K).CurY).Tiletype = 0 Or Map(Monster(K).CurX + 1, Monster(K).CurY).Tiletype = 11 Then
Call Pathfinding(K, "left")
Else
.Left = .Left + 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
 Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
End If
End If
If Monster(K).CurY > MainPlayer.ypos Then
If Map(Monster(K).CurX, Monster(K).CurY - 1).Tiletype = 0 Or Map(Monster(K).CurX, Monster(K).CurY - 1).Tiletype = 11 Then
Call Pathfinding(K, "down")
Else
.Top = .Top - 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
 Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
End If
End If
If Monster(K).CurY < MainPlayer.ypos Then
If Map(Monster(K).CurX, Monster(K).CurY + 1).Tiletype = 0 Or Map(Monster(K).CurX, Monster(K).CurY + 1).Tiletype = 11 Then
Call Pathfinding(K, "up")
Else
.Top = .Top + 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
 Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
End If
End If

If Monster(K).CurX = MainPlayer.xpos And Monster(K).CurY = MainPlayer.ypos Then
If Monster(K).Stamina >= 0 Then
If Monster(K).MonsterAttack2 > PlayerSkills.MeleeDefense + MainPlayer.Armor Or MainPlayer.Stamina <= 0 Then
Monster(K).Stamina = Monster(K).Stamina - 1
DamageNumber = Monster(K).MonsterAttack2 - PlayerSkills.MeleeDefense + MainPlayer.Armor

If DamageNumber <= 0 Then
DamageNumber = 0
End If
Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True
Form1.lblAttack = Monster(K).PlayerName & " has hit you for " & DamageNumber
Form1.TimerAttacklbl.Enabled = True
MainPlayer.Health = MainPlayer.Health - DamageNumber
Monster(K).Stamina = Monster(K).Stamina - 1
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
TOOLS.LoadStats
Call CheckDead
PlayerSkills.MeeleNextLevel = PlayerSkills.MeeleNextLevel - Monster(K).Level
Call RPG.CheckIfStatLeveled

Else

Monster(K).Stamina = Monster(K).Stamina - 1
MainPlayer.Stamina = MainPlayer.Stamina - 1
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
TOOLS.LoadStats
Form1.MediaPlayer2.Filename = App.Path & "\evade.wav"
Form1.lblAttack.Visible = True
Form1.lblAttack = "You evaded " & Monster(K).PlayerName & " attack"
Form1.TimerAttacklbl.Enabled = True
PlayerSkills.MeeleNextLevel = PlayerSkills.MeeleNextLevel - Monster(K).Level

Call RPG.CheckIfStatLeveled
Call Form1.AddPlayerStats("Xp", Monster(K).Level)
End If
End If
End If

Else
MonsterAttacking = False
Monster(K).MonsterAttack = False
End If
End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
nofile:

Exit Sub
End If
End If

End Sub
Public Sub MonsterStealth()
'Makes a monster use stealth
Dim K As Integer 'monsternumber
K = 0
For K = K + 1 To 250
If Monster(K).Health = 0 Then
 K = K + 1
End If
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
If Monster(K).PlayerName = "Stalker" Then


.Visible = False
Monster(K).PlayerVisible = False
End If

End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next

nofile:

Exit Sub

End Sub
Public Sub MonsterScan()
'monster scans for player
Dim K As Integer 'monsternumber
Call Monsters.BumpIntoMonster
If MainPlayer.Stealth = True Then
K = 0
For K = K + 1 To 250
If Monster(K).Health = 0 Then
 K = K + 1
End If
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
    If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width And .Visible = True Then
        If Monster(K).Scan >= PlayerSkills.Stealth Then
            Form1.lblAttack.Visible = True
            Form1.lblAttack = Monster(K).PlayerName & " has spotted you!"
            Form1.TimerAttacklbl.Enabled = True
            MainPlayer.Stealth = False
            TOOLS.CmdStealth.FontBold = False
        Else

            Form1.lblAttack.Visible = True
            Form1.lblAttack = Monster(K).PlayerName & " Fails to spot you"
            Form1.TimerAttacklbl.Enabled = True
            PlayerSkills.StealthNextLevel = PlayerSkills.StealthNextLevel - Monster(K).Level
            Call RPG.CheckIfStatLeveled
        End If

End If

End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
End If
nofile:

Exit Sub

End Sub
Public Sub MonsterCastSpell()
Dim DamageNumber As Integer
'Randomize
'DamageNumber = (Rnd * 17)
Dim K As Integer 'monsternumber
If MapEd = True Then Exit Sub
'if the character can see the monster then the monster attacks
If Form1.Image1.Left >= Form1.AttackShape.Left And Form1.Image1.Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And Form1.Image1.Top >= Form1.AttackShape.Top And Form1.Image1.Top <= Form1.AttackShape.Top + Form1.AttackShape.Width Then
If MainPlayer.Stealth = True Then 'if your in stealth make sure the monster doesn't see you
Else

K = 0
For K = K + 1 To 250
If Monster(K).Health = 0 Then
 K = K + 1
End If
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
'if the monster is the shape then it casts spells on  you
If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width And .Visible = True Then
'now check if the player resists the spell
If Monster(K).Mana >= 15 Then
Monster(K).Mana = Monster(K).Mana - 15
If PlayerSkills.MagicDefense >= Monster(K).CastSpell Then
MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
Call Form1.AddPlayerStats("Xp", Monster(K).Level)

Form1.lblAttack.Visible = True

Form1.lblAttack = " You resist " & Monster(K).PlayerName & " harm spell!"
Form1.TimerAttacklbl.Enabled = True
PlayerSkills.MagicDefenseNextLevel = PlayerSkills.MagicDefenseNextLevel - Monster(K).Level
Call RPG.CheckIfStatLeveled

Else
DamageNumber = Monster(K).CastSpell - PlayerSkills.MagicDefense
If DamageNumber <= 0 Then
DamageNumber = 0
End If
If DamageNumber > 16 Then
DamageNumber = 16
End If


MainPlayer.XP = MainPlayer.XP + Monster(K).Level
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - Monster(K).Level
Call Form1.AddPlayerStats("Xp", Monster(K).Level)

Form1.MediaPlayer2.Filename = App.Path & "\hit.wav"
Form1.lblAttack.Visible = True

Form1.lblAttack = Monster(K).PlayerName & " harms you for " & DamageNumber
MainPlayer.Health = MainPlayer.Health - DamageNumber
Call CheckDead
Form1.TimerAttacklbl.Enabled = True
PlayerSkills.MagicDefenseNextLevel = PlayerSkills.MagicDefenseNextLevel - Monster(K).Level
Call RPG.CheckIfStatLeveled
End If
TOOLS.LoadStats
End If
End If
End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
nofile:

Exit Sub


End If
End If

End Sub
Public Sub ScanForMonsters()
Dim K As Integer 'monsternumber
K = 0
For K = K + 1 To 250
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
If Monster(K).MapName = "nowhere" Then
Exit Sub
End If
If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width Then
If Monster(K).PlayerVisible = False Then
If PlayerSkills.Scan >= Monster(K).Stealth Then
    Monster(K).PlayerVisible = False
    Form1.lblAttack.Visible = True
    Form1.lblAttack = "You have spotted " & Monster(K).PlayerName
    Form1.TimerAttacklbl.Enabled = True
Else
    Form1.lblAttack.Visible = True
    Form1.lblAttack = "You find no one"
    Form1.TimerAttacklbl.Enabled = True
End If
Else
    Form1.lblAttack.Visible = True
    Form1.lblAttack = "You find no one"
    Form1.TimerAttacklbl.Enabled = True
End If
Else
    Form1.lblAttack.Visible = True
    Form1.lblAttack = "You find no one"
    Form1.TimerAttacklbl.Enabled = True
End If
End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
nofile:

Exit Sub

End Sub
Private Sub MonsterWarp()
Dim K As Integer 'monsternumber
K = 0
For K = K + 1 To 250
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
If Monster(K).MapName = "nowhere" Then
Exit Sub
End If
If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width Then
If Monster(K).PlayerVisible = True Then
If Monster(K).PlayerName = "DarkDrake" Then
 Monster(K).CurX = MainPlayer.xpos
 Monster(K).CurY = MainPlayer.ypos
 Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
 Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
 .Left = Form1.Image1.Left
 .Top = Form1.Image1.Top
 
 
Form1.MediaPlayer2.Filename = "warp.wav"
End If
End If
End If
End With

If Monster(K).MaxHealth = 0 Then
       Exit For
       End If

Next

nofile:

Exit Sub

End Sub
Public Sub MonsterReturn()
K = 0
For K = K + 1 To 250
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51

If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width Then

If Monster(K).MonsterAttack = True Then
    Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
    Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
If Monster(K).CurX > Monster(K).xpos Then
.Left = .Left - 32
End If
If Monster(K).CurX < Monster(K).xpos Then
.Left = .Left + 32

End If

If Monster(K).CurY > Monster(K).ypos Then
.Top = .Top - 32
End If
If Monster(K).CurY < Monster(K).ypos Then
.Top = .Top + 32

End If

End If

Else
Monster(K).MonsterAttack = False

End If

End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
nofile:

Exit Sub

End Sub
Private Sub UpdateWeaponskill(Level As Integer)
'checks if skills leved

If MainPlayer.CurrentWeaponSkill = "axe" Then
PlayerSkills.AxeNextLevel = PlayerSkills.AxeNextLevel - Level
End If
If MainPlayer.CurrentWeaponSkill = "bow" Then
PlayerSkills.BowNextLevel = PlayerSkills.BowNextLevel - Level
End If
If MainPlayer.CurrentWeaponSkill = "bombs" Then
PlayerSkills.BombsNextLevel = PlayerSkills.BombsNextLevel - Level
End If
If MainPlayer.CurrentWeaponSkill = "fail" Then
PlayerSkills.FailNextLvl = PlayerSkills.FailNextLvl - Level
End If
If MainPlayer.CurrentWeaponSkill = "mace" Then
PlayerSkills.MaceNextLevel = PlayerSkills.MaceNextLevel - Level
End If
If MainPlayer.CurrentWeaponSkill = "dagger" Then
PlayerSkills.DaggerNxtLvl = PlayerSkills.DaggerNxtLvl - Level
End If
If MainPlayer.CurrentWeaponSkill = "spear" Then
PlayerSkills.SpearNextLvl = PlayerSkills.SpearNextLvl - Level
End If
If MainPlayer.CurrentWeaponSkill = "staff" Then
PlayerSkills.StaffNxtLvl = PlayerSkills.StaffNxtLvl - Level
End If
If MainPlayer.CurrentWeaponSkill = "force" Then
PlayerSkills.theforceNextLvl = PlayerSkills.theforceNextLvl - Level
End If
If MainPlayer.CurrentWeaponSkill = "unarmed" Then
PlayerSkills.UnarmedAttackNextLevel = PlayerSkills.UnarmedAttackNextLevel - Level
End If
If MainPlayer.CurrentWeaponSkill = "sword" Then
PlayerSkills.SwordNextLvl = PlayerSkills.SwordNextLvl - Level
End If

Call RPG.CheckIfStatLeveled
End Sub
Public Sub BumpIntoMonster()
'Checks if play walks into monsters
Dim K As Integer
If MainPlayer.Stealth = True Then
K = 0
For K = K + 1 To 250

On Error GoTo nofile:

    If MainPlayer.xpos = Monster(K).CurX And MainPlayer.ypos = Monster(K).CurY Then
        MainPlayer.Stealth = False
        Form1.lblAttack.Visible = True
        Form1.lblAttack = "You have bumped into " & Monster(K).PlayerName
        Form1.TimerAttacklbl.Enabled = True
        TOOLS.CmdStealth.FontBold = False

    End If
    If Monster(K).MaxHealth = 0 Then
       Exit For
    End If
Next
End If
nofile:

Exit Sub
End Sub
Public Sub MonstersDrop(Level As Integer)
'You can also add dropping of items
MainPlayer.Money = MainPlayer.Money + Level
TOOLS.LoadStats

End Sub
Private Sub MonQuest(K As Integer)
'Sets Rewards for Monster Quest
Dim Reward As Integer
Randomize
Reward = Rnd * MaxItems
If Monster(K).MonsterIsQuest = 1 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You did it!!!  You have completed your quest!" & vbCrLf
PXItem(Reward).Quanity = PXItem(Reward).Quanity + 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "Your reward is " & PXItem(Reward).ItemName & vbCrLf

Monster(K).MapName = "nowhere"
Form1.ShowMonsters
End If

End Sub

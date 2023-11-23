Attribute VB_Name = "Stuff"
'glFirstLocalPort is the first port to be used by the host computer.
'The host computer will use all ports from glFirstLocalPort to (glFirstLocalPort + giMaxConnections - 1).
Global glFirstLocalPort As Long

'glFirstRemotePort is the first port to be used by connecting computers.
'The connecting computers will use all ports from glFirstRemotePort to (glFirstRemotePort + giMaxConnections - 1).
Global glFirstRemotePort As Long

'giMaxConnections defines the maximum number of connections that can be made to the server.
'Problems occur when using larger numbers.  I would not recommend using more than 16.
Global giMaxConnections As Integer

Global zz As Integer
'API calls used for reading and writing of preferences
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Type MainPlayerType 'your main Player
    PlayerName As String
    ClanName As String
    Admin As Boolean
    CurrentWeaponSkill As String
    MinDamge As Integer
    MaxDamge As Integer
    Cheater As Boolean
    Level As Integer
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
    Food As Integer
    SkillsPercent As Integer
    Money As Long
    KILLS As Integer
    Deaths As Integer
    SkillPoints As Integer
    ServerIP As String
    Burden As Long
    SpendXp As Long
    Armor As Integer
    AccountName As String
    Password As String
    MakeFolder As Boolean
    XP As Long
    Xp4nextLevel As Long
    CurrentWeapon As String
    BankAccountMoney As Long
    AttributePoints As Long
    WeaponSpeed As Integer
End Type

Global MainPlayer As MainPlayerType


Public Sub Fishing()
Dim Normal As Integer, Big As Integer, Small As Integer
Dim fish As Integer


Randomize
fish = (Rnd * 21)
'Find the item numbers of the fish
I = 0
For I = I + 1 To MaxItems
If PXItem(I).ItemName = "SmallFish" Then
Small = I
End If
If PXItem(I).ItemName = "BigFish" Then
Big = I
End If
If PXItem(I).ItemName = "NormalFish" Then
Normal = I
End If
Next I

If fish = 11 Then
PXItem(Big).Quanity = PXItem(Big).Quanity + 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You caught a huge fish!" & vbNewLine

End If

If fish = 14 Or fish = 15 Or fish = 16 Then
PXItem(Normal).Quanity = PXItem(Normal).Quanity + 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You caught a normal fish!" & vbNewLine

End If
If fish = 7 Or fish = 8 Or fish = 9 Or fish = 10 Or fish = 2 Then
PXItem(Small).Quanity = PXItem(Small).Quanity + 1
Form1.txtDialog.Text = Form1.txtDialog.Text & "You caught a small fish!" & vbNewLine

End If

If fish = 0 Or fish = 1 Or fish = 3 Or fish = 4 Or fish = 5 Or fish = 6 Or fish = 13 Or fish = 12 Or fish = 17 Or fish = 18 Or fish = 19 Or fish = 20 Or fish = 21 Then

Form1.txtDialog.Text = Form1.txtDialog.Text & "You failed to catch a fish" & vbNewLine

End If
MainPlayer.XP = MainPlayer.XP + 10
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 10
TOOLS.LoadStats
Form1.AddPlayerStats "Xp", 10
End Sub
Public Sub Mineforgold()
Dim mine As Integer
Randomize
mine = (Rnd * 99)

If mine < 35 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You failed to mine gold" & vbNewLine
End If
If mine > 35 And mine < 55 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You mined 1 gold" & vbNewLine
MainPlayer.Money = MainPlayer.Money + 1
End If
If mine > 55 And mine < 65 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You mined 5 gold" & vbNewLine
MainPlayer.Money = MainPlayer.Money + 5
End If
If mine > 65 And mine < 75 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You mined 10 gold" & vbNewLine
MainPlayer.Money = MainPlayer.Money + 10
End If
If mine > 75 And mine < 85 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You mined 20 gold" & vbNewLine
MainPlayer.Money = MainPlayer.Money + 20
End If
If mine > 85 And mine < 95 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You mined 50 gold" & vbNewLine
MainPlayer.Money = MainPlayer.Money + 50
End If
If mine > 95 And mine < 99 Then
Form1.txtDialog.Text = Form1.txtDialog.Text & "You mined 100 gold" & vbNewLine
MainPlayer.Money = MainPlayer.Money + 100
End If
MainPlayer.XP = MainPlayer.XP + 10
MainPlayer.Xp4nextLevel = MainPlayer.Xp4nextLevel - 10
TOOLS.LoadStats
Form1.AddPlayerStats "Xp", 10
End Sub

Public Sub createNpc(X As Single, y As Single, K As Integer)
Dim C As Integer
Static I As Integer


   If NPc(K).PlayerName = "" Then
   
    Open App.Path & "\monsters\" & AddMonster.File2.Filename For Input As #1
    Line Input #1, textline1 'name
    Line Input #1, textline2 'damage
    Line Input #1, textline3 'speed
    Line Input #1, textline4 'monstermove
    Line Input #1, textline5 'weapon
    Line Input #1, textline6 'armor

    Close #1

    zz = zz + 1
      NPc(K).PlayerName = textline1
      NPc(K).NpcText = AddMonster.Text1.Text
      NPc(K).NpcItem = textline3

      NPc(K).xpos = AddMonster.X.Text
      NPc(K).ypos = AddMonster.y.Text
   'Set P51 = Form1.Controls.Add("VB.Image", "NpcImage" & zz)
C = Form1.PICnpc.UBound + 1
Load Form1.PICnpc(C)
   'With P51
      Form1.PICnpc(C).Enabled = True
      Form1.PICnpc(C).Top = AddMonster.y.Text
      Form1.PICnpc(C).Left = AddMonster.X.Text
      Form1.PICnpc(C).Height = AddMonster.Picture3.Height
      Form1.PICnpc(C).Width = AddMonster.Picture3.Width
      Form1.PICnpc(C).Picture = LoadPicture(App.Path & "\monsters\" & NPc(K).PlayerName & ".gif")
      Form1.PICnpc(C).Visible = True
      Form1.PICnpc(C).ToolTipText = NPc(K).PlayerName

  ' End With
     
   NPc(K).ImageName = C 'zz
   NPc(K).MapName = MainPlayer.MapName
    End If
    
End Sub

Public Sub createMonster(X As Single, y As Single, K As Integer)

   If Monster(K).PlayerName = "" Then
   
    Open App.Path & "\monsters\" & AddMonster.File1.Filename For Input As #1
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
    Monster(K).xpos = AddMonster.X.Text
    Monster(K).ypos = AddMonster.y.Text
   Set P51 = Form1.Controls.Add("VB.Image", "MonsterImage" & Z)

   With P51
      .Enabled = True
      .Top = AddMonster.y.Text
      .Left = AddMonster.X.Text
      .Height = AddMonster.Picture1.Height
      .Width = AddMonster.Picture1.Width
      .Picture = LoadPicture(App.Path & "\monsters\" & Monster(K).PlayerName & ".gif")
      .Visible = True
      .ToolTipText = Monster(K).PlayerName & " Level " & Monster(K).Level & " Health " & Monster(K).Health & "\" & Monster(K).MaxHealth
     ' Form1.Caption = .Name & "mon"
   End With
     
'   Form1.p21(z).Top = y
'   Form1.p21(z).Left = X
   Monster(K).ImageName = Z
   Monster(K).MapName = MainPlayer.MapName
'  Monster(K).CurX = (Snap(Form1.p21(Z).Left, 32) \ 32) + HScroll
' Monster(K).CurY = (Snap(Form1.p21(Z).Top, 32) \ 32) + VScroll
Call MovenewMonster2
   ' Form1.p21(z).Visible = True
  ' Call MovenewMonsterk, z)
    'If Quest = 1 Then
   ' Monster(K).MonsterIsQuest = 1
   ' End If
    
   
   ' Else
    'Creates a monster after the last monster made
    
   'For K = K + 1 To 200
   '' If Monster(K).PlayerName = "" Then
  '  Call createMonster(X, y, K)
   ' Exit For
   ' End If
    
     End If
    
End Sub

Public Sub MovenewMonster2()

Dim newmon As String
Dim K As Integer
K = 0

For K = K + 1 To 250
If Monster(K).MonsterAttack = True Then
 K = K + 1
End If
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51

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
End With
If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
Next
nofile:

Exit Sub
End Sub
Public Sub MovenewNpc()
Dim newnpc As String
Dim K As Integer
Dim C As Integer
K = 0

For K = K + 1 To 250
'newnpc = "NpcImage" & k
On Error GoTo nofile:
 ' Set P51 = Form1.Controls(newnpc)
'With P51

With Form1.PICnpc(K)
NPc(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
NPc(K).CurY = (Snap(.Top, 32) \ 32) + VScroll


If NPc(K).CurX > NPc(K).xpos Then
.Left = .Left - 32
End If
If NPc(K).CurX < NPc(K).xpos Then
.Left = .Left + 32

End If

If NPc(K).CurY > NPc(K).ypos Then
.Top = .Top - 32
End If
If NPc(K).CurY < NPc(K).ypos Then
.Top = .Top + 32

End If
End With
If NPc(K).PlayerName = "" Then
       Exit For
       End If
Next
nofile:

Exit Sub
End Sub
Public Sub MoveAttack()
Dim newmon As String
Dim K As Integer
K = 0
For K = K + 1 To 250
newmon = "MonsterImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width Then
If Monster(K).OrgX = "" Then
    Monster(K).OrgX = Monster(K).xpos
    Monster(K).OrgY = Monster(K).ypos
End If

 
    Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
    Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
If MainPlayer.xpos > Monster(K).CurX Then
    Monster(K).xpos = Monster(K).xpos + 1
End If
If MainPlayer.xpos < Monster(K).CurX Then
    Monster(K).xpos = Monster(K).xpos - 1
End If
If MainPlayer.ypos > Monster(K).CurY Then
    Monster(K).ypos = Monster(K).ypos + 1
End If
If MainPlayer.ypos < Monster(K).CurY Then
    Monster(K).ypos = Monster(K).ypos - 1
End If
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

If Monster(K).MaxHealth = 0 Then
       Exit For
       End If
       
       End If
 End With
Next
nofile:

Exit Sub
End Sub


'Npc Speech Functions

Public Sub InitSpeech(ArrayIndex As Integer)
TalkingTo = ArrayIndex
Dim tInt As Integer
If Not NPCTalk(ArrayIndex).AlreadyTalkedTo = True Then
    NPCTalk(ArrayIndex).AlreadyTalkedTo = True
    tInt = NPCTalk(ArrayIndex).SIndex
Else: tInt = NPCTalk(ArrayIndex).ATIndex
End If
CSpeech = tInt
Form1.Pitanje.Caption = NPCTalk(ArrayIndex).NPCName & " : " & vbCrLf & "::: " & Speech(tInt).Pitanje
Form1.Choice(0).Top = Form1.Pitanje.Top + Form1.Pitanje.Height + 60
Form1.Choice(0).Caption = ">>> " & Speech(tInt).Odgovor(0)
For I = 1 To Speech(tInt).nOdgovor - 1
    Load Form1.Choice(Form1.Choice.Count)
    Form1.Choice(Form1.Choice.Count - 1).Left = 120
    Form1.Choice(Form1.Choice.Count - 1).Top = Form1.Choice(Form1.Choice.Count - 2).Top + Form1.Choice(Form1.Choice.Count - 2).Height + 60
    Form1.Choice(Form1.Choice.Count - 1).Caption = ">>> " & Speech(tInt).Odgovor(I)
    Form1.Choice(Form1.Choice.Count - 1).AutoSize = True
    Form1.Choice(Form1.Choice.Count - 1).Visible = True
Next
Form1.PicTXT.Visible = True
UpdateScrollBar
End Sub

Public Sub CloseSpeech()
Form1.PicTXT.Visible = False


For I = 0 To Form1.Choice.Count - 2
    Unload Form1.Choice(Form1.Choice.Count - 1)
Next
End Sub

Public Sub RedirectSpeech(Index As Integer, NPCIndex As Integer)
Form1.TextScroll.Value = 0
If Index = 0 Then
    CloseSpeech
    Exit Sub
End If
For I = 0 To Form1.Choice.Count - 2
    Unload Form1.Choice(Form1.Choice.Count - 1)
Next
tInt = Index
CSpeech = tInt
Form1.Pitanje.Caption = NPCTalk(NPCIndex).NPCName & " : " & vbCrLf & "::: " & Speech(tInt).Pitanje
Form1.Choice(0).Top = Form1.Pitanje.Top + Form1.Pitanje.Height + 60
Form1.Choice(0).Caption = ">>> " & Speech(tInt).Odgovor(0)
For I = 1 To Speech(tInt).nOdgovor - 1
    Load Form1.Choice(Form1.Choice.Count)
    Form1.Choice(Form1.Choice.Count - 1).Left = 120
    Form1.Choice(Form1.Choice.Count - 1).Top = Form1.Choice(Form1.Choice.Count - 2).Top + Form1.Choice(Form1.Choice.Count - 2).Height + 60
   Form1.Choice(Form1.Choice.Count - 1).Caption = ">>> " & Speech(tInt).Odgovor(I)
    Form1.Choice(Form1.Choice.Count - 1).AutoSize = True
    Form1.Choice(Form1.Choice.Count - 1).Visible = True
Next
UpdateScrollBar
End Sub

Public Sub UpdateScrollBar()
Dim tInt As Integer
tInt = Form1.Choice(Form1.Choice.Count - 1).Top - Form1.Choice(Form1.Choice.Count - 1).Height
Form1.TextScroll.Max = tInt / (Form1.PicTXT.Height / 2)
Form1.PicTXT.SetFocus
End Sub
'## end speech functions
Public Sub CenterPerson() 'Centers The Character
'Now improved so it works faster
Dim A As Integer, B As Integer, C As Integer, d As Integer, counter4 As Integer
A = 0
B = 0
C = 0
d = 0
counter4 = 0
Warping = True
Do Until d = 1 And C = 1 Or A = 1 And d = 1 Or B = 1 And C = 1 Or B = 1 And d = 1 '  And c = 1 Or d = 1
MainPlayer.xpos = (Snap(Form1.Image1.Left, 32) \ 32) + HScroll
MainPlayer.ypos = (Snap(Form1.Image1.Top, 32) \ 32) + VScroll
cx = MainPlayer.xpos
cy = MainPlayer.ypos
If Form1.Image1.Top < 256 Then
If Form1.VScroll1.Value < Form1.VScroll1.Max Then
If Form1.VScroll1.Value > Form1.VScroll1.Min Then
If Form1.Image1.Top + 32 > 256 Then A = 1
    Form1.Image1.Top = Form1.Image1.Top + 32
    Form1.VScroll1.Value = Form1.VScroll1.Value - 1
End If
'Form1.RedrawMap
End If
'a = 1
End If
If Form1.Image1.Top > 256 Then
If Form1.VScroll1.Value > Form1.VScroll1.Min Then

If Form1.VScroll1.Value < Form1.VScroll1.Max Then
    If Form1.Image1.Top - 32 < 256 Then B = 1
    Form1.Image1.Top = Form1.Image1.Top - 32
    Form1.VScroll1.Value = Form1.VScroll1.Value + 1
End If
'Form1.RedrawMap
End If
Else
'B = 1
End If
If Form1.Image1.Left < 264 Then
If Form1.Image1.Left = 264 Or Form1.Image1.Left = 256 Or Form1.Image1.Left = 240 Then Exit Sub
If Form1.HScroll1.Value < Form1.HScroll1.Max Then


If Form1.HScroll1.Value > Form1.HScroll1.Min Then
If Form1.Image1.Left + 32 > 264 Then C = 1
    Form1.Image1.Left = Form1.Image1.Left + 32
    Form1.HScroll1.Value = Form1.HScroll1.Value - 1
   ' Form1.RedrawMap
End If
End If
Else

End If
If Form1.Image1.Left > 264 Then
If Form1.Image1.Left = 264 Or Form1.Image1.Left = 256 Or Form1.Image1.Left = 240 Then Exit Sub
If Form1.HScroll1.Value > Form1.HScroll1.Min Then

If Form1.HScroll1.Value < Form1.HScroll1.Max Then
If Form1.Image1.Left - 32 < 264 Then d = 1
    Form1.Image1.Left = Form1.Image1.Left - 32
    Form1.HScroll1.Value = Form1.HScroll1.Value + 1
   ' Form1.RedrawMap
End If
Form1.lblPlayerName(0).Top = Form1.Image1.Top - 20
Form1.lblPlayerName(0).Left = Form1.Image1.Left
End If
Else

End If

MainPlayer.xpos = (Snap(Form1.Image1.Left, 32) \ 32) + HScroll
MainPlayer.ypos = (Snap(Form1.Image1.Top, 32) \ 32) + VScroll
If cx = MainPlayer.xpos And cy = MainPlayer.ypos Then
counter4 = counter4 + 1
If counter4 > 5 Then
A = 1
B = 1
C = 1
d = 1
End If
End If
Loop
Warping = False
Form1.RedrawMap
End Sub
Public Sub MoreEvents() 'For the town in 7.4
'I would convert this information to a mapfile
'I didn't do this when i start due i could not save extra data
If MainPlayer.MapName = "7.4.map" Then
If MainPlayer.xpos = 39 And MainPlayer.ypos = 47 Then
  Open App.Path & "\maps\motel.map" For Input As #1: Close
Open App.Path & "\maps\motel.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\htiles.bmp")

MainPlayer.MapName = "motel.map"
Form1.MapName = "motel.map"
Form1.RedrawMap
Form1.Image1.Left = 288
'
Form1.Image1.Top = 256
Form1.VScroll1.Value = 6
Form1.HScroll1.Value = 0
End If
If MainPlayer.xpos = 48 And MainPlayer.ypos = 55 Then
  Open App.Path & "\maps\house2.map" For Input As #1: Close
Open App.Path & "\maps\house2.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\htiles.bmp")

MainPlayer.MapName = "house2.map"
Form1.MapName = "house2.map"
Form1.RedrawMap
Form1.Image1.Left = 296
'
Form1.Image1.Top = 352
Form1.VScroll1.Value = 1
Form1.HScroll1.Value = 0
End If
If MainPlayer.xpos = 28 And MainPlayer.ypos = 53 Then
  Open App.Path & "\maps\house3.map" For Input As #1: Close
Open App.Path & "\maps\house3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\htiles.bmp")

MainPlayer.MapName = "house3.map"
Form1.MapName = "house3.map"
Form1.RedrawMap
Form1.Image1.Left = 292
'
Form1.Image1.Top = 352
Form1.VScroll1.Value = 1
Form1.HScroll1.Value = 2
End If

End If
If MainPlayer.MapName = "motel.map" Then
If MainPlayer.xpos = 8 And MainPlayer.ypos = 15 Then
  Open App.Path & "\maps\7.4.map" For Input As #1: Close
Open App.Path & "\maps\7.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\tiles.bmp")

MainPlayer.MapName = "7.4.map"
Form1.MapName = "7.4.map"
Form1.RedrawMap
Form1.Image1.Left = 296
'
Form1.Image1.Top = 304
Form1.VScroll1.Value = 40
Form1.HScroll1.Value = 31
End If
End If

If MainPlayer.MapName = "house2.map" Then
If MainPlayer.xpos = 10 And MainPlayer.ypos = 14 Then
  Open App.Path & "\maps\7.4.map" For Input As #1: Close
Open App.Path & "\maps\7.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\tiles.bmp")

MainPlayer.MapName = "7.4.map"
Form1.MapName = "7.4.map"
Form1.RedrawMap

Form1.VScroll1.Value = 44
Form1.HScroll1.Value = 38
End If
End If

If MainPlayer.MapName = "house4.map" Then
If MainPlayer.xpos = 11 And MainPlayer.ypos = 6 Then
  Open App.Path & "\maps\house3.map" For Input As #1: Close
Open App.Path & "\maps\house3.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close


MainPlayer.MapName = "house3.map"
Form1.MapName = "house3.map"
Form1.RedrawMap

Form1.Image1.Top = 224
Form1.VScroll1.Value = 0
Form1.HScroll1.Value = 0
Exit Sub
End If
End If
If MainPlayer.MapName = "house3.map" Then
If MainPlayer.xpos = 14 And MainPlayer.ypos = 14 Then
  Open App.Path & "\maps\7.4.map" For Input As #1: Close
Open App.Path & "\maps\7.4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\tiles.bmp")

MainPlayer.MapName = "7.4.map"
Form1.MapName = "7.4.map"
Form1.RedrawMap
Form1.VScroll1.Value = 44
Form1.HScroll1.Value = 16
Exit Sub
End If
If MainPlayer.xpos = 11 And MainPlayer.ypos = 6 Then
  Open App.Path & "\maps\house4.map" For Input As #1: Close
Open App.Path & "\maps\house4.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close

MainPlayer.MapName = "house4.map"
Form1.MapName = "house4.map"
Form1.RedrawMap
Form1.Image1.Top = 256
Form1.VScroll1.Value = 0
Form1.HScroll1.Value = 0
End If


If MainPlayer.xpos = 14 And MainPlayer.ypos = 6 Then
  ExitMap = MainPlayer.MapName
  ExitHscroll = Form1.HScroll1.Value
  ExitVscroll = Form1.VScroll1.Value
  ExitTop = Form1.Image1.Top
  ExitLeft = Form1.Image1.Left
  ExitTiles = "htiles.bmp"
  ExitDirection = Form1.Direction.Caption
Open App.Path & "\maps\c1.map" For Input As #1: Close
Open App.Path & "\maps\c1.map" For Binary Access Read Lock Read As #1
Form1.EraseAll
Get #1, , Map
Close
Form1.Tiles.Picture = LoadPicture(App.Path & "\tiles.bmp")


MainPlayer.MapName = "c1.map"
Form1.MapName = "c1.map"
Form1.RedrawMap
Form1.Image1.Top = 256
Form1.VScroll1.Value = 9
Form1.HScroll1.Value = 10
End If

End If

End Sub
Public Sub CreateQuestMonster(X As Single, y As Single, K As Integer, Quest As Integer, Mapx As Integer, Mapy As Integer)
Dim ChoseMon As Integer
Dim MonsterFile As String
'It chooses a monster at random to fight against
Randomize
ChoseMon = Rnd * 10
If ChoseMon = 0 Then
MonsterFile = "Chicken.bad"
End If
If ChoseMon = 1 Then
MonsterFile = "blob.bad"
End If
If ChoseMon = 2 Then
MonsterFile = "elf.bad"
End If
If ChoseMon = 3 Then
MonsterFile = "evil santa.bad"
End If
If ChoseMon = 4 Then
MonsterFile = "twister.bad"
End If
If ChoseMon = 5 Then
MonsterFile = "undead.bad"
End If
If ChoseMon = 6 Then
MonsterFile = "witch.bad"
End If
If ChoseMon = 7 Then
MonsterFile = "warlock.bad"
End If
If ChoseMon = 8 Then
MonsterFile = "Fire Demon.bad"
End If
If ChoseMon = 9 Then
MonsterFile = "dragon.bad"
End If
If ChoseMon = 10 Then
MonsterFile = "darkdrake.bad"
End If
   If Monster(K).PlayerName = "" Then
   
    Open App.Path & "\monsters\" & MonsterFile For Input As #1
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
    Monster(K).xpos = X
    Monster(K).ypos = y
   Set P51 = Form1.Controls.Add("VB.Image", "MonsterImage" & Z)

   Monster(K).ImageName = Z
   Monster(K).MapName = Mapx & "." & Mapy & ".map"
If Mapx = 5 And Mapy = 5 Then
    Monster(K).MapName = "newbie"
End If
    If Quest = 1 Then
        Monster(K).MonsterIsQuest = 1
    End If
   With P51
      .Enabled = True
      .Top = y
      .Left = X
      .Height = 32
      .Width = 32
      .Picture = LoadPicture(App.Path & "\monsters\" & Monster(K).PlayerName & ".gif")
      .Visible = True
      .ToolTipText = Monster(K).PlayerName & " Level " & Monster(K).Level & " Health " & Monster(K).Health & "\" & Monster(K).MaxHealth
   End With
   
Else
    'Creates a monster after the last monster made
    
    For K = K + 1 To 200
        If Monster(K).PlayerName = "" And Quest = 1 Then
            Call CreateQuestMonster(X, y, K, 1, Mapx, Mapy)
            Exit For
        End If
    Next
    
End If
    
End Sub

Function CalculateBurden()
'This function is not used in the game yet but you can use
'it to find the burden on the player with all there items
'You also have to edit the item files
'to include there burden amount or else it will be zero
Dim Burden As Integer
I = 0
For I = I + 1 To MaxItems
    If PXItem(I).Quanity > 1 Then
        Burden = Burden + (PXItem(I).Burden * PXItem(I).Quanity)
    End If
Next I
CalculateBurden = Burden
End Function

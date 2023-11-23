Attribute VB_Name = "ModRndMap"
'This is pretty neat makes random towns and worlds you could use
'this to make a random world each time you play!!

'I am still working on this but call this
'call ModRndMap.RandomMapGenerator("forest", 1000, False, 1, True)
'It's pretty cool!!!!!!!!!!!!!
'to do code random town generator, sewer, and island
Dim Counter As Integer
Dim X As Integer
Dim y As Integer
Public Sub RandomMapGenerator(MapType As String, Trees As Integer, Water As Boolean, Flowers As Integer, TotallyRandom As Boolean)
Dim RanX As Integer
Dim RanY As Integer
Dim RanTreeType As Integer
Randomize
If MapType = "forest" Then
Call LoadGrass

Counter = Trees
   Do While Counter > 0   '  loop.
      Counter = Counter - 1   ' Increment Counter.
     RanTreeType = (Rnd * 2)
     X1 = 23 * mWidth
     Y1 = 16 * mHeight
     RanX = (Rnd * X1)
     RanY = (Rnd * Y1)
       If RanTreeType = 0 Then
        Map(RanX, RanY).TileX = 1440
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = Tree
       End If
       If RanTreeType = 1 Then
        Map(RanX, RanY).TileX = 352
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = Tree
       End If
       If RanTreeType = 2 Then
        Map(RanX, RanY).TileX = 736
        Map(RanX, RanY).Tiley = 64
        Map(RanX, RanY).Tiletype = Tree
       End If
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
'flowers
Counter = Flowers
   Do While Counter > 0
      Counter = Counter - 1
        X1 = 23 * mWidth
        Y1 = 16 * mHeight
        RanX = (Rnd * X1)
        RanY = (Rnd * Y1)
        Map(RanX, RanY).TileX = 352
        Map(RanX, RanY).Tiley = 32
        Map(RanX, RanY).Tiletype = NONWALKABLE
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
Call BigTrees(Trees, "forest")
End If
If MapType = "desert" Then
Call LoadDesert
Counter = Trees
   Do While Counter > 0   '  loop.
      Counter = Counter - 1   ' Increment Counter.
     RanTreeType = (Rnd * 2)
     X1 = 23 * mWidth
     Y1 = 16 * mHeight
     RanX = (Rnd * X1)
     RanY = (Rnd * Y1)
     
        Map(RanX, RanY).TileX = 6976
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = WALKABLE
      If RanTreeType = 2 Then
      If RanX = 92 Then
      Else
        Map(RanX, RanY).TileX = 6912
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = WALKABLE
        Map(RanX + 1, RanY).TileX = 6944
        Map(RanX + 1, RanY).Tiley = 0
        Map(RanX + 1, RanY).Tiletype = WALKABLE
       End If
       End If
    
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
'Loads desert plants
Counter = Flowers
   Do While Counter > 0   '  loop.
      Counter = Counter - 1   ' Increment Counter.
     RanTreeType = (Rnd * 2)
     X1 = 23 * mWidth
     Y1 = 16 * mHeight
     RanX = (Rnd * X1)
     RanY = (Rnd * Y1)
     
       If RanTreeType = 0 Then
        Map(RanX, RanY).TileX = 7840
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = NONWALKABLE
       End If
       If RanTreeType = 1 Then
        Map(RanX, RanY).TileX = 7872
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = NONWALKABLE
       End If
       If RanTreeType = 2 Then
        Map(RanX, RanY).TileX = 768
        Map(RanX, RanY).Tiley = 32
        Map(RanX, RanY).Tiletype = NONWALKABLE
       End If
    
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop

End If

If MapType = "town" Then
Call LoadGrass
'Trees are buildings and flowers are trees Kinda odd isn't
Call TownCreate(Trees, Flowers)
End If

If MapType = "sewer" Then
Call LoadSewer
End If
If MapType = "ocean" Then
Call LoadWater
End If

If MapType = "island" Then
Call LoadWater

End If

If MapType = "snow" Then
Call LoadSnow
'flowers
Counter = Flowers
   Do While Counter > 0
      Counter = Counter - 1
        X1 = 23 * mWidth
        Y1 = 16 * mHeight
        RanX = (Rnd * X1)
        RanY = (Rnd * Y1)
        Map(RanX, RanY).TileX = 896
        Map(RanX, RanY).Tiley = 64
        Map(RanX, RanY).Tiletype = WALKABLE
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
'Loads snow trees
Counter = Trees
   Do While Counter > 0   '  loop.
      Counter = Counter - 1   ' Increment Counter.
     RanTreeType = (Rnd * 2)
     X1 = 23 * mWidth
     Y1 = 16 * mHeight
     RanX = (Rnd * X1)
     RanY = (Rnd * Y1)
       If RanTreeType = 0 Then
        Map(RanX, RanY).TileX = 768
        Map(RanX, RanY).Tiley = 64
        Map(RanX, RanY).Tiletype = Tree
       End If
       If RanTreeType = 1 Then
        Map(RanX, RanY).TileX = 800
        Map(RanX, RanY).Tiley = 64
        Map(RanX, RanY).Tiletype = Tree
       End If
       If RanTreeType = 2 Then
        Map(RanX, RanY).TileX = 1440
        Map(RanX, RanY).Tiley = 64
        Map(RanX, RanY).Tiletype = Tree
       End If
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
End If
If MapType = "cliffs" Then
Dim CliffX As Integer
Call LoadGrass
    CliffX = (Rnd * (16 * mHeight))
    Counter = 23 * mWidth
       Do While Counter > 0
    If CliffX >= 89 Then CliffX = 88
        Map(Counter, CliffX).TileX = 14304
        Map(Counter, CliffX).Tiley = 0
        Map(Counter, CliffX).Tiletype = NONWALKABLE
        Map(Counter, CliffX + 1).TileX = 14304
        Map(Counter, CliffX + 1).Tiley = 32
        Map(Counter, CliffX + 1).Tiletype = NONWALKABLE
        Map(Counter, CliffX + 2).TileX = 14304
        Map(Counter, CliffX + 2).Tiley = 64
        Map(Counter, CliffX + 2).Tiletype = NONWALKABLE
        
       Counter = Counter - 1
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
'flowers
Counter = Flowers
   Do While Counter > 0
      Counter = Counter - 1
       
        X1 = 23 * mWidth
        Y1 = 16 * mHeight
        RanX = (Rnd * X1)
        RanY = (Rnd * Y1)
        If Map(RanX, RanY).Tiletype = NONWALKABLE Or Map(RanX, RanY).Tiletype = Tree Or Map(RanX, RanY).Tiletype = SWIMABLE Then
        Else
        Map(RanX, RanY).TileX = 352
        Map(RanX, RanY).Tiley = 32
        Map(RanX, RanY).Tiletype = NONWALKABLE
        End If
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
Call GrassTrees(Trees)
End If
If MapType = "light forest" Then
Call LoadGrass
Call GrassTrees(Trees)
'flowers
Counter = Flowers
   Do While Counter > 0
      Counter = Counter - 1
        X1 = 23 * mWidth
        Y1 = 16 * mHeight
        RanX = (Rnd * X1)
        RanY = (Rnd * Y1)
        Map(RanX, RanY).TileX = 352
        Map(RanX, RanY).Tiley = 32
        Map(RanX, RanY).Tiletype = NONWALKABLE
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
End If

If MapType = "river" Then
Dim River As Integer
Dim RiverX As Integer
Dim RiverY As Integer
Call LoadGrass
River = (Rnd * 3)
    If River = 1 Then 'north river
    RiverY = (Rnd * (23 * mWidth))
    Counter = 16 * mHeight
       Do While Counter > 0
     
        Map(RiverY, Counter).TileX = 6816
        Map(RiverY, Counter).Tiley = 0
        Map(RiverY, Counter).Tiletype = SWIMABLE
       Counter = Counter - 1
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
    End If
    If River = 2 Then 'west river
    RiverX = (Rnd * (16 * mHeight))
    Counter = 23 * mWidth
       Do While Counter > 0
        Map(Counter, RiverX).TileX = 6592
        Map(Counter, RiverX).Tiley = 32
        Map(Counter, RiverX).Tiletype = SWIMABLE
       Counter = Counter - 1
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
    End If
    If River = 3 Then 'east river
    RiverX = (Rnd * (16 * mHeight))
    Counter = 23 * mWidth
       Do While Counter > 0
    
        Map(Counter, RiverX).TileX = 6592
        Map(Counter, RiverX).Tiley = 32
        Map(Counter, RiverX).Tiletype = SWIMABLE
       Counter = Counter - 1
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
    End If
    If River = 0 Then 'south river
    RiverY = (Rnd * (23 * mWidth))
    Counter = 16 * mHeight
       Do While Counter > 0
     
        Map(RiverY, Counter).TileX = 6816
        Map(RiverY, Counter).Tiley = 0
        Map(RiverY, Counter).Tiletype = SWIMABLE
       Counter = Counter - 1
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
    End If
'flowers
Counter = Flowers
   Do While Counter > 0
      Counter = Counter - 1
       
        X1 = 23 * mWidth
        Y1 = 16 * mHeight
        RanX = (Rnd * X1)
        RanY = (Rnd * Y1)
        If Map(RanX, RanY).Tiletype = NONWALKABLE Or Map(RanX, RanY).Tiletype = Tree Or Map(RanX, RanY).Tiletype = SWIMABLE Then
        Else
        Map(RanX, RanY).TileX = 352
        Map(RanX, RanY).Tiley = 32
        Map(RanX, RanY).Tiletype = NONWALKABLE
        End If
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
Call GrassTrees(Trees)
End If
If MapType = "wlandeast" Then
Dim WPos As Integer

WPos = (Rnd * (23 * mWidth))
If WPos < 10 Then WPos = 10
Call LoadWater
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
    If X >= WPos Then
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 320
        Map(X, y).Tiley = 0
    End If
    Next
Next
Call BigTrees(Trees, "forest")
End If
If MapType = "wlandwest" Then


WPos = (Rnd * (23 * mWidth))
If WPos < 50 Then WPos = 55
Call LoadWater
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
    If X <= WPos Then
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 320
        Map(X, y).Tiley = 0
    End If
    Next
    Next
Call BigTrees(Trees, "forest")
End If
If MapType = "wlandsouth" Then
Call LoadWater
WPos = (Rnd * (16 * mHeight))
If WPos > 50 Then WPos = 55
Call LoadWater
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
    If y >= WPos Then
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 320
        Map(X, y).Tiley = 0
    End If
    Next
    Next
Call BigTrees(Trees, "forest")
End If
If MapType = "wlandnorth" Then
Call LoadWater
WPos = (Rnd * (16 * mHeight))
If WPos < 10 Then WPos = 11
Call LoadWater
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
    If y <= WPos Then
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 320
        Map(X, y).Tiley = 0
    End If
    Next
    Next
Call BigTrees(Trees, "forest")
End If
If MapType = "deadforest" Then
Call LoadGrass
Counter = Trees
   Do While Counter > 0   '  loop.
      Counter = Counter - 1   ' Increment Counter.
     RanTreeType = (Rnd * 2)
     X1 = 23 * mWidth
     Y1 = 16 * mHeight
     RanX = (Rnd * X1)
     RanY = (Rnd * Y1)
     
        Map(RanX, RanY).TileX = 1440
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = Tree
      If RanTreeType = 2 Then
        Map(RanX, RanY).TileX = 8032
        Map(RanX, RanY).Tiley = 64
        Map(RanX, RanY).Tiletype = NONWALKABLE
       End If
    
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop

End If
If MapType = "mine" Then
Call LoadMine
End If
End Sub
Private Sub LoadDesert()
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 256
        Map(X, y).Tiley = 0
    Next
Next
Form1.RedrawMap
End Sub
Private Sub LoadWater()
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        Map(X, y).Tiletype = 2
        Map(X, y).TileX = 6784
        Map(X, y).Tiley = 64
    Next
Next
Form1.RedrawMap
End Sub
Private Sub LoadGrass()
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 320
        Map(X, y).Tiley = 0
    Next
Next
Form1.RedrawMap
End Sub
Private Sub LoadSnow()
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 288
        Map(X, y).Tiley = 0
    Next
Next
Form1.RedrawMap
End Sub
Private Sub LoadSewer()
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        Map(X, y).Tiletype = 1
        Map(X, y).TileX = 11840
        Map(X, y).Tiley = 0
    Next
Next
Form1.RedrawMap
End Sub
Private Sub LoadMine()
For y = 0 To UBound(Map, 2)
    For X = 0 To UBound(Map, 1)
        Map(X, y).Tiletype = Gold
        Map(X, y).TileX = 14048
        Map(X, y).Tiley = 0
    Next
Next
Form1.RedrawMap
End Sub
Private Sub GrassTrees(Trees As Integer)
Dim RanX As Integer
Dim RanY As Integer
Dim RanTreeType As Integer
Randomize
Counter = Trees
   Do While Counter > 0   '  loop.
      Counter = Counter - 1   ' Increment Counter.
     RanTreeType = (Rnd * 2)
     X1 = 23 * mWidth
     Y1 = 16 * mHeight
     RanX = (Rnd * X1)
     RanY = (Rnd * Y1)
      If Map(RanX, RanY).Tiletype = NONWALKABLE Or Map(RanX, RanY).Tiletype = Tree Or Map(RanX, RanY).Tiletype = SWIMABLE Then
        Else
       If RanTreeType = 0 Then
        Map(RanX, RanY).TileX = 1440
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = Tree
       End If
       If RanTreeType = 1 Then
        Map(RanX, RanY).TileX = 352
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = Tree
       End If
       If RanTreeType = 2 Then
        Map(RanX, RanY).TileX = 736
        Map(RanX, RanY).Tiley = 64
        Map(RanX, RanY).Tiletype = Tree
       End If
       End If
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
End Sub
Private Sub BigTrees(Trees As Integer, TreeType As String)
If TreeType = "forest" Then
Dim RanX As Integer
Dim RanY As Integer
Dim RanTreeType As Integer
Randomize
Counter = Trees
   Do While Counter > 0   '  loop.
      Counter = Counter - 1   ' Increment Counter.
    
     X1 = 23 * mWidth
     Y1 = 16 * mHeight
     RanX = (Rnd * X1)
     RanY = (Rnd * Y1)
     If RanY >= 60 Then RanY = 59
    If RanX >= 88 Then RanX = 87
      If Map(RanX, RanY).Tiletype = NONWALKABLE Or Map(RanX, RanY).Tiletype = Tree Or Map(RanX, RanY).Tiletype = SWIMABLE Then
        Else
             If Map(RanX + 1, RanY).Tiletype = NONWALKABLE Or Map(RanX + 1, RanY).Tiletype = Tree Or Map(RanX + 1, RanY).Tiletype = SWIMABLE Then
                Else
                     If Map(RanX, RanY + 1).Tiletype = NONWALKABLE Or Map(RanX, RanY + 1).Tiletype = Tree Or Map(RanX, RanY + 1).Tiletype = SWIMABLE Then
                        Else
                             If Map(RanX + 1, RanY + 1).Tiletype = NONWALKABLE Or Map(RanX + 1, RanY + 1).Tiletype = Tree Or Map(RanX + 1, RanY + 1).Tiletype = SWIMABLE Then
                                Else
        Map(RanX, RanY).TileX = 544
        Map(RanX, RanY).Tiley = 0
        Map(RanX, RanY).Tiletype = NONWALKABLE
        Map(RanX + 1, RanY).TileX = 576
        Map(RanX + 1, RanY).Tiley = 0
        Map(RanX + 1, RanY).Tiletype = NONWALKABLE
        Map(RanX + 1, RanY + 1).TileX = 576
        Map(RanX + 1, RanY + 1).Tiley = 32
        Map(RanX + 1, RanY + 1).Tiletype = NONWALKABLE
        Map(RanX, RanY + 1).TileX = 544
        Map(RanX, RanY + 1).Tiley = 32
        Map(RanX, RanY + 1).Tiletype = NONWALKABLE
       End If
       End If
       End If
       End If
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
Loop
End If
If TreeType = "snow" Then

End If
End Sub
Public Sub Pathfinding(MonsterNumber As Integer, Direction As String)
'Ye old pathingfinding code i never could get this to work
'if you get it to work send it to me
Dim K As Integer
K = MonsterNumber

X = Monster(K).CurX
y = Monster(K).CurY

newmon = "MonsterImage" & K
'On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51

If .Left >= Form1.AttackShape.Left And .Left <= Form1.AttackShape.Left + Form1.AttackShape.Width And .Top >= Form1.AttackShape.Top And .Top <= Form1.AttackShape.Top + Form1.AttackShape.Width And .Visible = True And .Tag = "" Then

Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
If Direction = "left" Then
If Monster(K).CurX > MainPlayer.xpos Then '
If Map(Monster(K).CurX - 1, Monster(K).CurY).Tiletype = 0 Or Map(Monster(K).CurX - 1, Monster(K).CurY).Tiletype = 11 Then
Call CheckRightLeft(K)
Call Pathfinding(K, "right")
Exit Sub
Else

.Left = .Left - 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
End If
End If
End If

If Direction = "right" Then
If Monster(K).CurX < MainPlayer.xpos Then
If Map(Monster(K).CurX + 1, Monster(K).CurY).Tiletype = 0 Or Map(Monster(K).CurX + 1, Monster(K).CurY).Tiletype = 11 Then
Call CheckRightLeft(K)
Call Pathfinding(K, "left")
Exit Sub
Else
.Left = .Left + 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
End If
End If
End If

If Direction = "up" Then
If Monster(K).CurY > MainPlayer.ypos Then
If Map(Monster(K).CurX, Monster(K).CurY + 1).Tiletype = 0 Or Map(Monster(K).CurX, Monster(K).CurY + 1).Tiletype = 11 Then
Call CheckUpDown(K)
Call Pathfinding(K, "down")
Exit Sub
Else
.Top = .Top - 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll

End If
End If
End If

If Direction = "down" Then
If Monster(K).CurY < MainPlayer.ypos Then
If Map(Monster(K).CurX, Monster(K).CurY - 1).Tiletype = 0 Or Map(Monster(K).CurX, Monster(K).CurY - 1).Tiletype = 11 Then
Call CheckUpDown(K)
Call Pathfinding(K, "up")
Exit Sub
Else
.Top = .Top + 32
Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
End If
End If
End If
End If
End With
If X = Monster(K).CurX And y = Monster(K).CurY Then
'Call MonsterStuck(k, direction)
End If

End Sub
Private Sub CheckRightLeft(K As Integer)
If Map(Monster(K).CurX + 1, Monster(K).CurY).Tiletype = 0 Or Map(Monster(K).CurX + 1, Monster(K).CurY).Tiletype = 11 Then
    Call Pathfinding(K, "down")
    If Map(Monster(K).CurX - 1, Monster(K).CurY).Tiletype = 0 Or Map(Monster(K).CurX - 1, Monster(K).CurY).Tiletype = 11 Then
Call Pathfinding(K, "up")

End If
End If
End Sub
Private Sub CheckUpDown(K As Integer)
If Map(Monster(K).CurX, Monster(K).CurY - 1).Tiletype = 0 Or Map(Monster(K).CurX, Monster(K).CurY - 1).Tiletype = 11 Then
    Call Pathfinding(K, "right")
    If Map(Monster(K).CurX, Monster(K).CurY + 1).Tiletype = 0 Or Map(Monster(K).CurX, Monster(K).CurY + 1).Tiletype = 11 Then

Call Pathfinding(K, "left")
End If
End If
End Sub
Private Sub MonsterStuck(K As Integer, Direction As String)
Dim amount As Integer
Dim amount2 As Integer

newmon = "MonsterImage" & K
'On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51

Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
If Direction = "left" Then
Form1.Caption = "DAGADG'2"
amount = 0 '61
amount2 = 0 '85
Do Until amount = 61
amount = amount + 1
If Monster(K).CurY - amount <= 0 Then
amount = 61
End If
If Monster(K).CurX + 1 > 85 Then
amount = 61
End If
'If Map(Monster(k).CurX, Monster(k).CurY - amount).TileType = 1 And Map(Monster(k).CurX, Monster(k).CurY - amount).TileType = 11 And Map(Monster(k).CurX + 1, Monster(k).CurY - amount).TileType = 0 And Map(Monster(k).CurX + 1, Monster(k).CurY - amount).TileType = 11 Then
If Map(Monster(K).CurX, Monster(K).CurY - amount).Tiletype = 1 And Map(Monster(K).CurX + 1, Monster(K).CurY - amount).Tiletype = 1 Then
      
   
amount = amount + 3
    Do Until amount = 0
        .Top = .Top - 32
        Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
        Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
        
        amount = amount - 1
    Loop
    .Left = .Left + 32
    Monster(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
    Monster(K).CurY = (Snap(.Top, 32) \ 32) + VScroll
    amount = 61
End If
Loop


End If

End With

End Sub
'***************************************************************************************'
' Purpose:  To create the height and length, and the color of the buildings        '                                                                        '
'***********************************************************************************'***'
Private Sub TownCreate(Buildings As Integer, Trees As Integer)
Randomize
buildingcolor = Int(Rnd * 3)
Counter = Buildings
Buildinglength = Int(Rnd * 10)
buildingheight = Int(Rnd * 8)
If buildingheight < 1 Then
buildingheight = 1
End If
If Buildinglength < 3 Then
Buildinglength = 3
End If
'MsgBox buildingcolor & " " & Buildinglength & " "
   Do While Counter > 0
      Counter = Counter - 1
        X1 = 23 * mWidth
        Y1 = 16 * mHeight
        RanX = Int(Rnd * X1)
        RanY = Int(Rnd * Y1)
        Buildinglength = Int(Rnd * 20)
    buildingheight = Int(Rnd * 20)
    If buildingheight < 1 Then
    buildingheight = 1
    End If
    If Buildinglength < 3 Then
    Buildinglength = 3
    End If
        If RanX + Buildinglength >= 90 Then
        RanX = RanX - Buildinglength - 2
        End If
        
        If buildingcolor = 0 Then ' blue
        '1st create the building bottom
        I = 0 '-1
        For I = I + 1 To Buildinglength
        If I = 1 Then ' create the left cornor
        Map(RanX + I, RanY).TileX = 2560
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
     '   Debug.Print I
        End If
       ' Debug.Print I
        If I > 1 And I < Buildinglength Then 'Create the middle
        Map(RanX + I, RanY).TileX = 2592
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
        If I = Buildinglength Then  ' create the right cornor
        Map(RanX + I, RanY).TileX = 2464
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
                'now create the roof
            For H = 0 To buildingheight
            If (RanY - H - 1) < 0 Then
                  Else
                      Map(RanX + I, RanY - H - 1).TileX = 2176
                      Map(RanX + I, RanY - H - 1).Tiley = 64
                      Map(RanX + I, RanY - H - 1).Tiletype = NONWALKABLE
            End If
            Next H

        Next I
        
        End If
        If buildingcolor = 1 Then ' red
'1st create the building bottom
        I = -1
        For I = I + 1 To Buildinglength
        If I = 0 Then ' create the left cornor
        Map(RanX + I, RanY).TileX = 2560
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
        If I > 0 And I < Buildinglength Then  'Create the middle
        Map(RanX + I, RanY).TileX = 2592
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
        If I = Buildinglength Then ' create the right cornor
        Map(RanX + I, RanY).TileX = 2464
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
                      'now create the roof
         For H = 0 To buildingheight
            If (RanY - H - 1) < 0 Then
                  Else
                      Map(RanX + I, RanY - H - 1).TileX = 2688
                      Map(RanX + I, RanY - H - 1).Tiley = 32
                      Map(RanX + I, RanY - H - 1).Tiletype = NONWALKABLE
            End If
            Next H
        
        Next I
        End If
        
        If buildingcolor = 2 Or buildingcolor = 3 Then ' wood Building
 '1st create the building bottom
        I = -1
        For I = I + 1 To Buildinglength
        If I = 0 Then ' create the left cornor
        Map(RanX + I, RanY).TileX = 2560
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
        If I > 0 And I < Buildinglength Then  'Create the middle
        Map(RanX + I, RanY).TileX = 2592
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
        If I = Buildinglength Then  ' create the right cornor
        Map(RanX + I, RanY).TileX = 2464
        Map(RanX + I, RanY).Tiley = 0
        Map(RanX + I, RanY).Tiletype = NONWALKABLE
        End If
                      'now create the roof
     For H = 0 To buildingheight
            If (RanY - H - 1) < 0 Then
                  Else
                      Map(RanX + I, RanY - H - 1).TileX = 2944
                      Map(RanX + I, RanY - H - 1).Tiley = 64
                      Map(RanX + I, RanY - H - 1).Tiletype = NONWALKABLE
            End If
            Next H
                
        Next I
        End If
      
      If Counter <= 0 Then   ' If condition is True.
      Exit Do   ' Exit loop.
      End If
      
Loop
Call GrassTrees(Trees)
Form1.RedrawMap

End Sub

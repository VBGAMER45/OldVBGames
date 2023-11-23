Attribute VB_Name = "Server"
'Well i have stoped work on online play right now
'Maybe one day I will get back to it

'Stuff Defined
'MainPlayer - is the main player on the screen
'Player(i) - is the other people
'NewItem - is the item on the screen that you see

'monster(i) - holds the information of the monsters on the screen
'NPc(i) holds the npc information

'PlayerSkill - are the persons skills not added yet
'to be able to pick up an item from any side
'on  form1 image1 is your main image for the player
'to do stop walking over npc's and monsters
'to do pathfinding for monsters
'finsh rnd map maker
Dim pPlayerNumber As Integer
Public Sub ServerCommands(Command As String, SecondInformation As String, Data As String)
Dim MainArray() As String  '20 is the max amount of seperate information the server is going to send
MainArray = Tokenize(Data)
Select Case Command
Case "[Board]" 'Handles Message Board Posts
    CurrentPost = CurrentPost + 1
    Messageboard(CurrentPost).Subject = MainArray(1)
    Messageboard(CurrentPost).Post = MainArray(2)
    Messageboard(CurrentPost).Author = MainArray(3)
    TempStr3 = ""
    TempStr3 = TempStr3 & Messageboard(CurrentPost).Subject
    Pics.Posts.AddItem TempStr3, Pics.Posts.ListIndex + 1

Case "[Cord]" 'Updates each player cords
    pPlayerNumber = -1
    If ReadField(2, Data, 44) = MainPlayer.PlayerName Then Exit Sub
    For pPlayerNumber = pPlayerNumber + 1 To 50
    If Player(pPlayerNumber).PlayerName = ReadField(2, Data, 44) Then
    Player(pPlayerNumber).MapName = ReadField(3, Data, 44)
    Player(pPlayerNumber).xpos = ReadField(4, Data, 44)
    Player(pPlayerNumber).ypos = ReadField(5, Data, 44)
    Call ItemInfo.MovenewPlayer
    End If
    Next
Case "[RestorePlayer]" ' Server command heals player
        If MainPlayer.PlayerName = Data Then
        MainPlayer.Health = MainPlayer.MaxHealth
        MainPlayer.Stamina = MainPlayer.MaxStamina
        MainPlayer.Mana = MainPlayer.MaxMana
        End If
        
Case "[Start]"  'Warps a player back to the begining
        If MainPlayer.PlayerName = Data Then
            MainPlayer.MapName = "newbie"
            Form1.MapName = "newbie"
            Open App.Path & "\maps\newbie" For Input As #1: Close
            Open App.Path & "\maps\newbie" For Binary Access Read Lock Read As #1
            Form1.EraseAll
            Get #1, , Map
            Close #1
            Form1.RedrawMap
            Form1.HScroll1.Value = "45"
            Form1.VScroll1.Value = "23"
        End If
Case "[WarpPlayer]" ' warps a player to a certain map
'to do handle the exact cords
    If MainPlayer.PlayerName = MainArray(0) Then
        MainPlayer.MapName = MainArray(1)
        Form1.MapName = MainArray(1)
        On Error Resume Next
        Open App.Path & "\maps\" & MainArray(1) For Input As #1: Close
        Open App.Path & "\maps\" & MainArray(1) For Binary Access Read Lock Read As #1
        Form1.EraseAll
        Get #1, , Map
        Close #1
        Call RPG.FastXYWarp(Int(MainArray(2)), Int(MainArray(3)))
        
        Form1.RedrawMap
    End If
Case "[MuzzlePlayer]"
    If MainPlayer.AccountName = MainArray(0) And MainPlayer.PlayerName = MainArray(1) Then
        If Form1.txtMessage.Visible = True Then
            Form1.txtMessage.Visible = False
            Form1.cmdSend.Visible = False
            Form1.txtDialog.Text = Form1.txtDialog.Text & "You have been muzzled" & vbCrLf
            Exit Sub
        End If
        If Form1.txtMessage.Visible = False Then
            Form1.txtMessage.Visible = True
            Form1.cmdSend.Visible = True
            Form1.txtDialog.Text = Form1.txtDialog.Text & "You have been unmuzzled" & vbCrLf
            Exit Sub
        End If
    End If
Case "[Jail]" 'sends player to jail
If MainPlayer.PlayerName = Data Then
    MainPlayer.MapName = "Jail.map"
    Form1.MapName = "Jail.map"
    Open App.Path & "\maps\Jail.map" For Input As #1: Close
    Open App.Path & "\maps\Jail.map" For Binary Access Read Lock Read As #1
    Form1.EraseAll
    Get #1, , Map
    Close #1
    Form1.RedrawMap
    Form1.HScroll1.Value = "47"
    Form1.VScroll1.Value = "28"
    Form1.Image1.Top = 208
    Form1.Image1.Left = 272
End If
Case "[Weather]"
'to handle the weather data
Case "[Tell]" ' a private message to the player
        If MainPlayer.PlayerName = MainArray(0) Then
        Form1.txtDialog.Text = Form1.txtDialog.Text & "Server" & " tells you! " & MainArray(1) & vbCrLf
        End If
Case "[Reconnect]"
'Only problem with this that i can think of is two message boxes appear
Call Form1.MessageBoxSystem(Data)
Unload Form1
Unload Form2
Unload Items

Unload TOOLS
Case "[WithDraw]"
If MainArray(0) = MainPlayer.PlayerName Then
    MainPlayer.BankAccountMoney = MainPlayer.BankAccountMoney - MainArray(1)
    MainPlayer.Money = MainPlayer.Money + MainArray(1)
End If
Case "[Deposit]"
If MainArray(0) = MainPlayer.PlayerName Then
    MainPlayer.BankAccountMoney = MainPlayer.BankAccountMoney - MainArray(1)
    MainPlayer.Money = MainPlayer.Money + MainArray(1)
End If
Case "[s]" ' storm command
If MainArray(3) = "3" Then
    Form1.TimerDraw.Enabled = True
Else
    Form1.TimerDraw.Enabled = False
End If
Case "[W]"
'I am going to check this out on the server looks alot like the code above
Form1.HScroll2.Value = MainArray(3)
If MainArray(2) = "3" Then
    Form1.TimerDraw.Enabled = True
Else
    Form1.TimerDraw.Enabled = False
End If
Case "[px]"
Form1.txtDialog.Text = Form1.txtDialog.Text & Data & vbCrLf
Case "[snd]"
If MainPlayer.PlayerName = MainArray(0) Then
    Form1.MediaPlayer2.Filename = MainArray(1)
End If
End Select
End Sub
Public Function Tokenize(pData As String)
Dim Num As Integer 'number of tokens
Dim CurWord As String 'current token
Dim WL() As String 'array of tokens
Dim ch As String 'current character
Dim SepChar As String 'separation character
SepChar = "," 'make sepchar a comma
ReDim WL(0) As String
Num = -1
pData = pData + SepChar 'add a comma to the end of the string, to
' make sure the last word/token is stored
For t% = 1 To Len(pData)
ch = Mid$(pData, t%, 1)
If ch = SepChar Then
If CurWord <> "" Then 'save word, start a new one
Num = Num + 1
ReDim Preserve WL(Num) As String
WL(Num) = CurWord 'save word to list
CurWord = "" 'reset CurWord to empty string
End If
Else
CurWord = CurWord + ch 'CH is not a comma, so add it to the
'current word...
End If
Next t%
Tokenize = WL 'returns an array of strings
End Function
Public Sub ServerLoop()
'Send Charecter List and send player information



End Sub
Public Sub Nocheating()
If MapEd = True Then
Else
If MainPlayer.Admin = False Then


End If
End If

End Sub
Public Sub LoadPlayerList()
I = -1
Form1.txtDialog.Text = Form1.txtDialog.Text & "Players:" & vbCrLf
For I = I + 1 To Form1.lstConnections.ListCount - 1
If Form1.lstConnections.List(I) = "[Open]" Then
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & Form1.lstConnections.List(I) & vbCrLf
End If


Next

End Sub
Public Sub AttackPlayer()
'This is the client side code for attacking other people
Dim newmon As String
Dim K As Integer
K = 0

For K = K + 1 To 150
newmon = "PlayerImage" & K
On Error GoTo nofile:
  Set P51 = Form1.Controls(newmon)
With P51
Player(K).CurX = (Snap(.Left, 32) \ 32) + HScroll
Player(K).CurY = (Snap(.Top, 32) \ 32) + VScroll

Call ItemInfo.UseWeapon
If TOOLS.AttackPlayers.Value = 1 Then
If MainPlayer.xpos = Player(K).CurX And MainPlayer.ypos = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If

'+y
If MainPlayer.xpos + 1 = Player(K).CurX And MainPlayer.ypos = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If
'-y
If MainPlayer.xpos - 1 = Player(K).CurX And MainPlayer.ypos = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If

'x-
If MainPlayer.xpos = Player(K).CurX And MainPlayer.ypos - 1 = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If

'x+
If MainPlayer.xpos = Player(K).CurX And MainPlayer.ypos + 1 = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If

'++
If MainPlayer.xpos + 1 = Player(K).CurX And MainPlayer.ypos + 1 = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If

'-1+1
If MainPlayer.xpos - 1 = Player(K).CurX And MainPlayer.ypos + 1 = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If

'+-
If MainPlayer.xpos + 1 = Player(K).CurX And MainPlayer.ypos - 1 = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If

'--
If MainPlayer.xpos - 1 = Player(K).CurX And MainPlayer.ypos - 1 = Player(K).CurY Then
Call Form1.SendClientData("PvsP", MainPlayer.PlayerName & "," & Player(K).PlayerName)
End If
End If
End With

Next K
nofile:

Exit Sub
End Sub
Function ReadField(ByVal pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
On Error Resume Next
'*****************************************************************
'Gets a field from a string
'*****************************************************************
Dim I As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For I = 1 To Len(Text)
    CurChar = Mid(Text, I, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = I
    End If
Next I

FieldNum = FieldNum + 1
If FieldNum = pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function

Public Sub SaveChatLog()
If OnlineGame = True And TOOLS.SaveChat.Value = 1 Then
Open App.Path & "\" & MainPlayer.PlayerName & ".log" For Append Shared As #5
Print #5, "**** Player Logged off " & Time
Print #5, Form1.txtDialog.Text
Close #5
End If
End Sub
Function LegalCharacter(KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************

'if backspace allow
If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If

'Only allow space,numbers,letters and special characters
If KeyAscii < 32 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If

'Check for bad special characters in between
If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If

'else everything is cool
LegalCharacter = True

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SIndex As Integer, ATIndex As Integer
Public PosX As Integer, PosY As Integer
Public NPCName As String, AlreadyTalkedTo As Boolean


Public Sub LoadMap(Path As String)
Dim tStr As String, t1 As String, t2 As String
Dim Tag As String
Open Path For Input As 1
Do Until EOF(1)
    Line Input #1, Tag
    VerifyTag Tag
Loop
Close
End Sub

Private Sub VerifyTag(Tag As String)
If Tag = "<PPOS>" Then
    InputPPos
ElseIf Tag = "<MAP>" Then
    InputMap
ElseIf Tag = "<NPC>" Then
    InputNPC
ElseIf Tag = "<SPEECH>" Then
    InputSpeech
End If
End Sub

Private Sub InputSpeech()
Dim tStr As String, i As Integer, j As Integer
Line Input #1, tStr
    Do Until EOF(1)
    i = Val(Mid(tStr, 2, Len(tStr) - 2))
    Line Input #1, Speech(i).Pitanje
    Do
        Line Input #1, tStr
        If Left(tStr, 1) = "<" Then
            Speech(i).nOdgovor = j
            Exit Do
        End If
        Speech(i).Redirect(j) = Val(Left(tStr, 4))
        Speech(i).Odgovor(j) = Mid(tStr, 6)
        j = j + 1
    Loop
    j = 0
Loop
End Sub

Private Sub InputNPC()
Dim tStr As String, i As Integer
Do
    Line Input #1, tStr
    If tStr = "</NPC>" Then Exit Do
    NPCTalk(i).NPCName = tStr
    Line Input #1, tStr
    NPCTalk(i).PosX = Val(tStr)
    Line Input #1, tStr
    NPCTalk(i).PosY = Val(tStr)
    Line Input #1, tStr
    NPCTalk(i).SIndex = Val(tStr)
    Line Input #1, tStr
    NPCTalk(i).ATIndex = Val(tStr)
    i = i + 1
Loop
End Sub

Private Sub InputPPos()
Dim pX As Integer, pY As Integer
Input #1, pX
Input #1, pY

End Sub

Private Sub InputMap()
Dim CounterX As Integer, CounterY As Integer, tStr As String
Line Input #1, tStr
Do Until tStr = "</MAP>"
    Do Until CounterX = Len(tStr)
        CounterX = CounterX + 1
        LenX = CounterX
       ' NpcMap(CounterY, CounterX - 1) = Asc(Mid(tStr, CounterX, 1))
    Loop
    CounterX = 0
    CounterY = CounterY + 1
    Line Input #1, tStr
Loop
LenY = CounterY
End Sub

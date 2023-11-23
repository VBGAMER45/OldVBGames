VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Train Skill"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   Icon            =   "TrainSKILL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "TrainSKILL.frx":030A
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   720
      OleObjectBlob   =   "TrainSKILL.frx":0384
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Technolgies"
      Height          =   2775
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton CmdTrain 
         Caption         =   "Train"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ListBox TechList 
         Height          =   1425
         ItemData        =   "TrainSKILL.frx":0586
         Left            =   360
         List            =   "TrainSKILL.frx":05B1
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.TextBox Points 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   0
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1620
      ItemData        =   "TrainSKILL.frx":0651
      Left            =   1920
      List            =   "TrainSKILL.frx":0653
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "TrainSKILL.frx":0655
      Left            =   240
      List            =   "TrainSKILL.frx":0657
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Train"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SkillPoints:"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
Unload Me
End Sub



Private Sub Command1_Click()
If List1.Text = "Melee Defense 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Melee Defense!" & vbNewLine
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbNewLine
Unload Me
End If
End If
If List1.Text = "Run 10" Then
If Points.Text >= 10 Then
Points.Text = Points.Text - 10
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Run!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Cast Spell 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Cast Spell!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Swim 4" Then
If Points.Text >= 4 Then
Points.Text = Points.Text - 4
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Swim!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Scan 4" Then
If Points.Text >= 4 Then
Points.Text = Points.Text - 4
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Scan!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Stealth 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Stealth!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Sword 18" Then
If Points.Text >= 18 Then
Points.Text = Points.Text - 18
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Sword!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Dagger 8" Then
If Points.Text >= 8 Then
Points.Text = Points.Text - 8
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Dagger!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Fail 10" Then
If Points.Text >= 10 Then
Points.Text = Points.Text - 10
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Fail!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Staff 12" Then
If Points.Text >= 12 Then
Points.Text = Points.Text - 12
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Staff!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Spear 10" Then
If Points.Text >= 10 Then
Points.Text = Points.Text - 10
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Spear!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "The Force 30" Then
If Points.Text >= 30 Then
Points.Text = Points.Text - 30
List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex
Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained TheForce!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text
Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Unarmed Attack 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20


List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Unarmed Attack!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Bow 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20


List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Bow!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Bombs 16" Then
If Points.Text >= 16 Then
Points.Text = Points.Text - 16


List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Bombs!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "BlackSmith 20" Then
If Points.Text >= 20 Then
Points.Text = Points.Text - 20

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Blacksmith!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Miner 18" Then
If Points.Text >= 18 Then
Points.Text = Points.Text - 18


List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Miner!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Hunter 14" Then
If Points.Text >= 14 Then
Points.Text = Points.Text - 14


List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Hunter!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Farmer 22" Then
If Points.Text >= 22 Then
Points.Text = Points.Text - 22

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Farmer!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Builder 16" Then
If Points.Text >= 16 Then
Points.Text = Points.Text - 16


List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained builder!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me
Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Magic Defense 24" Then
If Points.Text >= 24 Then
Points.Text = Points.Text - 24


List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained MagicDefense!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me


Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
If List1.Text = "Mace 6" Then
If Points.Text >= 6 Then
Points.Text = Points.Text - 6

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Mace!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me


Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me

End If
End If
If List1.Text = "Axe 6" Then
If Points.Text >= 6 Then
Points.Text = Points.Text - 6

List2.AddItem List1.Text
List1.RemoveItem List1.ListIndex


Form1.txtDialog.Text = Form1.txtDialog.Text & "You trained Axe!" & vbCrLf
SaveList List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt"
SaveList List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt"
MainPlayer.SkillPoints = Points.Text


Unload Me


Else
Form1.txtDialog.Text = Form1.txtDialog.Text & "You do not have enough skillpoints!" & vbCrLf
Unload Me
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'load player skillpoints
Points.Text = MainPlayer.SkillPoints

Call LoadList(List1, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skillsleft.txt")

Call LoadList(List2, App.Path & "\OfflineSaves\" & MainPlayer.PlayerName & "skills.txt")
If SkinFilename = "" Then Exit Sub
  Skn.LoadSkin SkinFilename ' Loads another skin into Skin component
  Skn.ApplySkin Me.hwnd
  Points.Visible = False
  SkinLabel1.Caption = "Skillpoints: " & Points.Text
  SkinLabel1.Visible = True
End Sub
Sub LoadList(Lst As ListBox, file As String)
'Call LoadList (List1,"C:\Windows\System\Saved.lst")
On Error GoTo error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
Lst.AddItem A$
Loop
Close 1
Exit Sub
error:
X = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub
Sub SaveList(Lst As ListBox, file As String)
'Call SaveList (List1,"C:\Windows\System\Saved.lst")
On Error GoTo error
Open file For Output As #1
For I = 0 To Lst.ListCount - 1
A$ = Lst.List(I)
Print #1, A$
Next
Close 1
Exit Sub
error:
X = MsgBox("There has been a error!", vbOKOnly, "Error")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call TOOLS.ReLoadSkills
  Call ImplodeForm(Me, 2, 500, 1)
End Sub


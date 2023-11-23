VERSION 5.00
Begin VB.Form frmjv 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Space Invaders"
   ClientHeight    =   6825
   ClientLeft      =   330
   ClientTop       =   615
   ClientWidth     =   8445
   Icon            =   "SPACE1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3120
      Picture         =   "SPACE1.frx":0442
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   18
      Top             =   6240
      Width           =   615
   End
   Begin VB.PictureBox explode 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "SPACE1.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox life4 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8160
      Picture         =   "SPACE1.frx":193A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      Top             =   5640
      Width           =   375
   End
   Begin VB.PictureBox life5 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7200
      Picture         =   "SPACE1.frx":2330
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   5640
      Width           =   375
   End
   Begin VB.PictureBox life1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7680
      Picture         =   "SPACE1.frx":2D26
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   6000
      Width           =   375
   End
   Begin VB.PictureBox life2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8160
      Picture         =   "SPACE1.frx":371C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      Top             =   6360
      Width           =   375
   End
   Begin VB.PictureBox life3 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7200
      Picture         =   "SPACE1.frx":4112
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   6360
      Width           =   375
   End
   Begin VB.PictureBox enpic5 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      Picture         =   "SPACE1.frx":4B08
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox enpic4 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "SPACE1.frx":59EA
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox enpic3 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "SPACE1.frx":6BB4
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox enpic2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "SPACE1.frx":7B46
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox enpic1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "SPACE1.frx":8AD8
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   125
      Left            =   1560
      Top             =   120
   End
   Begin VB.PictureBox Enemy2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2160
      MousePointer    =   2  'Cross
      Picture         =   "SPACE1.frx":9ECA
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox enemy3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3480
      MousePointer    =   2  'Cross
      Picture         =   "SPACE1.frx":AE5C
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox enemy4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "SPACE1.frx":BDEE
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox enemy5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      MousePointer    =   2  'Cross
      Picture         =   "SPACE1.frx":CFB8
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox enemy1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "SPACE1.frx":DE9A
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   4980
      Left            =   -1560
      Picture         =   "SPACE1.frx":F28C
      Top             =   480
      Width           =   8655
   End
   Begin VB.Label scorelbl 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Space Invaders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu sover 
         Caption         =   "Start &Over"
      End
      Begin VB.Menu mnuinterface 
         Caption         =   "&Interface"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu instruct 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmjv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response As String
Dim done As Boolean
Dim lives As Integer
Dim pic As Integer
Dim score As Integer

Private Sub about_Click()
frmAbout.Show vbModal, Me

End Sub

Private Sub enemy1_click()
    enemy1.Picture = explode.Picture
    For i = 1 To 250000
    Next i
    score = score + 10
    Call choosepic(1)
    enemy1.Top = 720
    done = False
    
    
End Sub

Private Sub Enemy2_click()
    Enemy2.Picture = explode.Picture
    For i = 1 To 250000
    Next i
    score = score + 10
    Call choosepic(2)
    Enemy2.Top = 720
    done = False
End Sub

Private Sub enemy3_click()
    enemy3.Picture = explode.Picture
    For i = 1 To 250000
    Next i
    score = score + 10
    Call choosepic(3)
    enemy3.Top = 720
    done = False
End Sub

Private Sub enemy4_click()
    enemy4.Picture = explode.Picture
    For i = 1 To 250000
    Next i
    score = score + 10
    Call choosepic(4)
    enemy4.Top = 720
    done = False
End Sub

Private Sub enemy5_click()
    enemy5.Picture = explode.Picture
    For i = 1 To 250000
    Next i
    score = score + 10
    Call choosepic(5)
    enemy5.Top = 720
    done = False
End Sub


Private Sub exit_Click()
    End
End Sub

Private Sub Form_Load()
    lives = 5
    MsgBox "Level 1", , "New Level"
End Sub

Private Sub instruct_Click()
MsgBox "LOOK IN READ ME FOR HELP"

End Sub

Private Sub mnuinterface_Click()

'form1.Show vbModal, Me

Unload Me

End Sub

Private Sub sover_Click()
    Call totop
    score = 0
    lives = 5
    life1.Visible = True
    life2.Visible = True
    life3.Visible = True
    life4.Visible = True
    life5.Visible = True
    Timer1.Interval = 125
End Sub

Private Sub timer1_timer()
    enemy1.Top = enemy1.Top + 70
    Enemy2.Top = Enemy2.Top + 100
    enemy3.Top = enemy3.Top + 80
    enemy4.Top = enemy4.Top + 60
    enemy5.Top = enemy5.Top + 90
    scorelbl.Caption = score
    
    If score = 100 And done = False Then
        MsgBox "Level 2", , "New Level"
        Call totop
        done = True
        Timer1.Interval = 75
    End If
    
    If score = 200 And done = False Then
        MsgBox "Level 3", , "New Level"
        Call totop
        done = True
        Timer1.Interval = 25
    End If
    
    If done = False And score = 300 Then
            MsgBox "Level 4", , "New Level"
            Call totop
            done = True
            Timer1.Interval = 10
    End If
    
    If score = 400 And done = False Then
         MsgBox "Level 5", , "New Level"
        Call totop
         done = True
        Timer1.Interval = 7
    End If
    
    
    If score = 500 Then
        MsgBox "Congratulations, You Win!", , "WINNER"
        MsgBox "You had " & score & " Points:"
        Call pgain
    End If
    
    
     
        
        
    If enemy1.Top >= 6300 Then
        Call loseship
        enemy1.Top = 720
    End If
    
    If Enemy2.Top >= 6300 Then
        Call loseship
        Enemy2.Top = 720
    End If
    
    If enemy3.Top >= 6300 Then
        Call loseship
        enemy3.Top = 720
    End If
    
    If enemy4.Top >= 6300 Then
        Call loseship
        enemy4.Top = 720
    End If
    
    If enemy5.Top >= 6300 Then
        Call loseship
        enemy5.Top = 720
    End If
    
End Sub

Public Sub choosepic(enemyname)
    pic = CInt((5 * Rnd) + 0.5)
    Select Case pic
        Case 1
            Select Case enemyname
                Case 1
                    
                    enemy1.Picture = enpic1.Picture
                 Case 2
                 
                    Enemy2.Picture = enpic1.Picture
                 Case 3
             
                    enemy3.Picture = enpic1.Picture
                 Case 4
                   
                    enemy4.Picture = enpic1.Picture
                 Case 5
                   
                    enemy5.Picture = enpic1.Picture
            End Select
            
        Case 2
            Select Case enemyname
                Case 1
              
                    enemy1.Picture = enpic2.Picture
                   
                 Case 2
              
                    Enemy2.Picture = enpic2.Picture
                    
                 Case 3
                
                    enemy3.Picture = enpic2.Picture
                   
                 Case 4
              
                    enemy4.Picture = enpic2.Picture
                   
                 Case 5
               
                    enemy5.Picture = enpic2.Picture
                   
            End Select
            
        Case 3
            Select Case enemyname
                Case 1
                    enemy1.Picture = enpic3.Picture
                     
                 Case 2
                    Enemy2.Picture = enpic3.Picture
                    
                 Case 3
                    enemy3.Picture = enpic3.Picture
                  
                 Case 4
                    enemy4.Picture = enpic3.Picture
                
                 Case 5
                    enemy5.Picture = enpic3.Picture
                 
            End Select
            
        Case 4
            Select Case enemyname
                Case 1
         
                    enemy1.Picture = enpic4.Picture
                 Case 2
           
                    Enemy2.Picture = enpic4.Picture
                 Case 3
              
                    enemy3.Picture = enpic4.Picture
                 Case 4
                  
                    enemy4.Picture = enpic4.Picture
                 Case 5
               
                    enemy5.Picture = enpic4.Picture
            End Select
            
            Case 5
            Select Case enemyname
                Case 1
        
                    enemy1.Picture = enpic5.Picture
                 Case 2
                
                    Enemy2.Picture = enpic5.Picture
                 Case 3
                 
                    enemy3.Picture = enpic5.Picture
                 Case 4
                  
                    enemy4.Picture = enpic5.Picture
                 Case 5
                
                    enemy5.Picture = enpic5.Picture
            End Select
       
        End Select
  End Sub


Public Sub loseship()
    MsgBox "YOU LOST A LIFE"
    lives = lives - 1
    If lives = 0 Then
        MsgBox "You Lose!"
        MsgBox "You had " & score & " Points:"
        Call pgain
    End If
    
    Select Case lives
        Case 4
            life5.Visible = False
        Case 3
            life4.Visible = False
        Case 2
            life3.Visible = False
        Case 1
            life2.Visible = False
    End Select
        
        Call totop
End Sub


Public Sub totop()
    enemy1.Top = 720
    Enemy2.Top = 720
    enemy3.Top = 720
    enemy4.Top = 720
    enemy5.Top = 720
    
End Sub

Public Sub pgain()
   response = MsgBox("Would you like to play again?", vbYesNo, "PLay Again?")
   If response = vbYes Then
    Call totop
    score = 0
    lives = 5
    life1.Visible = True
    life2.Visible = True
    life3.Visible = True
    life4.Visible = True
    life5.Visible = True
    Timer1.Interval = 125
   Else
    End
    End If
End Sub

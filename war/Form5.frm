VERSION 5.00
Object = "{2E32FE95-CB5C-11D3-8FE2-0040D00AB433}#1.0#0"; "MOON.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Wheater\Map"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin moon.UserControl1 UserControl11 
      Height          =   735
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Moon Rotation"
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox picRadar 
      BackColor       =   &H00000000&
      Height          =   2655
      Left            =   960
      ScaleHeight     =   2595
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Timer tmrRadar 
      Interval        =   10
      Left            =   360
      Top             =   2160
   End
   Begin VB.Timer Timer5 
      Interval        =   550
      Left            =   3120
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1680
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2400
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Moon Rotation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Radar Offline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   3120
      Picture         =   "Form5.frx":030A
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   720
      Picture         =   "Form5.frx":0614
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   2280
      Picture         =   "Form5.frx":091E
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   1680
      Picture         =   "Form5.frx":0C28
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4080
      Picture         =   "Form5.frx":0F32
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4080
      Picture         =   "Form5.frx":123C
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1680
      Picture         =   "Form5.frx":1546
      ToolTipText     =   "Rain"
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3120
      Picture         =   "Form5.frx":1850
      ToolTipText     =   "Sun"
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2400
      Picture         =   "Form5.frx":1B5A
      ToolTipText     =   "Lighting"
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "Form5.frx":1E64
      ToolTipText     =   "Cloudy"
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "        5 day Forecast"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Unload Me

End Sub

Private Sub Command2_Click()
MsgBox " To make the radar work you first have to advance to the age of radar then build a radar station"

End Sub

' Set form topmost
Private Sub Form_Activate()

    SetWindowPos Form5.hwnd, -1, 0, 0, 0, 0, 83
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = (27) Then
End
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
    End
End If
If KeyAscii = 18 Then
Unload Me

Form3.Show
End If
If KeyAscii = 112 Then
Unload Me
frmAbout.Show
End If


End Sub

Private Sub Form_Load()
Form5.KeyPreview = True
If form2.Picture10.ToolTipText = "Radar1" Then
Label1.Caption = "Radar Online"
End If


End Sub

' -----------------------------------------------------------------
' Reset main toolbar when unloading window
Private Sub Form_Unload(Cancel As Integer)
    
   ' fEditor.tbrEditor.Buttons("TOOLRADAR").Value = tbrUnpressed
    
End Sub

' -----------------------------------------------------------------
' Goto new map position when user clicks on radar map
Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim nNewPosX As Integer, nNewPosY As Integer  ' Variables to hold new map position
    
    ' Get clicked position
    nNewPosX = X - 3
    nNewPosY = Y - 3
    
    If NewPosX > -1 And nNewPosY > -1 Then
    
        ' Do not allow movement out of map extend
        If nNewPosX > Map.ExtentX - Map.DisplayExtentX Then nNewPosX = Map.ExtentX - Map.DisplayExtentX
        If nNewPosY > Map.ExtentY - Map.DisplayExtentY Then nNewPosY = Map.ExtentY - Map.DisplayExtentY
        
        'Set new position
        form2.hscmap.Value = nNewPosX
        form2.vscmap.Value = nNewPosY
        
        ' Exit radar window
        Unload Me
        
    End If
    
End Sub
' -----------------------------------------------------------------
' Draw radar map
Private Sub tmrRadar_Timer()

    Call DrawRadar
    Me.tmrRadar.Enabled = False
    
End Sub


Private Sub Image6_Click()
'Call Module1.CreateControl
'Call GameEngine.CreateControl
End Sub





Private Sub Timer1_Timer()
'If Timer1.Interval = 300 Then
'year = year + 10
'Call TIMER
Call choosepic(1)
'year = year + 10
'End If
Image1.Visible = True
End Sub
 Public Sub choosepic(enemyname)
    enemyname = CInt((5 * Rnd) + 0.5)
    Select Case pic
        Case 1
           Select Case enemyname
                Case 1
                    Image1.Picture = Image6.Picture
                    'enemy1.Picture = enpic1.Picture
               '     i = sndPlaySound(ByVal CStr(App.Path & "\SCREAM1.wav"), SND_ASYNC) = i = sndPlaySound(ByVal CStr(App.Path & "\SCREAM1..wav"), SND_ASYNC)
                 Case 2
                 Image2.Picture = Image6.Picture
                    'Enemy2.Picture = enpic1.Picture
                  '  i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC) = i = sndPlaySound(ByVal CStr(App.Path & "\SCREAM1..wav"), SND_ASYNC)
                 Case 3
             Image3.Picture = Image6.Picture
                    'enemy3.Picture = enpic1.Picture
               '     i = sndPlaySound(ByVal CStr(App.Path & "\PLACBLDG.wav"), SND_ASYNC) = i = sndPlaySound(ByVal CStr(App.Path & "\SCREAM1..wav"), SND_ASYNC)
                 Case 4
                   Image4.Picture = Image6.Picture
              '      enemy4.Picture = enpic1.Picture
                 Case 5
                   Image5.Picture = Image6.Picture
                   
           '         enemy5.Picture = enpic1.Picture
          End Select
            
       Case 2
            Select Case enemyname
                Case 1
              Image1.Picture = Image7.Picture
               '     enemy1.Picture = enpic2.Picture
               '
                 Case 2
              Image2.Picture = Image7.Picture
                 '   Enemy2.Picture = enpic2.Picture
                    
                 Case 3
                Image3.Picture = Image7.Picture
                   ' enemy3.Picture = enpic2.Picture
                   
                 Case 4
              Image4.Picture = Image7.Picture
                 '   enemy4.Picture = enpic2.Picture
                '
                 Case 5
               Image5.Picture = Image7.Picture
                   ' enemy5.Picture = enpic2.Picture
                   
            End Select
            
        Case 3
            Select Case enemyname
                Case 1
                  '  enemy1.Picture = enpic3.Picture
                     Image1.Picture = Image8.Picture
                 Case 2
                  '  Enemy2.Picture = enpic3.Picture
                    Image2.Picture = Image8.Picture
                 Case 3
                  '  enemy3.Picture = enpic3.Picture
                  Image3.Picture = Image8.Picture
                 Case 4
                   ' enemy4.Picture = enpic3.Picture
                Image4.Picture = Image8.Picture
                 Case 5
                    'enemy5.Picture = enpic3.Picture
                 Image5.Picture = Image8.Picture
            End Select
            
        Case 4
            Select Case enemyname
                Case 1
         Image1.Picture = Image9.Picture
             '       enemy1.Picture = enpic4.Picture
                 Case 2
           Image2.Picture = Image9.Picture
                '   ' Enemy2.Picture = enpic4.Picture
                 Case 3
              Image3.Picture = Image9.Picture
                  '  enemy3.Picture = enpic4.Picture
                 Case 4
                  Image4.Picture = Image9.Picture
                  '  enemy4.Picture = enpic4.Picture
                 Case 5
               Image5.Picture = Image9.Picture
                   ' enemy5.Picture = enpic4.Picture
            End Select
            '
            Case 5
            Select Case enemyname
                Case 1
        Image1.Picture = Image10.Picture
                    'enemy1.Picture = enpic5.Picture
                Case 2
                Image2.Picture = Image10.Picture
                 '   Enemy2.Picture = enpic5.Picture
                 Case 3
                 Image3.Picture = Image10.Picture
                '    enemy3.Picture = enpic5.Picture
                 Case 4
                  Image4.Picture = Image10.Picture
                 '   enemy4.Picture = enpic5.Picture
                 Case 5
                Image5.Picture = Image10.Picture
                '    enemy5.Picture = enpic5.Picture
            End Select
       
        End Select
End Sub

'jonathan valentin 1999
Private Sub Timer2_Timer()
Image3.Visible = True
Image5.Visible = True
Image2.Visible = True
Image4.Visible = True
Call choosepic(2)
End Sub
'Type ControlGroup
 ' ControlAmount As Integer
 ' ControlObjects(MAXCONTROLAMOUNT) As ControlObject
'End Type
'Type PropertyArray
'  Properties(MAXPROPERTIES) As Variant
'End Type
'Type ControlObject
 ' Outline As RECT
 ' ControlProperties As PropertyArray
 ' ControlType As Integer
'End Type
'Public Function CreateControl(ControlType, X, Y, Width, Height, Content) As ControlObject
'Dim NewControl As ControlObject
'NewControl.ControlType = ControlType
'Select Case ControlType
'Case CONTROLTYPE_PICTUREBOX
 ' NewControl.ControlProperties.Properties(CONTROLPROPERTY_PICTURE) = Content
 ' NewControl.Outline.Top = Y - Int(Height / 2)
 ' NewControl.Outline.bottom = Y + Int(Height / 2)
 ' NewControl.Outline.Left = X - Int(Width / 2)
'  NewControl.Outline.Right = X + Int(Width / 2)
'End Select
'End Sub
''  Properties(MAXPROPERTIES) As Variant
'End Type
'Type ControlObject
  ''C''on'trolType As Integer
''E'
Private Sub Timer3_Timer()
Call choosepic(3)
End Sub

Private Sub Timer4_Timer()
Call choosepic(4)

End Sub

Private Sub Timer5_Timer()
Call choosepic(5)
End Sub

Private Sub year_Click()

End Sub



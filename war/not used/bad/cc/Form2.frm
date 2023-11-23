VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "War is War"
   ClientHeight    =   6510
   ClientLeft      =   -255
   ClientTop       =   465
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   Begin VB.PictureBox Picture22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5520
      ScaleHeight     =   585
      ScaleWidth      =   1065
      TabIndex        =   34
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox Picture21 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4440
      ScaleHeight     =   585
      ScaleWidth      =   1065
      TabIndex        =   33
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3480
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   32
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   825
      TabIndex        =   31
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2520
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   30
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1680
      ScaleHeight     =   585
      ScaleWidth      =   825
      TabIndex        =   29
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      ScaleHeight     =   585
      ScaleWidth      =   825
      TabIndex        =   28
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4080
      ScaleHeight     =   585
      ScaleWidth      =   825
      TabIndex        =   27
      Top             =   2400
      Width           =   855
   End
   Begin VB.PictureBox Picture7 
      Height          =   615
      Left            =   6960
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   26
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6960
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   25
      ToolTipText     =   "Construction yard"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Down"
      Height          =   375
      Left            =   8880
      TabIndex        =   24
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Up"
      Height          =   375
      Left            =   8400
      TabIndex        =   23
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Down"
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Up"
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   5880
      Width           =   615
   End
   Begin VB.PictureBox Picture16 
      Height          =   735
      Left            =   8400
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture15 
      Height          =   735
      Left            =   8400
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture14 
      Height          =   855
      Left            =   8400
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture13 
      Height          =   735
      Left            =   8400
      ScaleHeight     =   675
      ScaleWidth      =   1155
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture12 
      Height          =   735
      Left            =   8400
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture11 
      Height          =   735
      Left            =   8400
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture10 
      Height          =   615
      Left            =   8400
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture6 
      Height          =   735
      Left            =   6960
      Picture         =   "Form2.frx":36A6
      ScaleHeight     =   675
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      Height          =   735
      Left            =   6960
      Picture         =   "Form2.frx":6328
      ScaleHeight     =   675
      ScaleWidth      =   1155
      TabIndex        =   10
      ToolTipText     =   "Refinary"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Height          =   735
      Left            =   6960
      Picture         =   "Form2.frx":6D0A
      ScaleHeight     =   675
      ScaleWidth      =   1155
      TabIndex        =   9
      ToolTipText     =   "Barrecks"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   615
      Left            =   6960
      Picture         =   "Form2.frx":7724
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   8
      ToolTipText     =   "Advanced Power Plant"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   6960
      Picture         =   "Form2.frx":A3B2
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   7
      ToolTipText     =   "Power Plant"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   7680
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Inventions"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dipomacly"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Communicate"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "News"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Map"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   6120
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   6600
      Y1              =   6480
      Y2              =   240
   End
   Begin VB.Label Power 
      Caption         =   "0"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Year"
      Height          =   255
      Left            =   6960
      TabIndex        =   19
      Top             =   480
      Width           =   735
   End
   Begin VB.Label year 
      Caption         =   "0"
      Height          =   255
      Left            =   7800
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label money 
      Caption         =   "5000"
      Height          =   255
      Left            =   8160
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuload 
         Caption         =   "Load"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnmain 
         Caption         =   "Main Menu"
      End
      Begin VB.Menu Mnurestart 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Options"
      Begin VB.Menu mnuset 
         Caption         =   "Setting"
      End
      Begin VB.Menu mnupassword 
         Caption         =   "Password"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grid As String
Dim red As String

Private Sub Command1_Click()
Dim Form9 As New Form9
Form9.Show vbModal, Me


End Sub

Private Sub Command2_Click()

Form5.Show vbModal, Me



End Sub

Private Sub Command3_Click()
Form7.Show vbModal, Me

End Sub

Private Sub Command4_Click()
Form4.Show vbModal, Me

End Sub

Private Sub Command5_Click()
Form6.Show vbModal, Me

End Sub

Private Sub Form_Load()
'If year.Caption >= 1800 Then
'Command2.Visible = True
'Else
'If year.Caption <= 1800 Then
'Command2.Visible = False
'End If
'End If
End Sub

Private Sub mnmain_Click()
Dim form3 As New form3
form3.Show vbModal, Me

End Sub

Private Sub MNUABOUT_Click()
frmAbout.Show

End Sub

Private Sub MNUEXIT_Click()
End
End Sub

Private Sub mnuload_Click()
CommonDialog1.ShowOpen

End Sub

Private Sub mnupassword_Click()
InputBox "Enter password,Cheater"
If Text = 2 Then
MsgBox " Cheater"

End If


End Sub

Private Sub mnusave_Click()
CommonDialog1.ShowSave
End Sub

Private Sub mnuset_Click()
Form8.Show vbModal, Me

End Sub

'Private Sub Picture1_Click()
Private Sub picture1_MouseDown(Button As Integer, _
        Shift As Integer, X As Single, Y As Single)

'If grid = True Then

    Picture1.Drag vbBeginDrag
    Set Picture1.DragIcon = LoadPicture("c:\windows\desktop\vbprojects\war\trash02a.ico") '("c:\Program _
        files\ Microsoft Visual _
       Basic\Icons\Dragdrop\Dragfldr.ico")
'End Sub
'do
Power = Power + 0
Picture2.Visible = True
money = money - 1000





'Else

'If grid = False Then
'jv sound
'End If
'End If
End Sub



Private Sub picture17_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture17.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub


Private Sub picture18_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture18.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub

Private Sub picture19_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture19.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub

Private Sub Picture2_Click()

'If grid = True Then
'do
Power = Power + 100
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
money = money - 200
'End If

'Else

'If grid = False Then
'jv sound
'End If


End Sub


Private Sub picture20_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture20.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub


Private Sub picture21_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture21.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub


Private Sub picture22_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture22.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub

Private Sub Picture3_Click()
money = money - 1000
'If grid = True Then
'do
Power = Power + 200

'End If
'Else

'If grid = False Then
'jv sound
'End If


End Sub

Private Sub Picture4_Click()
If grid = True Then
'do
money = money - 500
Power = Power - 20
Picture16.Visible = True
Picture15.Visible = True
Picture14.Visible = True
Picture13.Visible = True



End If

Else

If grid = False Then
'jv sound
End If


End Sub

Private Sub Picture5_Click()

If grid = True Then
'do
Power = Power - 20
money = money - 100
'
End If

Else

If grid = False Then
'jv sound
End If


End Sub

Private Sub Picture6_Click()
money = money - 1500
Power = Power - 20

End Sub

'Private Sub Picture8_Click()
Private Sub picture8_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture8.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _
       'files\Microsoft Visual _
       ' Basic\Icons\Office\Files03a.ico")
'End Sub
End Sub


Private Sub picture9_DragOver(Source As Control, _
     X As Single, Y As Single, State As Integer)
   Source.Drag vbEndDrag
    Source.Visible = True
    Picture9.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")      '( "c:\Program _

End Sub

Private Sub Timer1_Timer()
'If Interval = 125 Then
'year = year + 10
'End If
'If Timer1.Interval = 300 Then

'End If
Call TIMER
year = year + 10
End Sub

Private Sub year_Click()
If year.Caption = 1520 Then
MsgBox "News travels quickey"
End If

End Sub
'Joathan Valentin 1999
'Private Sub Image1_MouseDown(Button As Integer, _
       ' Shift As Integer, X As Single, Y As Single)
   ' Image1.Drag vbBeginDrag
   ' Set Image1.DragIcon = LoadPicture("c:\Program _
      '  files\ Microsoft Visual _
      '  Basic\Icons\Dragdrop\Dragfldr.ico")
'End Sub
    
'Adding a DragOver event procedure to Image2 allows you to terminate dragging when the source enters the target. This example closes the file cabinet when Image1 is passed over Image2.
'Private Sub Image2_DragOver(Source As Control, _
  '      X As Single, Y As Single, State As Integer)
  '  Source.Drag vbEndDrag
  '  Source.Visible = False
 '   Image2.Picture = LoadPicture("c:\Program _
'        files\Microsoft Visual _
        Basic\Icons\Office\Files03a.ico")
'End Sub
    
'A'dding a third Image control to the form demonstrates canceling a drag operation. In this example the Image3 Picture property contains an icon of a trash can. Using the DragOver event and the source argument, dragging the files over Image3 cancels the drag operation.
'Private Sub Image3_DragOver(Source As Control, _
    '   ' X As Single, Y As Single, State As Integer)
    'Source.Drag vbCancel
'End Sub



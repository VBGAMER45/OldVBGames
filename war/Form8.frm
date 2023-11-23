VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Settings"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   2880
      Top             =   3960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sound"
      Height          =   2895
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      Begin VB.CheckBox Check2 
         Caption         =   "on"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Off"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   855
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   120
         Max             =   2
         Min             =   1
         TabIndex        =   9
         Top             =   600
         Value           =   1
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Wheater"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Game Speed"
      Height          =   2895
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         Max             =   3
         Min             =   1
         TabIndex        =   4
         Top             =   360
         Value           =   1
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Scoll Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "3DFX"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      Begin VB.OptionButton Option4 
         Caption         =   "None"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Other"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Diamond"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Voodo"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2040
      Picture         =   "Form8.frx":030A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   960
      Picture         =   "Form8.frx":0614
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   360
      Picture         =   "Form8.frx":091E
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1560
      Picture         =   "Form8.frx":0C28
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   240
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Check1_Click()
'If vbChecked Then
form2.Image27.Visible = True
'End If
End Sub

Private Sub Check2_Click()
'If vbChecked Then
form2.Image27.Visible = False
'End If
End Sub

Private Sub Command1_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
Call HScroll
Unload Me

End Sub

Public Sub HScroll()
If HSCOLL1 = 1 Then
Call TIME3
If HSCOLL1 = 2 Then
Call TIME1
If HSCOLL1 = 3 Then
Call TIME2
End If
End If
End If

End Sub

Private Sub Command2_Click()
i = sndPlaySound(ByVal CStr(App.Path & "\ramenu1.wav"), SND_ASYNC)
frmAbout.Show vbModal, Me

End Sub

'jonathan valentin 1999
Private Sub HScroll1_Change()
If HScroll1.Value = 3 Then
form2.Timer1 = -HScroll1.Value
form2.Timer1.Interval = 30
End If
If HScroll1.Value = 1 Then
form2.Timer1 = HScroll1.Value
form2.Timer1.Interval = 2000
End If
End Sub

Private Sub HScroll3_Change()
If HScroll3.Value = 2 Then
'form2.MediaPlayer1.Volume = form2.MediaPlayer1.Volume + 9
End If
If HScroll3.Value = 1 Then
'form2.MediaPlayer1.Volume = 0
End If
End Sub

Private Sub HScroll3_Scroll()
If HScroll3.Value = 2 Then
'form2.MediaPlayer1.Volume = 9
End If
If HScroll3.Value = 1 Then
'form2.MediaPlayer1.Volume = 9
End If
End Sub

Private Sub Timer1_Timer()
Image3.Top = Image3.Top + 30
Image3.Left = Image3.Left + 30
Image4.Top = Image4.Top + 30
Image4.Left = Image4.Left - 30
Image1.Top = Image1.Top + 30
Image1.Left = Image1.Left + 30
Image2.Top = Image2.Top + 30
Image2.Left = Image2.Left - 30
End Sub

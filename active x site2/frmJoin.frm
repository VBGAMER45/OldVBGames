VERSION 5.00
Begin VB.Form frmJoin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Join a session"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join Session"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ListBox lstSessions 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblStatus 
      Caption         =   "Please select a session and click 'Join'"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
' Basic DirectPlay Chat Program. NOTE: You can run
' and test this program on ONE computer by starting a
' TCP/IP session in one copy of the program, and then
' starting another copy and JOINING that session. When
' prompted for an IP address, just leave it blank.
' (this will search the localhost for sessions)
'
' As always, feel free to use this code, modify it,
' steal it, whatever! Lucky don't care.
'
' - Lucky
' Lucky's VB Gaming Site
' http://members.home.net/theluckyleper
'*****************************************************

Option Explicit

Dim blnUnloadWithoutDismantle As Boolean    'Should we dismatle directplay and return to the previous menu, or not?

'*****************************************************
' Purpose:  Unload the form
'*****************************************************

Private Sub cmdCancel_Click()

    'Unload the form
    Unload Me

End Sub

'*****************************************************
' Purpose:  Join the currently selected session
'*****************************************************

Private Sub cmdJoin_Click()

    'Try to join the session
    If Not (gobjDPlay.JoinGame(lstSessions.ListIndex, frmDPlayDevices.txtPlayerName)) Then
        Unload Me
        Exit Sub
    End If
    
    'Hide this form
    Me.Hide
    
    'Display the frmStart form
    frmStart.Show vbModal, Me
    
    'Start the message checking loop
    frmStart.MessageLoop
    
    'Unload without dismantling, we want to keep gobjDPlay alive here since we've joined a session!
    blnUnloadWithoutDismantle = True
  '  Unload Me

End Sub

'*****************************************************
' Purpose:  Reset directplay and redisplay the
'           frmDPlayDevices
'*****************************************************

Private Sub Form_Unload(Cancel As Integer)
       
    'If we don't want to dismantle directplay, skip the rest of this sub
    If blnUnloadWithoutDismantle Then Exit Sub
       
    'Reset directplay
    gobjDPlay.Dismantle
       
    'And allow the device listing form to be displayed again
  '  frmDPlayDevices.ReDisplay
    
End Sub


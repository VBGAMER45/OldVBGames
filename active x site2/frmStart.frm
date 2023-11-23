VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic Zone Chat"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   5280
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   840
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   1560
   End
   Begin VB.CommandButton cmdSendText 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   4680
      Width           =   4335
   End
   Begin VB.ListBox lstText 
      Height          =   2400
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   5775
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "Disable Join"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ListBox lstPlayers 
      Height          =   1035
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Send Text:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Chat Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Caption         =   "Current People:   VisualBasicZone.Com"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmStart"
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

'Our message checking loop control variable
Dim mblnCheckingMessages As Boolean

Private Sub Command1_Click()
MsgBox "To close press the x in the top right area", vbExclamation

'MsgBox "start"
End Sub

'*****************************************************
' Purpose:  Update the playerlist
'*****************************************************

Private Sub Form_Load()

    'Display the player's names
    gobjDPlay.UpdatePlayers lstPlayers
            
End Sub

'*****************************************************
' Purpose:  Unload the form
'*****************************************************

Private Sub cmdCancel_Click()

    'Unload the form
    Unload Me

End Sub

'*****************************************************
' Purpose:  Disable joining!
'*****************************************************

Private Sub cmdDisable_Click()

    'Disable joining...
    gobjDPlay.DisableJoin

End Sub

'*****************************************************
' Purpose:  Send the chat string
'*****************************************************

Private Sub cmdSendText_Click()
If txtChat.Text = "" Then

txtChat.Text = "."

End If
    'Send the chat string message
    gobjDPlay.SendChatText txtChat.Text
    
    'Display it on our listbox too
    lstText.AddItem frmDPlayDevices.txtPlayerName & ": " & txtChat.Text

End Sub

'*****************************************************
' Purpose:  Reset directplay and redisplay the
'           frmgobjDPlayDevices
'*****************************************************

Private Sub Form_Unload(Cancel As Integer)
    
    'Terminate the message checking loop
    mblnCheckingMessages = False
    
    'Reset directplay
    gobjDPlay.Dismantle
       
    'And allow the device listing form to be displayed again
  '  frmDPlayDevices.ReDisplay

End Sub

'*****************************************************
' Purpose:  Update the playerlist, start the message
'           checking loop
'*****************************************************

Public Sub MessageLoop()

    'Start the loop
    mblnCheckingMessages = True
    Do While mblnCheckingMessages
        'Check the messages
        gobjDPlay.GetMessages
        'Give windows a chance
        DoEvents
    Loop

End Sub



Private Sub Timer2_Timer()
cmdDisable.Enabled = False
End Sub

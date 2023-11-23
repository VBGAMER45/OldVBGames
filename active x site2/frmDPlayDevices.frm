VERSION 5.00
Begin VB.Form frmDPlayDevices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Internet!!"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3405
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtGameName 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "Visual Basic"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtPlayerName 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "Vb"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ListBox lstServiceProviders 
      Height          =   1425
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Session"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join Session"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Jonathan Valentin 2000"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Select Internet Option for Vb chat!!"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Session Name:"
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Your Name:"
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   2220
      Width           =   975
   End
End
Attribute VB_Name = "frmDPlayDevices"
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

'*****************************************************
' Purpose:  Create a gobjDPlay object and populate the
'           listbox with available service providers
'*****************************************************

Private Sub Form_Load()

    '(re)initialize the gobjDplay object
    Set gobjDPlay = Nothing
    Set gobjDPlay = New DirPlay
    
    'Populate the listbox
    If Not (gobjDPlay.InitConnectionList(lstServiceProviders)) Then
        'If there's an error, destroy directplay
        Set gobjDPlay = Nothing
        '..and unload this form
        Unload Me
    End If
                
End Sub

'*****************************************************
' Purpose:  Once the user has selected a device, allow
'           him to search for sessions to join
'*****************************************************

Private Sub cmdJoin_Click()

    'Ensure that the players enter a name
    If txtPlayerName.Text = "" Then
        MsgBox "You must fill out the Player Name field."
        Exit Sub
    End If
    
    'List the sessions. If there is an error, don't load the next form
    If Not (gobjDPlay.ListSessions(lstServiceProviders.ListIndex, frmJoin.lstSessions)) Then
        Set gobjDPlay = Nothing
        ReDisplay
        Exit Sub
    End If
    
    'Hide this form
    Me.Hide

    'Show the Join form
    frmJoin.Show vbModal, Me

End Sub

'*****************************************************
' Purpose:  Once the user has selected a device, allow
'           him to create a new session for other
'           people to join
'*****************************************************

Private Sub cmdStart_Click()

    'Ensure that the players enter a name
    If txtPlayerName.Text = "" Or txtGameName.Text = "" Then
        MsgBox "You must fill out both the Player Name and Game Name fields."
        Exit Sub
    End If
    
    'Start the session. If there is an error, reset directplay
    If Not (gobjDPlay.HostGame(lstServiceProviders.ListIndex, txtPlayerName.Text, txtGameName.Text, 10)) Then
        Set gobjDPlay = Nothing
        ReDisplay
        Exit Sub
    End If
    
    'Enable the "Disable Join" command button
    frmStart.cmdDisable.Enabled = True
    
    'Hide this form
    Me.Hide

    'Show the Start form
    frmStart.Show vbModal, Me
    
    'Start the message checking loop
    frmStart.MessageLoop

End Sub

'*****************************************************
' Purpose:  Reinitialize DPlay and populate the
'           listbox
'*****************************************************

Public Sub ReDisplay()

    '(re)initialize the gobjDplay object
    Set gobjDPlay = Nothing
    Set gobjDPlay = New DirPlay

    'Populate the listbox
    If Not (gobjDPlay.InitConnectionList(lstServiceProviders)) Then
        'If there's an error, destroy directplay
        Set gobjDPlay = Nothing
        '..and unload this form
     '   Unload Me
    End If
    
    'Show the form
'    Me.Show

End Sub

'*****************************************************
' Purpose:  Unload DPlay
'*****************************************************

Private Sub Form_Unload(Cancel As Integer)

    'Unload gobjDplay
    Set gobjDPlay = Nothing

End Sub


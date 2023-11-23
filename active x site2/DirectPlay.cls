VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'1111111 '*****************************************************
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

'Our message type constants
Const MSG_CHATTEXT = 0

'The main DirectX7 variable
Private mobjdx As New DirectX7
'The main DirectPlay variable
Private mobjdp As DirectPlay4
'Object used to enumerate sessions
Private mobjEnumSessions As DirectPlayEnumSessions
'Object used to enumerate connections
Private mobjEnumConnections As DirectPlayEnumConnections
'The DirectPlay assigned Player ID
Private mlngPlayerID As Long

'*****************************************************
' Purpose:  Initialize the DirectPlay object
'*****************************************************

Private Sub Class_Initialize()

    'Enable error handling
    On Error GoTo FAILED
        
    'Create a DirectPlay instance
    Set mobjdp = mobjdx.DirectPlayCreate("")

    'Exit before we hit the error code
    Exit Sub

FAILED:
    MsgBox ("Failed to initialize DirectPlay.")
    

End Sub

'*****************************************************
' Purpose:  Call the dismantle procedure
'*****************************************************

Private Sub Class_Terminate()

    Dismantle

End Sub

'*****************************************************
' Purpose:  Dismantle DirectPlay and destroy any
'           associated objects
'*****************************************************

Public Sub Dismantle()

    Set mobjEnumConnections = Nothing
    Set mobjEnumSessions = Nothing
    Set mobjdp = Nothing

End Sub

'*****************************************************
' Purpose:  Enumerate the possible connections, and
'           display them in a listbox
'*****************************************************

Public Function InitConnectionList(LstBox As ListBox) As Boolean

Dim lngNumConnections As Long   'How many different service providers are there?
Dim strName As String           'What are their names?
Dim i As Long
 
    'Handle errors
    On Error GoTo FAILED
 
    'Erase the contents of the listbox
    LstBox.Clear
  
    'Enumerate the connections
    Set mobjEnumConnections = mobjdp.GetDPEnumConnections("", DPCONNECTION_DIRECTPLAY)
  
    'Determine the number of connections available
    lngNumConnections = mobjEnumConnections.GetCount
    
    'Step through each available connection
    For i = 1 To lngNumConnections
        'Determine the connection's name
        strName = mobjEnumConnections.GetName(i)
        'List this name in a listbox
        LstBox.AddItem strName
    Next
 
    'Initialize the listbox (no item should be selected)
    LstBox.ListIndex = 0

    'We have successfully listed the connections
    InitConnectionList = True

    'Exit before we hit the error code
    Exit Function

FAILED:
    MsgBox ("Unable to list connections.")
    InitConnectionList = False

End Function

'*****************************************************
' Purpose:  Enumerate and list all of the sessions
'           available on the current connection
'*****************************************************

Public Function ListSessions(Index As Long, LstBox As ListBox) As Boolean

Dim intSessionCount As Integer              'How many sessions are available?
Dim i As Integer
Dim objSessionData As DirectPlaySessionData 'How is a session described?
Dim strDetails As String                    'A string display the session's description to the user
  
    'If we failed to make a connection, exit the subroutine
    If Not (InitConnection(Index)) Then Exit Function
  
    'Erase the contents of the listbox
    LstBox.Clear

    'Handle errors
    On Error GoTo FAILED
    
    'Make a new SessionData object
    Set objSessionData = mobjdp.CreateSessionData
    
    'Fill this object with the criterea we would like to use to descriminate sessions during enumeration
    objSessionData.SetGuidApplication APP_GUID
    objSessionData.SetSessionPassword ""
    
    'Enumerate sessions with the current connection
    Set mobjEnumSessions = mobjdp.GetDPEnumSessions(objSessionData, 0, DPENUMSESSIONS_AVAILABLE)
 
    'Get info on each session, and add a description to the listbox
    intSessionCount = mobjEnumSessions.GetCount
    For i = 1 To intSessionCount
        Set objSessionData = mobjEnumSessions.GetItem(i)
        strDetails = objSessionData.GetSessionName & " (" & objSessionData.GetCurrentPlayers & "/" & objSessionData.GetMaxPlayers & ")"
        LstBox.AddItem strDetails
    Next
  
    'Initialize the listbox (no item should be selected)
    If LstBox.ListCount > 0 Then LstBox.ListIndex = 0
    
    'We have successfully listed the sessions
    ListSessions = True
    
    'Exit before we hit the error code
    Exit Function
  
FAILED:
    MsgBox ("Unable to list available games.")
    ListSessions = False

End Function

'*****************************************************
' Purpose:  Create a new session using the currently
'           selected connection
'*****************************************************

Public Function HostGame(Index As Long, strPlayerName As String, strSessionName As String, Optional intMaxPlayers As Integer = 2) As Boolean

Dim objSessionData As DirectPlaySessionData 'How is the session described?
    
    'If we failed to make a connection, exit the subroutine
    If Not (InitConnection(Index)) Then Exit Function
    
    'Handle errors
    On Error GoTo FAILED
    
    'Make a new SessionData object
    Set objSessionData = mobjdp.CreateSessionData
    
    'Assign the maximum number of players
    objSessionData.SetMaxPlayers intMaxPlayers
    'Assign the session's name
    objSessionData.SetSessionName strSessionName
    'Set the GUID that the session will be known by
    objSessionData.SetGuidApplication APP_GUID
    'Use the directplay protocol for packet transmission, and allow migration of the host
    objSessionData.SetFlags DPSESSION_DIRECTPLAYPROTOCOL Or DPSESSION_MIGRATEHOST
    
    'Create (and join) the session.
    Call mobjdp.Open(objSessionData, DPOPEN_CREATE)
    
    'Create a player in this new session, and assign a name
    mlngPlayerID = mobjdp.CreatePlayer(strPlayerName, "", 0, 0)
    
    'We have successfully hosted the game
    HostGame = True
        
    'Exit before the error code
    Exit Function

FAILED:
    MsgBox ("Unable to create game.")
    HostGame = False

End Function

'*****************************************************
' Purpose:  Join the session currently highlighted in
'           the listbox
'*****************************************************

Public Function JoinGame(Index As Long, strPlayerName As String) As Boolean

Dim objSessionData As DirectPlaySessionData 'How is the session described?
 
    'Handle errors
    On Error GoTo FAILED
 
    'Get the session data from the selected item
    Set objSessionData = mobjEnumSessions.GetItem(Index + 1)
    
    'Join the session
    Call mobjdp.Open(objSessionData, DPOPEN_JOIN)
    
    'Create a player in this new session, and assign a name
    mlngPlayerID = mobjdp.CreatePlayer(strPlayerName, "Player", 0, 0)

    'We have successfully joined the game
    JoinGame = True

    'Exit before the error code
    Exit Function

FAILED:
    'If the error arose during player creation...
    If Err.Number = 5 Then
        MsgBox ("A player with the same name already exists in that session.")
    Else
    'Otherwise, it was a generic connection error
        MsgBox ("Unable to join game.")
    End If
    JoinGame = False

End Function

'*****************************************************
' Purpose:  Initialize a connection with the selected
'           service provider
'*****************************************************

Private Function InitConnection(Index As Long) As Boolean

Dim objDPAddress As DirectPlayAddress   'What is the address we're using with this service provider?
 
    'Handle errors
    On Error GoTo FAILED
    
    'Get the address for this connection
    Set objDPAddress = mobjEnumConnections.GetAddress(Index + 1)
    
    'Initialize this address
    Call mobjdp.InitializeConnection(objDPAddress)

    'We have successfully initialized the connection
    InitConnection = True

    'Exit before the error code
    Exit Function
 
FAILED:
    MsgBox ("Failed to initialize connection.")
    InitConnection = False
  
End Function

'*****************************************************
' Purpose:  Stop people from joining the session
'*****************************************************

Public Sub DisableJoin()

Dim objSessionData As DirectPlaySessionData

    'Make a new session data object
    Set objSessionData = mobjdp.CreateSessionData
    
    'Get the session description
    mobjdp.GetSessionDesc objSessionData
    
    'Set the joined disabled flag
    objSessionData.SetFlags objSessionData.GetFlags Or DPSESSION_JOINDISABLED
    
    'Set the session description (now that we've changed the flags)
    mobjdp.SetSessionDesc objSessionData

End Sub

'*****************************************************
' Purpose:  List all of the players currently
'           connected to the session
'*****************************************************

Public Function UpdatePlayers(LstBox As ListBox) As String
  
Dim i As Integer
Dim strPlayerName As String
Dim intPlayersWaiting As Integer
Dim objEnumPlayers As DirectPlayEnumPlayers
Dim objSessionData As DirectPlaySessionData
  
    'Handle errors
    On Error GoTo FAILED
    
    'Get the EnumPlayers object
    Set objEnumPlayers = mobjdp.GetDPEnumPlayers("", 0)
    
    'Determine how many players it contains
    intPlayersWaiting = objEnumPlayers.GetCount
  
    'Create a session data object
    Set objSessionData = mobjdp.CreateSessionData
    
    'Get the current session data
    mobjdp.GetSessionDesc objSessionData
    
    'Display the number of players in the session
    UpdatePlayers = intPlayersWaiting & " people in the session."
    
    'Erase the contents of the listbox
    LstBox.Clear
    
    'For every player in the session, extract the short name, and add to a listbox
    For i = 1 To intPlayersWaiting
        strPlayerName = objEnumPlayers.GetShortName(i)
        LstBox.AddItem strPlayerName
    Next
      
    'Exit before the error code
    Exit Function
    
FAILED:
    MsgBox ("Error listing players.")
    
End Function

'*****************************************************
' Purpose:  Check for any messages, and process those
'           that are detected
'*****************************************************

Public Sub GetMessages()

Dim lngFromPlayerID As Long     'ID of player that message is from
Dim lngToPlayerID As Long       'ID of player to whom message was sent
Dim lngMsgType As Long          'A value indicating the type of message this is
Dim dpMsg As DirectPlayMessage  'An object that contains all message data
Dim lngMsgCount As Long         'How many messages are in the queue?
Dim strFromPlayerName As String 'What is the short name of the player who sent the message?
Dim strChatText As String       'The chat text string

    'Handle errors that may occur if the session or player have not yet been created
    On Error GoTo NOTINIT
    
    'Find out how many messages are waiting for the player indicated by lngPlayerID
    lngMsgCount = mobjdp.GetMessageCount(mlngPlayerID)
      
    'Allow normal error handling to continue
    On Error GoTo 0
    
    'Process each message sequentially
    Do While lngMsgCount > 0
        'Receive the message, and assign the "To" and "From" ID's
        Set dpMsg = mobjdp.Receive(lngFromPlayerID, lngToPlayerID, DPRECEIVE_ALL)
        'Read the first LONG from the message to determine the type
        lngMsgType = dpMsg.ReadLong()
        'Decrease the number of messages by one (we've just read one)
        lngMsgCount = lngMsgCount - 1
          
        'Check if this message was sent by the system
        If lngFromPlayerID = DPID_SYSMSG Then
            'Take action depending on the type of system message sent
            Select Case lngMsgType
                'If a player has joined or left, update the listbox
                Case DPSYS_DESTROYPLAYERORGROUP, DPSYS_CREATEPLAYERORGROUP
                    'Update the listbox
                    frmStart.lblStatus.Caption = gobjDPlay.UpdatePlayers(frmStart.lstPlayers)
            End Select
        'If not sent by system, this must be a application-defined message
        Else
            'Get friendly name of sending player
            strFromPlayerName = mobjdp.GetPlayerFriendlyName(lngFromPlayerID)
            'Take action depending on the type of application message
            Select Case lngMsgType
                'If we've recieved chat text, apply it
                Case MSG_CHATTEXT
                    'Get the string
                    strChatText = dpMsg.ReadString
                    'Display it
                    frmStart.lstText.AddItem strFromPlayerName & ": " & strChatText
            End Select
        End If
    Loop
      
    'Exit before the error code
    Exit Sub

NOTINIT:
    Exit Sub

End Sub

'*****************************************************
' Purpose:  Send a "MSG_CHATTEXT" message
'*****************************************************

Public Sub SendChatText(strChatText As String)

Dim dpMsg As DirectPlayMessage  'Create a message object

    'Make a new message
    Set dpMsg = mobjdp.CreateMessage
    
    'Assign it as a initgame message
    dpMsg.WriteLong MSG_CHATTEXT
    
    'Write the chat text
    dpMsg.WriteString strChatText
    
    'Send the message
    mobjdp.SendEx mlngPlayerID, DPID_ALLPLAYERS, DPSEND_GUARANTEED, dpMsg, 0, 0, 0
    
End Sub

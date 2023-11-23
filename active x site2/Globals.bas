Attribute VB_Name = "Globals"
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

'The global DirectPlay object we'll continually use
Global gobjDPlay As DirPlay

'Our GUID
Global Const APP_GUID = "{64823540-E39A-11D3-9651-0080C8E4028D}"

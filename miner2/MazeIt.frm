VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "MazeIt"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   5880
      Picture         =   "MazeIt.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   5280
      Picture         =   "MazeIt.frx":0C42
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   4680
      Picture         =   "MazeIt.frx":1884
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   4080
      Picture         =   "MazeIt.frx":24C6
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   3480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2880
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   2280
      Picture         =   "MazeIt.frx":3108
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   1680
      Picture         =   "MazeIt.frx":3D4A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   8160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      Picture         =   "MazeIt.frx":498C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngPlayerXPos As Long 'dim a long to hold the X position of the player
Dim lngPlayerYPos As Long 'dim a long to hold the Y position of the player
Dim lngOldPlayerXPos As Long 'dim a long to hold the previous X position of the player
Dim lngOldPlayerYPos As Long 'dim a long to hold the previous Y position of the player
Dim hCleanMapDC As Long 'dim a handle to a DC for holding a clean piece of the map
Dim hCleanMapBitmap As Long 'handle to a Bitmap to be selected into the dc
Dim hPrevCleanMapBitmap As Long 'the handle to the old bitmap that needs to be reselected to delete the object

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    lngOldPlayerXPos = lngPlayerXPos 'store the current X position into a variable before updating
    lngOldPlayerYPos = lngPlayerYPos 'store the current Y position into a variable before updating
    Select Case KeyCode 'find out which key was pressed
        Case vbKeyUp 'the up arrow was pressed
            lngPlayerYPos = lngPlayerYPos - 32 'update the player location
        Case vbKeyDown ' the down arrow was pressed
            lngPlayerYPos = lngPlayerYPos + 32 'update the player location
        Case vbKeyLeft 'the left arrow was pressed
            lngPlayerXPos = lngPlayerXPos - 32 'update the player location
        Case vbKeyRight 'the right arrow was pressed
            lngPlayerXPos = lngPlayerXPos + 32 'update the player location
        Case Else 'no valid key was pressed, so
            Exit Sub 'exit the sub
    End Select
    Call BlitPlayer(lngPlayerXPos, lngPlayerYPos) 'call bitplayer with the updated X and Y player position
End Sub

Private Sub Form_Load()
    Dim strMapName As String 'dim a string variable to hold the map name
    Dim intFileFree As Integer 'dim an integer to hold a handle to the next free file
    Dim lngRecordNumber As Long 'dim a variable to track what record we are on
    Dim lngRecordLength As Long 'dim a variable to hold the length of a record
    Dim lngYCount As Long 'dim a variable for the Y coordinate loop
    Dim lngXCount As Long 'dim a variable for the X coordinate loop
    Dim LoadMap As typeTileStructure  'dim a variable that will be used to determine the length of a record in the file

    'Create an area to hold a clean section of the map to be blitted back when the player moves
    hCleanMapDC = CreateCompatibleDC(Form1.hdc)
    hCleanMapBitmap = CreateCompatibleBitmap(Form1.hdc, 32, 32)
    hPrevCleanMapBitmap = SelectObject(hCleanMapDC, hCleanMapBitmap)
    
    '***PLACE your own map file name here to use it***
    'strMapName = App.Path & "\examplemap.map" 'get the file name so it can be used to save the map
     strMapName = App.Path & "\gg.map"
    intFileFree = FreeFile 'get the next available free file handle

    'opens a file for binary input, using the map name as the file name, and the available free file
    Open strMapName For Binary As #intFileFree
    
    lngRecordNumber = 1 'set the first record
    lngRecordLength = CLng(Len(LoadMap)) 'sets the length of a record to equal the length of our user-defined type
    For lngXCount = 0 To NUMBER_X_TILES  'loop through the first element in our array
        For lngYCount = 0 To NUMBER_Y_TILES 'loop through the second element in our array
            Get #intFileFree, lngRecordNumber, Tile(lngXCount, lngYCount) 'retrieve this tiles' information in this recordnumber
            lngRecordNumber = lngRecordNumber + lngRecordLength 'increment the recordnumber by the recordlength to the next recordnumber
        Next
    Next
    Close #intFileFree 'close the open file
    Call InitializeMap 'call the subroutine which draws the loaded map to the picturebox
    
    '***ENTER your starting coordinates here***
    lngPlayerXPos = 192 'set the starting X coordinate for the player
    lngPlayerYPos = 128 'set the starting Y coordinate for the player
    Call BlitPlayer(lngPlayerXPos, lngPlayerYPos) 'call the subroutine that draws the player into place

    Exit Sub

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Delete any objects and DC's created
    hCleanMapBitmap = SelectObject(hCleanMapDC, hPrevCleanMapBitmap)
    DeleteObject hCleanMapBitmap
    DeleteDC hCleanMapDC
End Sub
Private Sub InitializeMap()
    Dim lngYCount As Long 'dim a variable for the Y coordinate loop
    Dim lngXCount As Long 'dim a variable for the X coordinate loop
    Dim lngXPos As Long 'X coordinate for the tile location in the picturebox
    Dim lngYPos As Long 'Y coordinate for the tile location in the picturebox
    
    For lngXCount = 0 To NUMBER_X_TILES  'loop through the first element in our array
        For lngYCount = 0 To NUMBER_Y_TILES 'loop through the second element in our array
            lngXPos = lngXCount * 32 'multiply the X count by 32 to get the X location the tile will be blitted to
            lngYPos = lngYCount * 32 'multiply the Y count by 32 to get the Y location the tile will be blitted to
            If Tile(lngXCount, lngYCount).lngTileImgNumber <> 0 Then 'if a tile exists here, blit it into place
                'blit the tile to the appropriate coordinate in the picturebox
                BitBlt Form1.hdc, lngXPos, lngYPos, 32, 32, Picture2(Tile(lngXCount, lngYCount).lngTileImgNumber).hdc, 0, 0, SRCCOPY
            End If
        Next
    Next

    Form1.Refresh 'refresh to reflect the changes made

End Sub
Private Sub BlitPlayer(lngXPos As Long, lngYPos As Long)
    '****************************************************************************************************************
    'This subroutine takes an X and Y coordinate and blits the player to it, as long as the player is moving to a
    'valid tile. It also copies a clean part of the map in memory to the previous player location, and then grabs the
    'clean section the player is moving onto and stores it in memory also.
    '****************************************************************************************************************
    Static boolStarted As Boolean 'dim a static variable so we know that something is in the memory DC
    
    If boolStarted Then 'if this routine has been accessed before
        'then copy a clean section of the map back to the old player location
        BitBlt Form1.hdc, lngOldPlayerXPos, lngOldPlayerYPos, 32, 32, hCleanMapDC, 0, 0, SRCCOPY
    Else 'otherwise,
        boolStarted = True 'let the program know that we now have something in the memory DC
    End If
    
    'if the tile the player is trying to step on is walkable then
    If Tile(lngXPos \ 32, lngYPos \ 32).lngTileType = WALKABLE Then
        'blit the new section into the memoryDC, and then
        BitBlt hCleanMapDC, 0, 0, 32, 32, Form1.hdc, lngXPos, lngYPos, SRCCOPY
        'blit the player sprite to the new location
        BitBlt Form1.hdc, lngXPos, lngYPos, 32, 32, picSprite.hdc, 32, 0, SRCAND
        BitBlt Form1.hdc, lngXPos, lngYPos, 32, 32, picSprite.hdc, 0, 0, SRCPAINT
    Else 'if the tile is not walkable, then
        lngPlayerXPos = lngOldPlayerXPos 'keep the player in the X location he is in
        lngPlayerYPos = lngOldPlayerYPos 'keep the player in the Y location he is in
        'reblit the player to the old position
        BitBlt Form1.hdc, lngPlayerXPos, lngPlayerYPos, 32, 32, picSprite.hdc, 32, 0, SRCAND
        BitBlt Form1.hdc, lngPlayerXPos, lngPlayerYPos, 32, 32, picSprite.hdc, 0, 0, SRCPAINT
    End If
    Form1.Refresh 'reflect any changes made to the form
End Sub


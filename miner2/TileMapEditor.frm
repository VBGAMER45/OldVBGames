VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Tile Map Editor"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   11730
      Picture         =   "TileMapEditor.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   2430
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
      Left            =   11220
      Picture         =   "TileMapEditor.frx":0C42
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   2400
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
      Left            =   10710
      Picture         =   "TileMapEditor.frx":1884
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   2400
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
      Left            =   10200
      Picture         =   "TileMapEditor.frx":24C6
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Non-Walkable Tiles"
      Height          =   375
      Left            =   10440
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Walkable Tiles"
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.PictureBox picOutline 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   11160
      Picture         =   "TileMapEditor.frx":3108
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   4920
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
      Left            =   11730
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   2430
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
      Left            =   11220
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   2430
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
      Left            =   10710
      Picture         =   "TileMapEditor.frx":34AD
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   2430
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
      Left            =   10200
      Picture         =   "TileMapEditor.frx":40EF
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   2430
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   120
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tile Type:"
      Height          =   195
      Left            =   10320
      TabIndex        =   16
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label lblTileType 
      AutoSize        =   -1  'True
      Caption         =   "Non-Walkable"
      Height          =   195
      Left            =   11100
      TabIndex        =   15
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label lblArrayNumber 
      AutoSize        =   -1  'True
      Caption         =   "Number of Array"
      Height          =   195
      Left            =   11160
      TabIndex        =   10
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Array Number:"
      Height          =   195
      Left            =   10080
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.Line Line4 
      X1              =   680
      X2              =   680
      Y1              =   161
      Y2              =   194
   End
   Begin VB.Line Line3 
      X1              =   816
      X2              =   816
      Y1              =   160
      Y2              =   193
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   681
      X2              =   814
      Y1              =   194
      Y2              =   194
   End
   Begin VB.Line Line2 
      X1              =   680
      X2              =   680
      Y1              =   192
      Y2              =   160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   681
      X2              =   814
      Y1              =   161
      Y2              =   161
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Currently selected tile"
      Height          =   195
      Left            =   10320
      TabIndex        =   6
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   10560
      Picture         =   "TileMapEditor.frx":4D33
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   960
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuSaveMap 
         Caption         =   "Save Map"
      End
      Begin VB.Menu menuLoadMap 
         Caption         =   "Load Map"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hWorkAreaDC As Long 'handle to a DC for a work area we will be needing
Dim hWorkAreaBitmap As Long 'handle to a compatible bitmap image for our work area
Dim hPrevWorkAreaBitmap As Long 'handle to the bitmap that will be selected back so we can delete the object
Dim hOldOutlineDC As Long 'handle to a DC for the old outline area that will be drawn to
Dim hOldOutlineBitmap As Long 'handle to a compatible bitmap for the outline area
Dim hPrevOldOutlineBitmap As Long 'handle to the bitmap that will be selected back so we can delete the object
Dim lngTileType As Long 'holds which tile type is being displayed
Dim lngTileImgNumber As Long 'holds which tile picture is currently selected
Dim boolUpdateOccurred As Boolean 'this indicates whether or not drawing has occured, to prevent the old outline from erasing any drawing
Private Sub Form_Load()
    'Create Work Area DC
    hWorkAreaDC = CreateCompatibleDC(Form1.hdc)
    hWorkAreaBitmap = CreateCompatibleBitmap(Form1.hdc, 32, 32)
    hPrevWorkAreaBitmap = SelectObject(hWorkAreaDC, hWorkAreaBitmap)
    'Create area to hold the old outline location bitmap
    hOldOutlineDC = CreateCompatibleDC(Form1.hdc)
    hOldOutlineBitmap = CreateCompatibleBitmap(Form1.hdc, 32, 32)
    hPrevOldOutlineBitmap = SelectObject(hOldOutlineDC, hOldOutlineBitmap)
    
    lngTileImgNumber = 1 'the app starts with tile img 1 loaded
    lngTileType = NONWALKABLE 'the app starts with non-walkable tiles being displayed
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Delete work wrea storage area
    hWorkAreaBitmap = SelectObject(hWorkAreaDC, hPrevWorkAreaBitmap)
    DeleteDC hWorkAreaBitmap
    DeleteObject hWorkAreaBitmap
    'Delete old outline storage area
    hOldOutlineBitmap = SelectObject(hOldOutlineDC, hPrevOldOutlineBitmap)
    DeleteObject hOldOutlineBitmap
    DeleteDC hOldOutlineDC
End Sub
Private Sub Command1_Click()
    Dim lngCount As Long 'dim a long variable used for the loop
    lngTileType = WALKABLE 'switch to walkable tiles
    For lngCount = 1 To 5 'make arrays 1 to 4 of the tile pictureboxes invisible
        Picture2(lngCount).Visible = False
    Next
    For lngCount = 5 To 8 'make arrays 5 to 8 of the tile pictureboxes visibe
        Picture2(lngCount).Visible = True
    Next
    lblTileType = "Walkable" 'change the label caption so the user knows what tile type is being displayed
End Sub
Private Sub Command2_Click()
    Dim lngCount As Long 'dim a long variable used for the loop
    lngTileType = NONWALKABLE 'switch to walkable tiles
    For lngCount = 1 To 5 'make arrays 1 to 4 of the tile pictureboxes visible
        Picture2(lngCount).Visible = True
    Next
    For lngCount = 5 To 8 'make arrays 5 to 8 of the tile pictureboxes invisibe
        Picture2(lngCount).Visible = False
    Next
    lblTileType = "Non-Walkable" 'change the label caption so the user knows what tile type is being displayed
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngPositionX As Long, lngPositionY As Long 'dim an X and Y position
    lngPositionX = (X \ 32) * 32 'take the X parameter from Picture1_MouseDown and bump it to the closest available 32 pixel location, assigning it to lngPositionX
    lngPositionY = (Y \ 32) * 32 'take the Y parameter from Picture1_MouseDown and bump it to the closest available 32 pixel location, assigning it to lngPositionY
    If Button = 1 Then 'if the left button was clicked, then
        Call AddTile(lngPositionX, lngPositionY) 'call AddTile with lngPositionX and lngPositionY
    ElseIf Button = 2 Then 'if the right button is down then
        Call DeleteTile(lngPositionX, lngPositionY) 'call DeleteTile with the X and Y coordinates
    End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngPositionX As Long, lngPositionY As Long 'dim an X and a Y position
    
    'take the X parameter from Picture1_MouseMove and bump it to the closest available 32 pixel location,
    'assigning it to lngPositionX
    lngPositionY = (Y \ 32) * 32
    
    'take the Y parameter from Picture1_MouseMove and bump it to the closest available 32 pixel location,
    'assigning it to lngPositionY
    lngPositionX = (X \ 32) * 32
    
    'update the Array Number label with the current array associated with this position on the map. This is useful
    'when you are planning on the starting position for the player, or any other kind of hard-coded interactions.
    lblArrayNumber = lngPositionX \ 32 & "," & lngPositionY \ 32
    
    If Button = 0 Then 'if no button is down then
        Call DrawOutline(lngPositionX, lngPositionY) ' call DrawOutline with lngPositionX and lngPositionY
    ElseIf Button = 1 Then 'if the left button is down then
        Call AddTile(lngPositionX, lngPositionY) 'call AddTile with lngPositionX and lngPositionY
    ElseIf Button = 2 Then 'if the right button is down then
        Call DeleteTile(lngPositionX, lngPositionY) 'call DeleteTile with the X and Y coordinates
    End If
End Sub
Private Sub Picture2_Click(Index As Integer)
    'take the selected picturebox, and load into image1. This will provide the user with a larger image of the tile.
    'This comes in handy when tiles are small, and the differences between tiles is subtle.
    Image1.Picture = Picture2(Index).Picture
    
    lngTileImgNumber = Index 'store which tile picture number is selected into lngTileImgNumber
    
    'blit the selected tile picture into the TilePicture DC
    BitBlt Picture2(lngTileImgNumber).hdc, 0, 0, 32, 32, Picture2(Index).hdc, 0, 0, SRCCOPY
    
End Sub
Private Sub menuLoadMap_Click()
    Dim LoadMap As typeTileStructure  'dim a variable that will be used to determine the length of a record in the file
    Dim strMapName As String 'dim a variable to hold the name of the map
    Dim intFileFree As Integer 'dim a variable to hold a handle to the next available free file
    Dim lngRecordNumber As Long 'dim a variable to track what record we are on
    Dim lngRecordLength As Long 'dim a variable to hold the length of a record
    Dim lngYCount As Long 'dim a variable for the Y coordinate loop
    Dim lngXCount As Long 'dim a variable for the X coordinate loop
    
    On Error GoTo ErrorHandler 'setup error handling in case cancel is selected

    'these lines initialize a common dialog control
    CommonDialog1.Filter = "Map files (*.map)|*.map"
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.DialogTitle = "Load Map"
    CommonDialog1.ShowOpen
    
    strMapName = CommonDialog1.filename 'get the file name so it can be used to save the map
    If strMapName = "" Then Exit Sub 'if no name was entered, exit the sub
    
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

    Exit Sub
'these last lines handle an error generated by the user clicking cancel in the common dialog control
ErrorHandler:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If
End Sub
Private Sub menuSaveMap_Click()
    Dim SaveMap As typeTileStructure 'dim a variable that will be used to determine the length of a record in the file
    Dim strMapName As String 'dim a variable to hold the name of the map
    Dim intFileFree As Integer 'dim a variable to hold a handle to the next available free file
    Dim lngRecordNumber As Long 'dim a variable to track what record we are on
    Dim lngRecordLength As Long 'dim a variable to hold the length of a record
    Dim lngYCount As Long 'dim a variable for the Y coordinate loop
    Dim lngXCount As Long 'dim a variable for the X coordinate loop
    
    On Error GoTo ErrorHandler 'setup error handling in case cancel is selected

    'these lines initialize a common dialog control
    CommonDialog1.Filter = "Map files (*.map)|*.map"
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.DialogTitle = "Save Map"
    CommonDialog1.ShowSave
    
    strMapName = CommonDialog1.filename 'get the file name so it can be used to save the map
    If strMapName = "" Then Exit Sub 'if no name was entered, exit the sub
    
    intFileFree = FreeFile 'get the next available free file handle
    
    'opens a file for binary input, using the map name as the file name, and the available free file
    Open strMapName For Binary As #intFileFree
    
    lngRecordNumber = 1 'set the first record
    lngRecordLength = CLng(Len(SaveMap)) 'sets the length of a record to equal the length of our user-defined type
    For lngXCount = 0 To NUMBER_X_TILES  'loop through the first element in our array
        For lngYCount = 0 To NUMBER_Y_TILES 'loop through the second element in our array
            Put #intFileFree, lngRecordNumber, Tile(lngXCount, lngYCount) 'record this tiles' information in this recordnumber
            lngRecordNumber = lngRecordNumber + lngRecordLength 'increment the recordnumber by the recordlength to the next recordnumber
        Next
    Next
    
    Close #intFileFree 'close the open file
    
    MsgBox "Map saved" 'let the user know that his map was saved successfully

    Exit Sub
'these last lines handle an error generated by the user clicking cancel in the common dialog control
ErrorHandler:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If
End Sub
Private Sub AddTile(lngXPos As Long, lngYPos As Long)
    '*******************************************************************************************************
    'This sub takes an X and Y position, and adds the tile to the Tile array with all neccessary information
    '*******************************************************************************************************
    Dim lngModifiedXPos As Long 'stores the X position after we perform some math to it
    Dim lngModifiedYPos As Long 'stores the Y position after we perform some math to it
    
    lngModifiedXPos = lngXPos \ 32 'modulus lngXPos to get the first dimension of our tile array
    lngModifiedYPos = lngYPos \ 32 'modulus lngYPos to get the second dimension of our tile array
    
    'check to make sure the modified X and Y positions fall within the array boundaries. If they don't, then exit the sub
    If lngModifiedXPos < 0 Or lngModifiedXPos > 19 Or lngModifiedYPos < 0 Or lngModifiedYPos > 19 Then Exit Sub
    
    Tile(lngModifiedXPos, lngModifiedYPos).lngTileType = lngTileType   'update the tile with the type of tile it is (i.e. walkable, non-walkable)
    Tile(lngModifiedXPos, lngModifiedYPos).lngTileImgNumber = lngTileImgNumber  'update the tile with the currently selected tile picture number
    BitBlt Picture1.hdc, lngXPos, lngYPos, 32, 32, Picture2(lngTileImgNumber).hdc, 0, 0, SRCCOPY 'blit the selected tile picture onto picture1
    Picture1.Refresh 'refresh the picturebox to reflect update
    boolUpdateOccurred = True 'let DrawOutline know that an update occurred
End Sub
Private Sub DeleteTile(lngXPos As Long, lngYPos As Long)
    '*************************************************************************************************************
    'This sub takes an X and Y position, and deletes the tile, updating Tile array with all neccessary information
    '*************************************************************************************************************
    Dim lngModifiedXPos As Long 'stores the X position after we perform some math to it
    Dim lngModifiedYPos As Long 'stores the Y position after we perform some math to it
    
    lngModifiedXPos = lngXPos \ 32 'modulus lngXPos to get the first dimension of our tile array
    lngModifiedYPos = lngYPos \ 32 'modulus lngYPos to get the second dimension of our tile array
    
    'check to make sure the modified X and Y positions fall within the array boundaries. If they don't, then exit the sub
    If lngModifiedXPos < 0 Or lngModifiedXPos > 19 Or lngModifiedYPos < 0 Or lngModifiedYPos > 19 Then Exit Sub
    
    'if the tile array at this location doesn't have an image number associated with it, exit the sub
    If Tile(lngModifiedXPos, lngModifiedYPos).lngTileImgNumber = 0 Then
        Exit Sub
    End If
    
    'all four lines of code here reset the values of the tile array at this location
    Tile(lngModifiedXPos, lngModifiedYPos).lngTileType = 0
    Tile(lngModifiedXPos, lngModifiedYPos).lngTileImgNumber = 0
    
    'blit white onto picture1 to reflect deletion of that tile
    BitBlt Picture1.hdc, lngXPos, lngYPos, 32, 32, Picture2(lngTileImgNumber).hdc, 0, 0, WHITENESS
    Picture1.Refresh 'refresh the picturebox to reflect update
    boolUpdateOccurred = True 'let DrawOutline know that an update occurred

End Sub
Private Sub DrawOutline(lngXPos As Long, lngYPos As Long)
    '****************************************************************************************************************
    'This sub stores the piece of the picturebox that the outline gets blitted to so that it can be replaced once the
    'mouse moves off of this location and also draws the outline
    '****************************************************************************************************************
    
    'Dim 2 longs to hold the X and Y positions to be blitted when the mouse moves off the tile, and make them
    'static so their value is preserved. Also, skip blitting the old location the first time the mouse is moved
    'over the picturebox, since nothing is contained in the picturebox when the app first starts. This is done using
    'a static boolean variable named boolStarted
    Static lngOldXPos As Long, lngOldYPos As Long, boolStarted As Boolean
    
    'if the app started, and drawing hasn't occured, then blit the oldoutline info to the previous X and Y position
    If boolStarted And boolUpdateOccurred = False Then
        BitBlt Picture1.hdc, lngOldXPos, lngOldYPos, 32, 32, hOldOutlineDC, 0, 0, SRCCOPY
    Else
        'otherwise, let the subroutine know that we have information located in the OldOutlineDC now, and
        'that drawing has not occured
        boolStarted = True
        boolUpdateOccurred = False
    End If
    
    lngOldXPos = lngXPos 'store the current X in lngOldXPos
    lngOldYPos = lngYPos 'store the current Y in lngOldYPos
    
    'blit into hOldOutlineDC the location that the mouse is over right now, and make it 32x32 pixels wide and high
    BitBlt hOldOutlineDC, 0, 0, 32, 32, Picture1.hdc, lngXPos, lngYPos, SRCCOPY
    
    'blit the info from hOldOutlineDC into hWorkAreaDC so we can perform some raster ops on it
    BitBlt hWorkAreaDC, 0, 0, 32, 32, hOldOutlineDC, 0, 0, SRCCOPY
    
    'blit the outline bitmap into the work area, inverting it. This will make it so any colors contained in the work
    'area will be inverted when it comes into contact with any non-zero colors in the outline. Since the only colors
    'outline are the dashes, only the dashes appear.
    BitBlt hWorkAreaDC, 0, 0, 32, 32, picOutline.hdc, 0, 0, SRCINVERT
    
    'Finally, blit the work area to the mouse location, and refresh the picture
    BitBlt Picture1.hdc, lngXPos, lngYPos, 32, 32, hWorkAreaDC, 0, 0, SRCCOPY
    Picture1.Refresh
End Sub
Private Sub InitializeMap()
    Dim lngYCount As Long 'dim a variable for the Y coordinate loop
    Dim lngXCount As Long 'dim a variable for the X coordinate loop
    Dim lngXPos As Long 'X coordinate for the tile location in the picturebox
    Dim lngYPos As Long 'Y coordinate for the tile location in the picturebox
    
    Picture1.Cls 'clear anything that is currently in the picturebox
    For lngXCount = 0 To NUMBER_X_TILES  'loop through the first element in our array
        For lngYCount = 0 To NUMBER_Y_TILES 'loop through the second element in our array
            lngXPos = lngXCount * 32 'multiply the X count by 32 to get the X location the tile will be blitted to
            lngYPos = lngYCount * 32 'multiply the Y count by 32 to get the Y location the tile will be blitted to
            If Tile(lngXCount, lngYCount).lngTileImgNumber <> 0 Then 'if a tile exists here, blit it into place
                'blit the tile to the appropriate coordinate in the picturebox
                BitBlt Picture1.hdc, lngXPos, lngYPos, 32, 32, Picture2(Tile(lngXCount, lngYCount).lngTileImgNumber).hdc, 0, 0, SRCCOPY
            End If
        Next
    Next

    Picture1.Refresh 'refresh to reflect the changes made

End Sub

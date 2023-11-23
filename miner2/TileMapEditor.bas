Attribute VB_Name = "Module1"
Option Explicit

'Setup constants
Public Const WALKABLE = 1   'this will be a walkable tile
Public Const NONWALKABLE = 0 'this will be a non-walkable tile
Public Const NUMBER_X_TILES = 19 'Constant to keep track of the number of tiles in the X direction
Public Const NUMBER_Y_TILES = 19 'Constant to keep track of the number of tiles in the Y direction

'Start a type-defintion
Public Type typeTileStructure
    lngTileType As Long 'The type of tile the tile is (walkable/non-walkable)
    lngTileImgNumber As Long 'stores which tile picture number this tile is associated with
End Type

'Initialize the tiles, using a 20X20 map grid
Public Tile(19, 19) As typeTileStructure

'************************************
' set up bit block transfers
'************************************
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
Public Const PATCOPY = &HF00021
Public Const PATPAINT = &HFB0A09
Public Const PATINVERT = &H5A0049
Public Const DSTINVERT = &H550009
Public Const BLACKNESS = &H0
Public Const WHITENESS = &HFFFFFF
Public Const PIXELS = 3

'*****************
'Setup Memory DC's
'*****************
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

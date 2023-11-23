Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpszDestinationString1 As Any, ByVal lpszSourceString2 As Any) As Long
Global KeyStates(250) As Boolean
Public View As ViewProperties
Global Const DIRECTION_UP = 1
Global Const DIRECTION_RIGHT = 2
Global Const DIRECTION_DOWN = 3
Global Const DIRECTION_LEFT = 4

Public Const CONTROLTYPE_PICTUREBOX = 1
Public Const CONTROLTYPE_BUTTONLARGE = 2
Public Const CONTROLTYPE_LABEL = 3
Public Const CONTROLTYPE_TEXTBOX = 4
'End Type
'Interrupt Reasons
Type tMap
    FileName As String               ' Full name and path of map
    CurrentX As Byte                 ' Current Position
    CurrentY As Byte                 '
    ExtentX As Byte                  ' Map dimension
    ExtentY As Byte                  '
    DisplayExtentX As Byte           ' Number of tiles to display
    DisplayExtentY As Byte           ' Number of tiles to display
   ' Squares() As tMapSquare          ' Squares
  '  AniLookup(99) As tTileAniLookup  ' Index of animated tiles currently visible
End Type
Public Map As tMap
Type ViewProperties
  MaxScrollSpeed As Integer
  ScrollSpeedEW As Integer
  ScrollSpeedNS As Integer
  Left As Integer
  Top As Integer
  MaxLeft As Integer
  MaxTop As Integer
  Width As Integer
  Height As Integer
End Type
Public Sub TIMER()
If Form5.year.Caption = form2.year.Caption Then
form2.Timer1.Interval = 300 And Form5.Timer1.Interval = 300
'year = year + 10
End If

End Sub

Public Sub CHECKED()
If Form7.Option5.Value = True Then
Form7.Text1 = form2.Label3.Caption
End If

End Sub
'Public Sub TEXT1()
'If Form7.TEXT1 = form2.Label3.Caption Then
'form2.La
'End Sub
'bel3.Caption = Form7.TEXT1
'End If
Public Sub TIME1()
form2.Timer1.Interval = 800

End Sub
Public Sub load()

Form4.Label5.Caption = Form11.Label9.Caption
Dim form2 As New form2
form2.Show vbModal ', Me

End Sub
Public Sub tree()
'form2.Image30.Picture = LoadPicture("c:\windows\desktop\vb projects\war\tr1.bmp")
End Sub
Public Sub TIME2()
form2.Timer1.Interval = 300
End Sub



Public Sub TIME3()
form2.Timer1.Interval = 100

End Sub
Public Sub sell()

If form2.Image43.Height = 750 And form2.Image43.Width = 1095 Then
'Image43.ToolTipText = "War Factory"
 
form2.money.Caption = form2.money.Caption + 500
form2.Power.Caption = form2.Power.Caption + 20

End If
 
 
 
End Sub

Public Sub WIN()
Form9.Show vbModal, Me

Form9.Label1.Caption = MISSION2

End Sub
Public Sub b()
If form2.year.Caption >= 1990 Then
form2.Image15.Visible = True
form2.Image9.Visible = False
End If
'If form2.year.Caption >= 2050 Then
'form2.Image16.Visible = True
''End If

End Sub
Sub DrawMap()

    Dim nRunX As Integer, nRunY As Integer          ' Variables to run through the map squares to be drawn
    
    ' Run through the visible map area
    For nRunX = 0 To Map.DisplayExtentX
        For nRunY = 0 To Map.DisplayExtentY
            'Call DrawSquare(nRunX, nRunY)
        Next
    Next

End Sub
Public Sub DrawRadar()

    Dim nRunX As Integer, nRunY As Integer  ' Variables to run through all squares
    
    ' Draw overview using API SetPixelV
    For nRunX = 0 To Map.ExtentX
        For nRunY = 0 To Map.ExtentY
           ' SetPixelV Form5.picRadar.hdc, nRunX + 3, nRunY + 3, TileSet.Tile(Map.Squares(nRunX, nRunY).TileIndex).RadarColor
        Next
    Next

    ' Draw visible rectangle using standard line
    Form5.picRadar.Line (Map.CurrentX + 3, Map.CurrentY + 3)-(Map.CurrentX + 3 + Map.DisplayExtentX, Map.CurrentY + 3 + Map.DisplayExtentY), RGB(250, 250, 250), B
    
End Sub

'jonathan vaLentin 1999

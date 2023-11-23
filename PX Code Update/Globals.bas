Attribute VB_Name = "modGlobals"
'****************************************************************
'Windows API/Global Declarations for :FreeDiskSpace
'****************************************************************
Private Declare Function GetDiskFreeSpace Lib "kernel32" _
                         Alias "GetDiskFreeSpaceA" _
                         (ByVal lpRootPathName As String, _
                          lpSectorsPerCluster As Long, _
                          lpBytesPerSector As Long, _
                          lpNumberOfFreeClusters As Long, _
                          lpTotalNumberOfClusters As Long) As Long

Public Enum ActionType
    Action_Trap_Single
    ACTION_TRAP_MULTI
    ACTION_TELEPORT
    ACTION_HEAL
    ACTION_LEVEL_CHANGE
    ACTION_SPAWN_MONSTER
    ACTION_SPAWN_ITEM
    ACTION_LOCK
    ACTION_UNLOCK
    ACTION_MESSAGE
    ACTION_PLAYER_START
End Enum
Const Max_AniTiles = 20 'Max amount of animated tiles on the screen
Const MAXSpells = 30
Const Max_Traps = 30
Const Max_Lights = 10 'These are the draymicly created lights in c1.map
Global TrapsArray(Max_Traps) As New clsAction


Global ShopWindowOpen As Boolean 'Is the shop window visible or not?

'Posting constants
Public Const MAX_POSTS_POSTS = 999
Global Messageboard(1 To MAX_POSTS_POSTS) As Postings
'Messageboard postings
Type Postings
    Subject As String
    Post As String
    Author As String
End Type
Type HaveSkills
Axe As Boolean
Mace As Boolean
Bow As Boolean
CastSpell As Boolean
MeleeDefense As Boolean
MagicDefense As Boolean
Scan As Boolean
Bombs As Boolean
Stealth As Boolean
Staff As Boolean
Spear As Boolean
Sword As Boolean
unarmed As Boolean
Force As Boolean
Fail As Boolean
Miner As Boolean
Dagger As Boolean
Blacksmith As Boolean
End Type
'Holds wether player has trained the skill or not
Global TrueSkills As HaveSkills

Global CurrentPost As Integer
Global MapEd As Boolean 'Are you in the map editor or not
Global ToggleTileType As Boolean

Type SpellsType
Spellname As String
Mana As Integer
AddHealth As Integer
AddStamina As Integer
AddMana As Integer
AddFood As Integer
Posion As Integer
SkillBonus As Integer
CSskill As Integer
Graphic As String
MaxDamage As Integer
MinDamage As Integer
EffectArea As Integer ' The area the spell effect 1 = 1tile
Target As String 'monster/other player or yourself where the spell is casted
Duration As Integer
Researched As Boolean
End Type
Global Spells(MAXSpells) As SpellsType
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
' delete a dc
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'load an image into an object (&H10 for load from file)
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'delete an object from memory
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'link to an object
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'gets the color of a pixel from a certain source and location
'we use this for pixel collision detection
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public fps As Long

Public RepeatTimes As Long 'This one calculates,
' but don't use in this sample. If You need, You
' can add simple checking at Timer1_Timer Procedure
Public TotalFrames As Long
Declare Function RegComCtl32 Lib "actskin4.OCX" _
Alias "DllRegisterServer" () As Long
Declare Function UnRegComCtl32 Lib "actskin4.OCX" _
Alias "DllUnregisterServer" () As Long
Const ERROR_SUCCESS = &H0
Global SkinFilename As String
Public Function LoadGif(sFile As String, aImg As Variant) As Boolean
    LoadGif = False
    If Dir$(sFile) = "" Or sFile = "" Then
       MsgBox "File " & sFile & " not found", vbCritical
       Exit Function
    End If
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim imgHeader As String, fileHeader As String
    Dim buf$, picbuf$
    Dim imgCount As Integer
    Dim I&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd As String
    GifEnd = Chr(0) & Chr(33) & Chr(249)
    For I = 1 To aImg.Count - 1
        Unload aImg(I)
    Next I
    
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
        buf = String(LOF(fNum), Chr(0))
        Get #fNum, , buf 'Get GIF File into buffer
    Close fNum
    
    I = 1
    imgCount = 0
    j = InStr(1, buf, GifEnd) + 1
    fileHeader = Left(buf, j)
    If Left$(fileHeader, 3) <> "GIF" Then
       MsgBox "This file is not a *.gif file", vbCritical
       Exit Function
    End If
    LoadGif = True
    I = j + 2
    If Len(fileHeader) >= 127 Then
        RepeatTimes& = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * 256&)
    Else
        RepeatTimes = 0
    End If

    Do ' Split GIF Files at separate pictures
       ' and load them into Image Array
        imgCount = imgCount + 1
        j = InStr(I, buf, GifEnd) + 3
        If j > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
                picbuf = String(Len(fileHeader) + j - I, Chr(0))
                picbuf = fileHeader & Mid(buf, I - 1, j - I)
                Put #fNum, 1, picbuf
                imgHeader = Left(Mid(buf, I - 1, j - I), 16)
            Close fNum
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256&)) * 10&
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&)
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256&)
                Load aImg(imgCount - 1)
                aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
                aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
            End If
            ' Use .Tag Property to save TimeWait interval for separate Image
            aImg(imgCount - 1).Tag = TimeWait
            aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
            Kill ("temp.gif")
            I = j
        End If
        DoEvents
    Loop Until j = 3
' If there are one more Image - Load it
    If I < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
            picbuf = String(Len(fileHeader) + Len(buf) - I, Chr(0))
            picbuf = fileHeader & Mid(buf, I - 1, Len(buf) - I)
            Put #fNum, 1, picbuf
            imgHeader = Left(Mid(buf, I - 1, Len(buf) - I), 16)
        Close fNum
        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256)
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256)
            Load aImg(imgCount - 1)
            aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
            aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(imgCount - 1).Tag = TimeWait
        aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
        Kill ("temp.gif")
    End If
    TotalFrames = aImg.Count - 1
    Exit Function
ErrHandler:
    MsgBox "Error No. " & Err.Number & " when reading file", vbCritical
    LoadGif = False
    On Error GoTo 0
End Function

Public Function FitText(ByRef Ctl As Control, _
                        ByVal strCtlCaption) As String

' Function FitText
' Author:   Jeff Cockayne
'
' Fit the caption text passed in strCtlCaption
' to the width of the passed Control, Ctl.
' There are a few ways to blow this function, like
' passing a control without a Caption Property, but
' this Function is for internal use, so...
'
' Example:
' If "C:\Program Files\Test.TXT" was too wide, the
' returned string might be: "C:\Pro...\Test.TXT"

Dim lngCtlLeft As Long
Dim lngMaxWidth As Long
Dim lngTextWidth As Long
Dim lngX As Long

' Store frequently referenced values to increase
' performance (saves some OLE lookup)
lngCtlLeft = Ctl.Left
lngMaxWidth = Ctl.Width
lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)


lngX = (Len(strCtlCaption) \ 2) - 2
While lngTextWidth > lngMaxWidth And lngX > 3
    ' Text is too wide for Ctl's width;
    ' shrink the caption from the middle,
    ' replacing the 3 middlemost characters
    ' with ellipses (...)
    strCtlCaption = Left(strCtlCaption, lngX) & "..." & _
                    Right(strCtlCaption, lngX)
    lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)
    lngX = lngX - 1
Wend

FitText = strCtlCaption

End Function

Public Function FormatFileSize(ByVal dblFileSize As Double, _
                               Optional ByVal strFormatMask As String) _
                               As String

' FormatFileSize:   Formats dblFileSize in bytes into
'                   X GB or X MB or X KB or X bytes depending
'                   on size (a la Win9x Properties tab)

Select Case dblFileSize
    Case 0 To 1023              ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575        ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823#       ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select

End Function

Public Function FormatTime(ByVal sglTime As Single) As String
                           
' FormatTime:   Formats time in seconds to time in
'               Hours and/or Minutes and/or Seconds

' Determine how to display the time
Select Case sglTime
    Case 0 To 59    ' Seconds
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599 ' Minutes Seconds
        FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " min " & _
                     Format(sglTime Mod 60, "0") & " sec"
    Case Else       ' Hours Minutes
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " hr " & _
                     Format(sglTime / 60 Mod 60, "0") & " min"
End Select

End Function

Public Function DiskFreeSpace(strDrive As String) As Double

' DiskFreeSpace:    returns the amount of free space on a drive
'                   in Windows9x/2000/NT4+

Dim SectorsPerCluster As Long
Dim BytesPerSector As Long
Dim NumberOfFreeClusters As Long
Dim TotalNumberOfClusters As Long
Dim FreeBytes As Long
Dim spaceInt As Integer

strDrive = QualifyPath(strDrive)

' Call the API function
GetDiskFreeSpace strDrive, _
                 SectorsPerCluster, _
                 BytesPerSector, _
                 NumberOFreeClusters, _
                 TotalNumberOfClusters

' Calculate the number of free bytes
DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector

End Function


Public Function QualifyPath(strPath As String) As String

' Make sure the path ends in "\"
QualifyPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")

End Function


Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String

' ReturnFileOrFolder:   Returns the filename or path of an
'                       MS-DOS file or URL.
'
' Author:   Jeff Cockayne 4.30.99
'
' Inputs:   FullPath:   String; the full path
'           ReturnFile: Boolean; return filename or path?
'                       (True=filename, False=path)
'           IsURL:      Boolean; Pass True if path is a URL.
'
' Returns:  String:     the filename or path
'

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
If intDelimiterIndex = 0 Then
    ReturnFileOrFolder = FullPath
Else
    ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))
End If

End Function
'used for nobadchat sub I use this to filter out bad words
Function sReplace(SearchLine As String, SearchFor As String, ReplaceWith As String)
    Dim vSearchLine As String, found As Integer
    
    found = InStr(SearchLine, SearchFor): vSearchLine = SearchLine
    If found <> 0 Then
        vSearchLine = ""
        If found > 1 Then vSearchLine = Left(SearchLine, found - 1)
        vSearchLine = vSearchLine + ReplaceWith
        If SearchFor = "disconnected." Then
        Form1.txtDialog.Text = Form1.txtDialog.Text & " You have been disconnected."
        End If
        If found + Len(SearchFor) - 1 < Len(SearchLine) Then _
            vSearchLine = vSearchLine + Right$(SearchLine, Len(SearchLine) - found - Len(SearchFor) + 1)
    End If
    sReplace = vSearchLine
    
End Function


'****The Following are the Encryption Functions
Function encdec(inputstrinG As String) As String

If Len(inputstrinG) = 0 Then Exit Function

Dim p As String
Dim o As String
Dim K As String
Dim S As String
Dim tempstr As String

For I = 1 To Len(inputstrinG)
p = Mid$(inputstrinG, I, 1)

o = Asc(p)
K = o Xor 2
S = Chr$(K)
tempstr = tempstr & S

Next I

encdec = tempstr
End Function


Private Sub EncryptFile(file2open As String, file2save As String)

Dim result      As Long
Dim fref1       As Long
Dim fref2       As Long
Dim tempstr     As String * 1
Dim starttime   As Long
Dim endtime     As Long

'starttime = GetTickCount
process = True

fref1 = FreeFile

Open file2open For Binary As #fref1

fref2 = FreeFile

Open file2save For Binary As #fref2

For result = 1 To LOF(fref1)

Get #fref1, , tempstr
Put #fref2, , encdec(tempstr)

Next result

Close #fref2
Close #fref1


process = False
'endtime = GetTickCount

End Sub
Private Sub DecryptFile(file2open As String, file2save As String)

Dim tempstr As String * 1
Dim result As Long
Dim fref1 As Long
Dim fref2 As Long
Dim starttime As Long
Dim endtime As Long

'starttime = GetTickCount
process = True

fref1 = FreeFile

Open file2open For Binary As #fref1

fref2 = FreeFile

Open file2save For Binary As #fref2

For result = 1 To LOF(fref1)

Get #fref1, , tempstr
Put #fref2, , encdec(tempstr)

Next result

Close #fref2
Close #fref1

'endtime = GetTickCount
process = False

End Sub

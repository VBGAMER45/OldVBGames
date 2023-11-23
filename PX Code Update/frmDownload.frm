VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmDownload.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1080
      Top             =   3360
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   3600
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3960
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1485
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   313
      FullHeight      =   41
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   2197
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblFILES 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of files remaining:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label RateLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "RateLabel"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2400
      Width           =   3765
   End
   Begin VB.Label TransferRate 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer rate:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label ToLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "ToLabel"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   2100
      Width           =   3765
   End
   Begin VB.Label DownloadTo 
      BackStyle       =   0  'Transparent
      Caption         =   "Download to:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2100
      Width           =   1575
   End
   Begin VB.Label TimeLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "TimeLabel"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   3765
   End
   Begin VB.Label SourceLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "SourceLabel"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5250
   End
   Begin VB.Label EstimatedTimeLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated time left:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label StatusLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "StatusLabel"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5235
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this form is to download updates to Project X
'I just added the unzipping functions to back downloading faster
Private CancelSearch As Boolean
Dim NumberOfFiles As Integer
Dim ZipFiles As Integer
Public Function DownloadFile(strURL As String, _
                             strDestination As String, _
                             Optional UserName As String = Empty, _
                             Optional Password As String = Empty) _
                             As Boolean

' Funtion DownloadFile: Download a file via HTTP
'
' Author:   Jeff Cockayne
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

Const CHUNK_SIZE As Long = 1024 ' Download chunk size
Const ROLLBACK As Long = 4096   ' Bytes to roll back on resume
                                ' You can be less conservative,
                                ' and roll back less, but I
                                ' don't recommend it.
Dim bData() As Byte             ' Data var
Dim blnResume As Boolean        ' True if resuming download
Dim intFile As Integer          ' FreeFile var
Dim lngBytesReceived As Long    ' Bytes received so far
Dim lngFileLength As Long       ' Total length of file in bytes
Dim lngX                        ' Temp long var
Dim sglLastTime As Single          ' Time last chunk received
Dim sglRate As Single           ' Var to hold transfer rate
Dim sglTime As Single           ' Var to hold time remaining
Dim strFile As String           ' Temp filename var
Dim strHeader As String         ' HTTP header store
Dim strHost As String           ' HTTP Host

On Local Error GoTo InternetErrorHandler

' Start with Cancel flag = False
CancelSearch = False

' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
              
SourceLabel = Empty
TimeLabel = Empty
ToLabel = Empty
RateLabel = Empty

' Pre-open the AVI
With Animation1
    .AutoPlay = True
    .Open App.Path & "\DOWNLD2.AVI"
End With

' Show the download status form
Show
' Move form into view
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

StartDownload:

If blnResume Then
    StatusLabel = "Resuming download..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    StatusLabel = "Getting file information..."
End If
' Give the system time to update the form gracefully
DoEvents

' Download file
With Inet1
    .Url = strURL
    .UserName = UserName
    .Password = Password
    ' GET file, sending the magic resume input header...
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
    ' While initiating connection, yield CPU to Windows
    While .StillExecuting
        DoEvents
        ' If user pressed Cancel button on StatusForm
        ' then fail, cancel, and exit this download
        If CancelSearch Then GoTo ExitDownload
    Wend

    StatusLabel = "Saving:"
    SourceLabel = FitText(SourceLabel, strHost & " from " & .RemoteHost)
    ToLabel = FitText(ToLabel, strDestination)

    ' Get first header ("HTTP/X.X XXX ...")
    strHeader = .GetHeader
End With

' Trap common HTTP response codes
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK
        ' If resuming, however, this is a failure
        If blnResume Then
            ' Delete partially downloaded file
            Kill strDestination
            ' Prompt
            If MsgBox("The server is unable to resume this download." & _
                      vbCr & vbCr & _
                      "Do you want to continue anyway?", _
                      vbExclamation + vbYesNo, _
                      "Unable to Resume Download") = vbYes Then
                    ' Yes - continue anyway:
                    ' Set resume flag to False
                    blnResume = False
                Else
                    ' No - cancel
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
            
    Case "206"  ' 206=Partial Content, which is GREAT when resuming!
    
    Case "204"  ' No content
        MsgBox "Nothing to download!", _
               vbInformation, _
               "No Content"
        CancelSearch = True
        GoTo ExitDownload
        
    Case "401"  ' Not authorized
        MsgBox "Authorization failed!", _
               vbCritical, _
               "Unauthorized"
        CancelSearch = True
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        MsgBox "The file, " & _
               """" & Inet1.Url & """" & _
               " was not found!", _
               vbCritical, _
               "File Not Found"
        CancelSearch = True
        GoTo ExitDownload
        
    Case vbCrLf ' Empty header
        MsgBox "Cannot establish connection." & vbCr & vbCr & _
               "Check your Internet connection and try again.", _
               vbExclamation, _
               "Cannot Establish Connection"
        CancelSearch = True
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MsgBox "The server returned the following response:" & _
               vbCr & vbCr & _
               strHeader, _
               vbCritical, _
               "Error Downloading File"
        CancelSearch = True
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
If blnResume = False Then
    ' Set timer for gauging download speed
    sglLastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If

' Check for available disk space first...
' If on a physical or mapped drive. Can't with a UNC path.
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, _
                          InStr(strDestination, "\"))) < lngFileLength Then
        ' Not enough free space to download file
        MsgBox "There is not enough free space on disk for this file." _
               & vbCr & vbCr & "Please free up some disk space and try again.", _
               vbCritical, _
               "Insufficient Disk Space"
        GoTo ExitDownload
    End If
End If

' Prepare display
'
' Progress Bar
With ProgressBar
    .Value = 0
    .Max = lngFileLength
End With

' Give system a chance to show AVI
DoEvents

' Reset bytes received counter if not resuming
If blnResume = False Then lngBytesReceived = 0


On Local Error GoTo FileErrorHandler

' Create destination directory, if necessary
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If

' If no errors occurred, then spank the file to disk
intFile = FreeFile()        ' Set intFile to an unused file.
' Open a file to write to.
Open strDestination For Binary Access Write As #intFile
' If resuming, then seek byte position in downloaded file
' where we last left off...
If blnResume Then Seek #intFile, lngBytesReceived + 1
Do
    ' Get chunks...
    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    Put #intFile, , bData   ' Put it into our destination file
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    TimeLabel = FormatTime(sglTime) & _
                   " (" & _
                   FormatFileSize(lngBytesReceived) & _
                   " of " & _
                   FormatFileSize(lngFileLength) & _
                   " copied)"
    RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
    ProgressBar.Value = lngBytesReceived
    Me.Caption = Format((lngBytesReceived / lngFileLength), "##0%") & _
                 " of " & strFile & " Completed"
Loop While UBound(bData, 1) > 0       ' Loop while there's still data...
Close #intFile

ExitDownload:
' Success if the # of bytes transferred = content length
If lngBytesReceived = lngFileLength Then
    StatusLabel = "Download completed!"
    DownloadFile = True
    If strHost = "update.txt" Then
    StatusLabel.Caption = "Checking files to download"
    Label1.Visible = True
    LblFILES.Visible = True
    Call LoadList(List1, App.Path & "\update\update.txt")
    LblFILES.Caption = List1.ListCount
    NumberOfFiles = List1.ListCount
    ZipFiles = List1.ListCount - 1
    Timer1.Enabled = True
    End If
    
Else
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        ' Resume? (If not cancelled)
        If CancelSearch = False Then
            If MsgBox("The connection with the server was reset." & _
                      vbCr & vbCr & _
                      "Click ""Retry"" to resume downloading the file." & _
                      vbCr & "(Approximate time remaining: " & FormatTime(sglTime) & ")" & _
                      vbCr & vbCr & _
                      "Click ""Cancel"" to cancel downloading the file.", _
                      vbExclamation + vbRetryCancel, _
                      "Download Incomplete") = vbRetry Then
                    ' Yes
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If
    ' No or unresumable failure:
    ' Delete partially downloaded file
    If Not Dir(strDestination) = Empty Then Kill strDestination
    DownloadFile = False
End If

CleanUp:
' Close AVI
Animation1.Close

' Make sure that the Internet connection is closed...
Inet1.Cancel
' ...and exit this function
'Unload Me
'add my code here!!!!!!!!!!!!!

'
Exit Function

InternetErrorHandler:
    ' Err# 9 occurs when UBound(bData,1) < 0
    If Err.Number = 9 Then Resume Next
    ' Other errors...
    MsgBox "Error: " & Err.Description & " occurred.", _
           vbCritical, _
           "Error Downloading File"
    Err.Clear
    GoTo ExitDownload
    
FileErrorHandler:
    MsgBox "Cannot write file to disk." & _
           vbCr & vbCr & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Error Downloading File"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
    
End Function


Private Sub CancelButton_Click()
    StatusLabel = "Cancelling..."
    Animation1.Close
    CancelSearch = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' For some reason, the Cancel=True and Default=True
' Properties of the "Cancel" button are not working...
' This is the first time that's ever happened to me.
' Any ideas?
If KeyCode = vbKeyEscape Then CancelButton_Click

End Sub

Private Sub Form_Load()

' Move form off-screen until it is ready (completely drawn)
Move -Width, -Height

End Sub


Private Sub Form_Unload(Cancel As Integer)

' Move form off-screen so that it disappears "instantly"
Move -Width, -Height
DoEvents

End Sub


Sub LoadList(Lst As ListBox, file As String)
'Call LoadList (List1,"C:\Windows\System\Saved.lst")
On Error GoTo error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
Lst.AddItem A$
Loop
Close 1
Exit Sub
error:
Exit Sub
'x = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

Private Sub Timer1_Timer()
'This checks how many files remaining then downloads them
If ProgressBar.Value = ProgressBar.Max Then
If NumberOfFiles > -1 Then
DownloadFile "http://" & MainPlayer.ServerIP & "/update/" & List1.List(NumberOfFiles), _
                         App.Path & "\update\" & List1.List(NumberOfFiles)

NumberOfFiles = NumberOfFiles - 1
LblFILES.Caption = LblFILES.Caption - 1
End If
End If
If NumberOfFiles = -1 Then
If ZipFiles > -1 Then
StatusLabel.Caption = "Files Downloaded! Updating Files!"
If List1.List(ZipFiles) = "exe.zip" Then
If List1.List(ZipFiles) = "" Then Exit Sub
UnZipModule.Unzip App.Path & "\update\" & List1.List(ZipFiles), App.Path & "\update\"
ZipFiles = ZipFiles - 1
  F = FreeFile
Open (App.Path & "\update\check.def") For Binary Access Write Lock Write As F
Put #F, , "Update" & vbCrLf
Put #F, , App.Path & "\" & App.EXEName & ".exe" & vbCrLf
Close F

Call Shell(App.Path & "\update\exepatcher.exe")
End
End If
If List1.List(ZipFiles) = "Monsters.zip" Then
If List1.List(ZipFiles) = "" Then Exit Sub
UnZipModule.Unzip App.Path & "\update\" & List1.List(ZipFiles), App.Path & "\update\monsters\"
ZipFiles = ZipFiles - 1
End If
If List1.List(ZipFiles) = "Maps.zip" Then
If List1.List(ZipFiles) = "" Then Exit Sub
UnZipModule.Unzip App.Path & "\update\" & List1.List(ZipFiles), App.Path & "\update\maps\"
ZipFiles = ZipFiles - 1
Else
If List1.List(ZipFiles) = "" Then Exit Sub
UnZipModule.Unzip App.Path & "\update\" & List1.List(ZipFiles), App.Path & "\update\"

ZipFiles = ZipFiles - 1
End If

Else 'when its done unzipping show the login screen
frmLogin.Show
Unload Me


End If
End If

End Sub
Function FileExists%(fname$)
 On Local Error Resume Next
 
 Dim ff%
 ff% = FreeFile
 Open fname$ For Input As ff%
 If Err Then
  FileExists% = False
 Else
  FileExists% = True
 End If
 
 Close ff%
End Function


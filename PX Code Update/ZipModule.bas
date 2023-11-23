Attribute VB_Name = "ZipModule"
Option Explicit

Public Enum ZTranslate
    CRLFtoLF = 1
    LFtoCRLF = 2
End Enum
'
' Collection of Files to Zip
'
Private mCollection As Collection
'
' Recurse Folders ?
'
Private miRecurseFolders As Integer
'
' Zip File Name
'
Private msZipFileName As String
'
' Encryption ?
'
Private miEncrypt As Integer
'
' System Files
'
Private miSystem As Integer
'
' Root Directory
'
Private msRootDirectory As String
'
' Verbose Zip
'
Private miVerbose As Integer
'
' Quiet Zip
'
Private miQuiet As Integer
'
' Translate CRLF / LF Chars
'
Private miTranslateCRLF As ZTranslate
'
' Updating Existing Zip ?
'
Private miUpdateZip As Integer

Public Function MakeZipFile(sFileName As String, sZFileName As String) As Long
    Dim zFileArray As ZIPnames
    Dim iIgnorePath As Integer
    Dim iRecurse As Integer

On Error GoTo vbErrorHandler
    
        zFileArray.s(0) = "bosdosya"
        zFileArray.s(1) = sFileName
    
    msZipFileName = sZFileName
    miEncrypt = 0
    miSystem = 0
    msRootDirectory = "\"
    miQuiet = 0
    miUpdateZip = 0
    miRecurseFolders = 0
    
    MakeZipFile = VBZip(2, msZipFileName, _
        zFileArray, iIgnorePath, _
        miRecurseFolders, miUpdateZip, _
        0, msRootDirectory)
   

    Exit Function

vbErrorHandler:

End Function




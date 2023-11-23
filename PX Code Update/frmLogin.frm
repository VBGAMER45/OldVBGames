VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2340
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3765
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1382.549
   ScaleMode       =   0  'User
   ScaleWidth      =   3535.131
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdDone 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   1140
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "New User"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   270
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   390
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Confirm:"
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


Private Sub Check1_Click()
If Check1.Value = 0 Then
    lblLabels(2).Enabled = False
    Text1.Enabled = False
End If
If Check1.Value = 1 Then
    lblLabels(2).Enabled = True
    Text1.Enabled = True
End If
End Sub

Private Sub CmdCancel_Click()

    Me.Hide
    Unload Me
    
    End
    
End Sub

Private Sub CmdDone_Click()
Dim Charascii As Integer, Loopc As Integer
If Check1.Value = 1 Then

If Text1.Text = "" Then
MsgBox "Your password has to be at least one letter or number!", vbExclamation
Exit Sub
End If
   
If txtUserName.Text = "" Then
MsgBox "Your username has to be at least one letter or number!", vbExclamation
Exit Sub
End If
 
If txtUserName.Text = "" Or txtPassword.Text = "" Then
MsgBox "Please enter a username or password!", vbExclamation
Exit Sub
End If
If txtPassword.Text = Text1.Text Then

    For Loopc = 1 To Len(txtPassword.Text)

    Charascii = Asc(Mid$(txtPassword.Text, Loopc, 1))
    If Server.LegalCharacter(Charascii) = False Then
        MsgBox ("Invalid Password because of use of specail characters")
       Exit Sub
    
    End If
    
Next Loopc
  For Loopc = 1 To Len(txtUserName.Text)

    Charascii = Asc(Mid$(txtUserName.Text, Loopc, 1))
    If Server.LegalCharacter(Charascii) = False Then
        MsgBox ("Invalid Username because of use of specail characters")
       Exit Sub
    
    End If
    
Next Loopc
Else
MsgBox "Passwords do not match", vbExclamation
Exit Sub
End If
'On Error GoTo nofile:
If Winsock1.State = sckConnected Then
Call SendLogin
Else
Winsock1.Connect
End If
Else
If txtUserName.Text = "" Then
MsgBox "Enter a username!", vbExclamation
Exit Sub
End If
If txtPassword.Text = "" Then
MsgBox "Enter a password!", vbExclamation
Exit Sub
End If

'On Error GoTo nofile:
If Winsock1.State = sckConnected Then
Call SendLogin
Else
Winsock1.Connect
End If
'nofile:
'Call SendLogin

'Exit Sub
End If
End Sub



Private Sub Form_Load()
Winsock1.RemotePort = "5774"
Winsock1.RemoteHost = MainPlayer.ServerIP

End Sub


Private Sub SendLogin()
Dim logintext As String
If Me.Visible = True Then
If Check1.Value = 0 Then
logintext = "login" & "," & frmLogin.txtUserName.Text & "," & frmLogin.txtPassword.Text
Winsock1.SendData logintext
Else
logintext = "New" & "," & frmLogin.txtUserName.Text & "," & frmLogin.txtPassword.Text
Winsock1.SendData logintext
End If
End If
End Sub

Private Sub Winsock1_Connect()
Dim logintext As String
If Check1.Value = 0 Then
logintext = "login" & "," & frmLogin.txtUserName.Text & "," & frmLogin.txtPassword.Text
Winsock1.SendData logintext
Else
logintext = "New" & "," & frmLogin.txtUserName.Text & "," & frmLogin.txtPassword.Text
Winsock1.SendData logintext
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim incomdata As String
Winsock1.GetData incomdata
DataProcess incomdata
End Sub
Sub DataProcess(DataStr As String)
'Online login code

If DataStr = "Login Failed" Then
MsgBox "Someone Else is using this account!", vbExclamation
End If
If Left(DataStr, 4) = "list" Then
'Loads Character List.
Call loadcharlist(DataStr)

End If
If Left(DataStr, 6) = "BadLog" Then
'If DataStr = "BadLog" Then
MsgBox "Invaild Password or UserName", vbExclamation
End If
If DataStr = "None" Then
MsgBox "The UserDoes not exist", vbExclamation
End If
If Left(DataStr, 7) = "LoginOk" Then
'If DataStr = "LoginOk" Then
Mutiplayer.currrentusers.AddItem txtUserName.Text
OnlineGame = True
MainPlayer.AccountName = txtUserName.Text
MainPlayer.Password = txtPassword.Text
'Winsock1.Close
Winsock1.SendData "chars" & "," & MainPlayer.AccountName
Mutiplayer.Show
'Me.Hide
Me.Visible = False
End If
'If DataStr = "AccountOk" Then
If Left(DataStr, 9) = "AccountOk" Then
MainPlayer.AccountName = txtUserName.Text
MainPlayer.Password = txtPassword.Text
Mutiplayer.currrentusers.AddItem txtUserName.Text
OnlineGame = True
'Winsock1.Close
Mutiplayer.Show
Me.Visible = False
End If
If DataStr = "Account Failed" Then
MsgBox "The account already exists or invaild charecters used for account name!", vbExclamation
End If

End Sub
Private Sub loadcharlist(Data As String)
Dim chararray As Variant
Dim X As Integer
Dim Account As String
chararray = Tokenize(Data)
X = chararray(1)
Account = chararray(2)
If MainPlayer.AccountName = Account Then
I = 2
Mutiplayer.List1.Clear
On Error GoTo nofile
For I = I + 1 To X + 3
Mutiplayer.List1.AddItem chararray(I)
Next
nofile:
Exit Sub
End If
End Sub

VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form12 
   Caption         =   "Under Ground"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":030A
   ScaleHeight     =   5895
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3100
      Left            =   2760
      Top             =   5520
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5295
      Left            =   7200
      Max             =   5000
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   5000
      TabIndex        =   10
      Top             =   5280
      Width           =   7455
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   8160
      Picture         =   "Form12.frx":154A5C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   7560
      Picture         =   "Form12.frx":154D66
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close windows"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture5 
      Height          =   1095
      Left            =   7440
      Picture         =   "Form12.frx":155070
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   6
      ToolTipText     =   "farm"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   0
         Picture         =   "Form12.frx":156AB2
         ScaleHeight     =   1035
         ScaleWidth      =   1275
         TabIndex        =   22
         ToolTipText     =   "communications"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         Begin ComctlLib.ProgressBar ProgressBar6 
            Height          =   135
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   238
            _Version        =   327682
            Appearance      =   1
         End
      End
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   135
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   7440
      Picture         =   "Form12.frx":157FD4
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   5
      ToolTipText     =   "elevator"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   150
         Left            =   0
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   265
         _Version        =   327682
         Appearance      =   1
         Min             =   1
         Max             =   10
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   7440
      Picture         =   "Form12.frx":1594B6
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   4
      ToolTipText     =   "tube right"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
      Begin ComctlLib.ProgressBar ProgressBar3 
         Height          =   135
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   7440
      Picture         =   "Form12.frx":15DD28
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   3
      ToolTipText     =   "tube left"
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
      Begin ComctlLib.ProgressBar ProgressBar4 
         Height          =   135
         Left            =   0
         TabIndex        =   15
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7440
      Picture         =   "Form12.frx":15ECAA
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   2
      ToolTipText     =   "Crosstube"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      Begin ComctlLib.ProgressBar ProgressBar5 
         Height          =   135
         Left            =   0
         TabIndex        =   16
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Build Windows"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image Image43 
      Height          =   615
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image42 
      Height          =   615
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image41 
      Height          =   615
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image40 
      Height          =   615
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image39 
      Height          =   855
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image38 
      Height          =   855
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image37 
      Height          =   855
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image36 
      Height          =   855
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image Image35 
      Height          =   975
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image34 
      Height          =   975
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image33 
      Height          =   975
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image32 
      Height          =   975
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image31 
      Height          =   615
      Left            =   3960
      OLEDropMode     =   2  'Automatic
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image30 
      Height          =   855
      Left            =   3960
      OLEDropMode     =   2  'Automatic
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image29 
      Height          =   975
      Left            =   3960
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image Image28 
      Height          =   615
      Left            =   4920
      OLEDropMode     =   2  'Automatic
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image27 
      Height          =   855
      Left            =   4920
      OLEDropMode     =   2  'Automatic
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image Image26 
      Height          =   975
      Left            =   4920
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image Image25 
      Height          =   975
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image24 
      Height          =   975
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image23 
      Height          =   975
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image22 
      Height          =   975
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image21 
      Height          =   975
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image20 
      Height          =   975
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image19 
      Height          =   975
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image18 
      Height          =   975
      Left            =   960
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image17 
      Height          =   975
      Left            =   1920
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image16 
      Height          =   615
      Left            =   6000
      OLEDropMode     =   2  'Automatic
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image15 
      Height          =   855
      Left            =   6000
      OLEDropMode     =   2  'Automatic
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Image14 
      Height          =   975
      Left            =   6000
      OLEDropMode     =   2  'Automatic
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image Image13 
      Height          =   975
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image12 
      Height          =   975
      Left            =   3960
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   975
   End
   Begin VB.Image Image11 
      Height          =   975
      Left            =   4920
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image10 
      Height          =   975
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image9 
      Height          =   975
      Left            =   3960
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image8 
      Height          =   975
      Left            =   2880
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Image Image7 
      Height          =   975
      Left            =   3960
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image6 
      Height          =   975
      Left            =   6000
      OLEDropMode     =   2  'Automatic
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   4920
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   6000
      OLEDropMode     =   2  'Automatic
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   4920
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   6000
      OLEDropMode     =   2  'Automatic
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   8340
      Left            =   -3240
      Picture         =   "Form12.frx":15FD64
      Top             =   0
      Width           =   12525
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Command4.Visible = True
Command5.Visible = True

End Sub

Private Sub Command3_Click()
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Command4.Visible = False
Command5.Visible = False
End Sub

Private Sub Command4_Click()
Picture6.Visible = False
End Sub

Private Sub Command5_Click()
Picture6.Visible = True
End Sub

Private Sub Form_Load()
'ProgressBar1.Align = vbAlignBottom
    ProgressBar1.Visible = False
    ProgressBar2.Visible = False
    ProgressBar3.Visible = False
    ProgressBar4.Visible = False
    ProgressBar5.Visible = False
    'Command1.Caption = "Initialize array"
End Sub

Private Sub HScroll1_Change()
  Image4.Left = -HScroll1.Value
End Sub

Private Sub Label1_Click()
If Label1.Caption = 4 Then
  Label1.Caption = 5
  End If
  If Label1.Caption = 3 Then
 Label1.Caption = 4
 End If
 If Label1.Caption = 2 Then
 Label1.Caption = 3
 End If
  If Label1.Caption = 1 Then
 Label1.Caption = 2
 End If
 If Label1.Caption = 0 Then
 Label1.Caption = 1
 End If
End Sub

Private Sub Label2_Click()
If Label2.Caption = 4 Then
  Label2.Caption = 5
  End If
  If Label2.Caption = 3 Then
 Label2.Caption = 4
 End If
 If Label2.Caption = 2 Then
 Label2.Caption = 3
 End If
  If Label2.Caption = 1 Then
 Label2.Caption = 2
 End If
 If Label2.Caption = 0 Then
 Label2.Caption = 1
 End If
End Sub

Private Sub Label3_Click()
If Label3.Caption = 4 Then
  Label3.Caption = 5
  End If
  If Label3.Caption = 3 Then
 Label3.Caption = 4
 End If
 If Label3.Caption = 2 Then
 Label3.Caption = 3
 End If
  If Label3.Caption = 1 Then
 Label3.Caption = 2
 End If
 If Label3.Caption = 0 Then
 Label3.Caption = 1
 End If
End Sub

Private Sub Label4_Click()
If Label4.Caption = 4 Then
  Label4.Caption = 5
  End If
  If Label4.Caption = 3 Then
 Label4.Caption = 4
 End If
 If Label4.Caption = 2 Then
 Label4.Caption = 3
 End If
  If Label4.Caption = 1 Then
 Label4.Caption = 2
 End If
 If Label4.Caption = 0 Then
 Label4.Caption = 1
 End If
End Sub

Private Sub Label5_Click()
If Label5.Caption = 4 Then
  Label5.Caption = 5
  End If
  If Label5.Caption = 3 Then
 Label5.Caption = 4
 End If
 If Label5.Caption = 2 Then
 Label5.Caption = 3
 End If
  If Label5.Caption = 1 Then
 Label5.Caption = 2
 End If
 If Label5.Caption = 0 Then
 Label5.Caption = 1
 End If
End Sub

Private Sub Picture1_Click()

  'If Label1.Caption = 5 Then
  'Label1.Caption = 0
  'End If
  'If Label1.Caption = 4 Then
 ' Label1.Caption = 5
  'End If
  'If Label1.Caption = 3 Then
 'Label1.Caption = 4
 'End If
 'If Label1.Caption = 2 Then
 'Label1.Caption = 3
 'End If
 ' If Label1.Caption = 1 Then
 'Label1.Caption = 2
 'End If
  ' If Label1.Caption = 0 Then
' Label1.Caption = 1
' End If

' Label1 = Label1 + 1

 Dim Counter As Integer
    Dim Workarea(30000) As String
    ProgressBar5.min = LBound(Workarea)
    ProgressBar5.Max = UBound(Workarea)
    ProgressBar5.Visible = True

'Set the Progress's Value to Min.
    ProgressBar5.Value = ProgressBar5.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar5.Value = Counter
    Next Counter
    ProgressBar5.Visible = False
    ProgressBar5.Value = ProgressBar5.min
    If Label1 >= 1 Then
    Label1 = Label1 - 1
    End If
    form2.money.Caption = form2.money.Caption - 350
Picture1.OLEDragMode = 1
Timer2.Enabled = True
End Sub

Private Sub Picture2_Click()
 Dim Counter As Integer
    Dim Workarea(20000) As String
    ProgressBar4.min = LBound(Workarea)
    ProgressBar4.Max = UBound(Workarea)
    ProgressBar4.Visible = True

'Set the Progress's Value to Min.
    ProgressBar4.Value = ProgressBar4.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar4.Value = Counter
    Next Counter
    ProgressBar4.Visible = False
    ProgressBar4.Value = ProgressBar4.min
   ' Label2 = Label2 - 1
   form2.money.Caption = form2.money.Caption - 200
   Picture2.OLEDragMode = 1
   Timer2.Enabled = True
     If Label2 >= 1 Then
    Label2 = Label2 - 1
    End If
End Sub

Private Sub Picture3_Click()
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar3.min = LBound(Workarea)
    ProgressBar3.Max = UBound(Workarea)
    ProgressBar3.Visible = True

'Set the Progress's Value to Min.
    ProgressBar3.Value = ProgressBar3.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar3.Value = Counter
    Next Counter
    ProgressBar3.Visible = False
    ProgressBar3.Value = ProgressBar3.min
    form2.money.Caption = form2.money.Caption - 200
    Picture3.OLEDragMode = 1
    Timer2.Enabled = True
     If Label3 >= 1 Then
    Label3 = Label3 - 1
    End If
End Sub

Private Sub Picture4_Click()
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar1.min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.min
    form2.money.Caption = form2.money.Caption - 400
    Picture4.OLEDragMode = 1
    Timer2.Enabled = True
  If Label4 >= 1 Then
    Label4 = Label4 - 1
    End If
'ProgressBar1.Visible = True
'ProgressBar1.Value = ProgressBar1.Min
End Sub



Private Sub Picture5_Click()
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar2.min = LBound(Workarea)
    ProgressBar2.Max = UBound(Workarea)
    ProgressBar2.Visible = True

'Set the Progress's Value to Min.
    ProgressBar2.Value = ProgressBar2.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar2.Value = Counter
    Next Counter
    ProgressBar2.Visible = False
    ProgressBar2.Value = ProgressBar2.min
    form2.money.Caption = form2.money.Caption - 1000
    Picture5.OLEDragMode = 1
    Timer2.Enabled = True
      If Label5 >= 1 Then
    Label5 = Label5 - 1
    End If
End Sub

Private Sub Picture6_Click()
 Dim Counter As Integer
    Dim Workarea(10000) As String
    ProgressBar6.min = LBound(Workarea)
    ProgressBar6.Max = UBound(Workarea)
    ProgressBar6.Visible = True

'Set the Progress's Value to Min.
    ProgressBar6.Value = ProgressBar6.min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar6.Value = Counter
    Next Counter
    ProgressBar6.Visible = False
    ProgressBar6.Value = ProgressBar6.min
    form2.money.Caption = form2.money.Caption - 600
    Picture6.OLEDragMode = 1
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Picture1.OLEDragMode = 0
Picture2.OLEDragMode = 0
Picture3.OLEDragMode = 0
Picture4.OLEDragMode = 0
Picture5.OLEDragMode = 0
Picture6.OLEDragMode = 0
Picture1.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cool\z1.bmp")
Picture2.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cool\z2.bmp")
Picture3.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cool\z3.bmp")
Picture4.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cool\z4.bmp")
Picture5.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cool\z8.bmp")
Picture6.Picture = LoadPicture("c:\windows\desktop\vb projects\war\cool\aa3.bmp")
Timer2.Enabled = False
End Sub

Private Sub VScroll1_Change()
    Image4.Top = -VScroll1.Value

End Sub

Private Sub Timer1_Timer()
If Picture1.Visible = True Then
Command3.Visible = True
End If
If Picture1.Visible = False Then
Command3.Visible = False
End If

End Sub
   'ProgressBar1.Align = vbAlignBottom
   ' ProgressBar1.Visible = False
   ' Command1.Caption = "Initialize array"
'End Sub

'Private Sub Command1_Click()
    'ProgressBar1.Align = vbAlignBottom
    'ProgressBar1.Visible = False
   ' Command1.Caption = "Initialize array"
'End Sub

'Private Sub Command1_Click()
   ' Dim Counter As Integer
    'Dim Workarea(10000) As String
   ' ProgressBar1.Min = LBound(Workarea)
   ' ProgressBar1.Max = UBound(Workarea)
   ' ProgressBar1.Visible = True

'Set the Progress's Value to Min.
   ' ProgressBar1.Value = ProgressBar1.Min

'Loop through the array.
   ' For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        'Workarea(Counter) = "Initial value" & Counter
       ' ProgressBar1.Value = Counter
   ' Next Counter
    'ProgressBar1.Visible = False
    'ProgressBar1.Value = ProgressBar1.Minn

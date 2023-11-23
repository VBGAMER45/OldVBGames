VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStoreEdit 
   Caption         =   "Store Editor 1.0 Jonathan Valentin 2003"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdLoadStore 
      Caption         =   "Load Store"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton CmdSellRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CmdBuyRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdSellAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton CmdBuyAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Store Name"
      Top             =   480
      Width           =   4695
   End
   Begin VB.ListBox LstSell 
      Height          =   1425
      Left            =   3600
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ListBox lstBuy 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblStore 
      Caption         =   "Store Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblSell 
      Caption         =   "Items the Npc Sells"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblBuy 
      Caption         =   "Items the Npc Buys"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "frmStoreEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Store Editor for Projectx
'Jonathan Valentin 2003
'http://www.visualbasiczone.com


Private Type StoreType
    StoreName As String
    Buys() As String
    Sells() As String
End Type
Dim Store As StoreType

Private Sub CmdBuyAdd_Click()
CommonDialog1.Filter = "PX Items (*.item)|*.item"


CommonDialog1.Flags = cdlOFNExplorer

CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then
Else
lstBuy.AddItem CommonDialog1.FileTitle

End If

End Sub

Private Sub CmdBuyRemove_Click()
If lstBuy.Text = "" Then
Else
lstBuy.RemoveItem lstBuy.ListIndex
End If

End Sub

Private Sub CmdDone_Click()
If txtName.Text = "" Then
    MsgBox "Please enter a name for your store"
Else
    ReDim Preserve Store.Buys(lstBuy.ListCount)
    ReDim Preserve Store.Sells(LstSell.ListCount)
    
'Load the data from the listboxes into the array
For i = 0 To lstBuy.ListCount
    Store.Buys(i) = lstBuy.List(i)
Next i

For i = 0 To LstSell.ListCount
    Store.Sells(i) = LstSell.List(i)
Next i

Store.StoreName = txtName.Text

'now write the file
Open App.Path & "\" & txtName.Text & ".shop" For Binary Access Write Lock Write As #1
    Put #1, , Store
Close #1

MsgBox "Shop has been created!", vbExclamation
End If

End Sub

Private Sub CmdLoadStore_Click()
CommonDialog1.Filter = "PX Shop (*.shop)|*.shop"


CommonDialog1.Flags = cdlOFNExplorer

CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then
Else
LstSell.Clear
lstBuy.Clear

Open CommonDialog1.FileName For Binary Access Read Lock Read As #1
    Get #1, , Store
Close #1

txtName.Text = Store.StoreName

For i = 0 To UBound(Store.Buys)
    If Store.Buys(i) = "" Then
    Else
    lstBuy.AddItem Store.Buys(i)
    End If
    
Next i

For i = 0 To UBound(Store.Sells)
    If Store.Sells(i) = "" Then
    Else
    LstSell.AddItem Store.Sells(i)
    End If
Next i

End If
End Sub

Private Sub CmdSellAdd_Click()
CommonDialog1.Filter = "PX Items (*.item)|*.item"


CommonDialog1.Flags = cdlOFNExplorer

CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then
Else

LstSell.AddItem CommonDialog1.FileTitle

End If
End Sub

Private Sub CmdSellRemove_Click()
If LstSell.Text = "" Then
Else
LstSell.RemoveItem LstSell.ListIndex
End If

End Sub


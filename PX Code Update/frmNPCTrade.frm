VERSION 5.00
Begin VB.Form frmNPCTrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   NPC Trade"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSellQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5535
      TabIndex        =   41
      Text            =   "1"
      Top             =   2430
      Width           =   1140
   End
   Begin VB.TextBox txtBuyQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2025
      TabIndex        =   38
      Text            =   "1"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.VScrollBar vsPlayer 
      Height          =   1755
      LargeChange     =   5
      Left            =   6540
      SmallChange     =   5
      TabIndex        =   36
      Top             =   540
      Width           =   315
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   14
      Left            =   5940
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   13
      Left            =   5340
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   12
      Left            =   4740
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   11
      Left            =   4140
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   10
      Left            =   3540
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   9
      Left            =   5940
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   8
      Left            =   5340
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   7
      Left            =   4740
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   4140
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   5
      Left            =   3540
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   4
      Left            =   5940
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   5340
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   2
      Left            =   4740
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   4140
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picPlayerItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   3540
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "Sell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4500
      TabIndex        =   20
      Top             =   3315
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   900
      TabIndex        =   19
      Top             =   3315
      Width           =   1215
   End
   Begin VB.VScrollBar vsNPC 
      Height          =   1755
      LargeChange     =   5
      Left            =   3060
      SmallChange     =   5
      TabIndex        =   18
      Top             =   540
      Width           =   315
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   14
      Left            =   2460
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   13
      Left            =   1860
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   12
      Left            =   1260
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   11
      Left            =   660
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   10
      Left            =   60
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1740
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   9
      Left            =   2460
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   8
      Left            =   1860
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   7
      Left            =   1260
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   660
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   5
      Left            =   60
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1140
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   4
      Left            =   2460
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   1860
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   2
      Left            =   1260
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   660
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.PictureBox picNPCItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   60
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   540
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2700
      TabIndex        =   2
      Top             =   3315
      Width           =   1215
   End
   Begin VB.Label lblMoney 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   43
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Current Money:"
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sell Quantity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3780
      TabIndex        =   40
      Top             =   2475
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Buy Quantity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   510
      TabIndex        =   39
      Top             =   2445
      Width           =   1500
   End
   Begin VB.Label lblCurrentItem 
      Alignment       =   2  'Center
      Caption         =   "lblCurrentItem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      TabIndex        =   37
      Top             =   2895
      Width           =   4275
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      Caption         =   "lblPlayerName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblNPCName 
      Alignment       =   2  'Center
      Caption         =   "lblNPCName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmNPCTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jonathan Valentin Feb 11, 2003
'For loading of custom stores
Private Type StoreType
    StoreName As String
    Buys() As String
    Sells() As String
End Type

Private Type StoreItemType
    Value As Integer
    Name As String
    Number As Integer 'location in the pxitem array so we don't have to look for it again
    Buy As Boolean
    ScrollPos As Integer
    Used As Boolean
End Type

Private Type PlayerItemType
    Value As Integer
    Name As String
    Number As Integer 'location in the pxitem array so we don't have to look for it again
    ScrollPos As Integer
End Type

Dim PlayerItem() As PlayerItemType

Dim StoreItem() As StoreItemType

Dim Store As StoreType

'Which item the user has selected to sell
Dim SelectedItemSell As Integer
'Hold the tag ot the item
Dim ItemSell As String
'Which item the user has selected to buy
Dim SelectedItemBuy As Integer
'Hold the tag ot the item
Dim ItemBuy As String

Private Sub cmdBuy_Click()
If lblCurrentItem.Caption = "" Or SelectedItemBuy = -1 Then
    MsgBox "Please select an item"
    Exit Sub
End If
If txtBuyQty.Text = "" Then
    MsgBox "Invaild Buy Qunatiy", vbInformation
    Exit Sub
End If

If txtBuyQty.Text <= 0 Then
    MsgBox "Invaild Buy Qunatiy", vbInformation
Else
    For I = 0 To UBound(StoreItem)
        If StoreItem(I).Name = lblCurrentItem.Caption Then
          
            If MainPlayer.Money >= (StoreItem(I).Value * txtBuyQty.Text) Then
                
                'remove money from player
                MainPlayer.Money = MainPlayer.Money - (StoreItem(I).Value * txtBuyQty.Text)
                lblMoney.Caption = MainPlayer.Money
                'add item to their inventory
                PXItem(StoreItem(I).Number).Quanity = PXItem(StoreItem(I).Number).Quanity + txtBuyQty.Text
                'reload the players inventory
                Call sLoadPlayerItems
            Else
                MsgBox "Sorry you do not have enought money to buy this item", , Store.StoreName
                
            End If
        
        End If
    Next I
End If
End Sub

Private Sub cmdClose_Click()
Call StoreExit
End Sub

Private Sub cmdSell_Click()
Dim PlayeritemNum As Integer
If lblCurrentItem.Caption = "" Or SelectedItemSell = -1 Then
    MsgBox "Please select an item"
    Exit Sub
End If

For I = 0 To UBound(PlayerItem)
    If PlayerItem(I).Name = picPlayerItem(SelectedItemSell).ToolTipText Then
    PlayeritemNum = I
    Exit For
    End If
Next

If txtSellQty.Text <= 0 Then
    MsgBox "Invaild Sell Qunatiy", vbInformation
Else
    If picPlayerItem(SelectedItemSell).Tag < txtSellQty.Text Then
        MsgBox "Invaild Sell Qunatiy", vbInformation
    Else
        'Now Check if the shop accepts that item
        For I = 0 To UBound(Store.Buys)
                ' MsgBox picPlayerItem(SelectedItemSell).ToolTipText
        If Store.Buys(I) = picPlayerItem(SelectedItemSell).ToolTipText & ".item" Then
       
            If MainPlayer.Money >= PlayerItem(PlayeritemNum).Value Then
                'add money to player
                MainPlayer.Money = MainPlayer.Money + (PlayerItem(PlayeritemNum).Value * txtSellQty.Text)
                lblMoney.Caption = MainPlayer.Money
                'remove item from player
                PXItem(PlayerItem(PlayeritemNum).Number).Quanity = PXItem(PlayerItem(PlayeritemNum).Number).Quanity - txtSellQty.Text
                picPlayerItem(SelectedItemSell).Tag = picPlayerItem(SelectedItemSell).Tag - txtSellQty.Text
                'reload player items
                Call LoadPlayerItems
                MsgBox "Sold: " & picPlayerItem(SelectedItemSell).ToolTipText & " of Qty# " & txtSellQty.Text & " for $" & PlayerItem(PlayeritemNum).Value * txtSellQty.Text
                'Redraw the Qty
                Call DrawQty
                
            Exit Sub
            End If
        
        End If
        Next I
        MsgBox "This Shop does not accept that item to be sold!", vbExclamation
        
    End If
    
End If

End Sub



Private Sub Form_Load()
ShopWindowOpen = True

lblCurrentItem.Caption = ""
'Loads store
lblPlayerName.Caption = MainPlayer.PlayerName
'Force textboxes numeric
RPG.ForceTextBoxNumeric txtBuyQty, True
RPG.ForceTextBoxNumeric txtSellQty, True

lblMoney.Caption = MainPlayer.Money
'set the music to the cool shop music
Form1.MediaPlayer1.Filename = App.Path & "\Inside_Shop.mid"

'Load the players items into the picture boxes
Call sLoadPlayerItems

'Call LoadStore(App.Path & "\scripts\test.shop")
End Sub

Sub LoadStore(Filename As String)


Erase StoreItem

If Filename = "" Then
Else

Open Filename For Binary Access Read Lock Read As #1
    Get #1, , Store
Close #1

lblNPCName.Caption = Store.StoreName
Me.Caption = Store.StoreName

'Make the array large enough to store the store items
ReDim StoreItem(UBound(Store.Buys) + UBound(Store.Sells))


'For I = 0 To UBound(Store.Buys)
  '  If Store.Buys(I) = "" Then
 '   Else

'    End If
'Next I

For I = 0 To UBound(Store.Sells)
    If Store.Sells(I) = "" Then
    Else
    'add item to picture box
        'Get the name of the item
    Open App.Path & "\monsters\" & Store.Sells(I) For Input As #1
    Line Input #1, textline1 'name
    Line Input #1, textline2 'damage
    Line Input #1, textline3 'speed
    Line Input #1, textline4 'monstermove
    Line Input #1, textline5 'weapon
    Line Input #1, textline6 'Item Type
    Close #1
   
   
        For G = 0 To MaxItems
            'Search for the items value
        '    MsgBox PXItem(G).ItemName
            If PXItem(G).ItemName = textline1 Then
           
            StoreItem(I).Buy = False
            StoreItem(I).Name = textline1
            StoreItem(I).Value = PXItem(G).Value
            StoreItem(I).Number = G
            G = MaxItems
            End If
            
        Next G
    'add item to picture box
    'load first 15 items into the boxes
    For pos = 0 To UBound(StoreItem)
    If StoreItem(I).Name = "" Then
    Else
        If pos <= 14 Then
        
        If picNPCItem(pos).ToolTipText = "" Then
        StoreItem(I).ScrollPos = pos
        picNPCItem(pos).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(I).Name & ".gif")
        picNPCItem(pos).ToolTipText = StoreItem(I).Name
        Exit For
        
        End If
        Else
        StoreItem(I).ScrollPos = pos
        End If
    End If
    Next pos
    
    
    End If
Next I

'set the scrollbars max values
'vsNPC.Max = UBound(Store.Sells) \ 5
vsNPC.Max = UBound(Store.Sells)
End If

End Sub
Private Sub LoadPlayerItem1()
'Loads the players items into the picture boxes

'Draw the Quantiy of each item into the picture box

'set the scrollbars max values
vsPlayer.Max = UBound(PXItem) \ 5
End Sub

Private Sub picNPCItem_Click(Index As Integer)
For I = 0 To 14
'clear all the other circles
    picNPCItem(I).Cls
Next

ItemBuy = picNPCItem(Index).Tag

'Draws a circle on selected item
picNPCItem(Index).Circle (picNPCItem(Index).Width \ 2 - 30, picNPCItem(Index).Height \ 2 - 30), picNPCItem(Index).Width \ 2 - 50, vbGreen
SelectedItemSell = -1
SelectedItemBuy = Index

lblCurrentItem.Caption = picNPCItem(Index).ToolTipText
End Sub

Private Sub picPlayerItem_Click(Index As Integer)
SelectedItemSell = Index
For I = 0 To 14
'clear all the other circles
    picPlayerItem(I).Cls
Next
Call DrawQty
ItemSell = picPlayerItem(Index).Tag
'Draws a circle on selected item
picPlayerItem(Index).Circle (picPlayerItem(Index).Width \ 2 - 30, picPlayerItem(Index).Height \ 2 - 30), picPlayerItem(Index).Width \ 2 - 50, vbGreen
SelectedItemBuy = -1
SelectedItemSell = Index

lblCurrentItem.Caption = picPlayerItem(Index).ToolTipText
End Sub
Private Sub StoreExit()
If Form1.Direction = "up" Then
    Form1.Image1.Top = Form1.Image1.Top + 32
    Form1.lblTileKind = "1"
    Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
If Form1.Direction = "down" Then
    Form1.Image1.Top = Form1.Image1.Top + 32
    Form1.lblTileKind = "1"
    Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
If Form1.Direction = "left" Then
    Form1.Image1.Top = Form1.Image1.Top + 32
    Form1.lblTileKind = "1"
    Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
If Form1.Direction = "right" Then
    Form1.Image1.Top = Form1.Image1.Top + 32
    Form1.lblTileKind = "1"
    Form1.Shape2.Top = Form1.Shape2.Top + 32

End If
If MapExtra.Music = "" Then MapExtra.Music = "z5oot[2].mid"

Form1.MediaPlayer1.Filename = App.Path & "\" & MapExtra.Music
Form1.lblPlayerName(0).Top = Form1.Image1.Top - 20
Form1.lblPlayerName(0).Left = Form1.Image1.Left
ShopWindowOpen = False
Unload Me

End Sub

Private Sub vsNPC_Change()

DrawShopitems
End Sub

Private Sub vsNPC_Scroll()

DrawShopitems
End Sub


Private Sub sLoadPlayerItems()
For I = 0 To 14
'clear all the picture boxes
    picPlayerItem(I).Picture = Nothing
    picPlayerItem(I).Tag = ""
    picPlayerItem(I).ToolTipText = ""
Next

ReDim PlayerItem(UBound(PXItem))
For I = 0 To UBound(PXItem)

    'load first 15 items into the boxes
    For pos = 0 To UBound(PXItem)
    
    If PXItem(I).ItemName = "" Or PXItem(I).Quanity <= 0 Then
   '' If PXItem(pos).ItemName = "" Or PXItem(pos).Quanity <= 0 Then
    Else
        If pos <= 14 Then
        
            If picPlayerItem(pos).ToolTipText = "" Then
      ''  PlayerItem(pos).Name = PXItem(pos).ItemName
       '' PlayerItem(pos).ScrollPos = pos
       '' PlayerItem(pos).Value = PXItem(pos).Value
       '' picPlayerItem(pos).Picture = LoadPicture(App.Path & "\monsters\" & PXItem(pos).ItemName & ".gif")
       ''picPlayerItem(pos).ToolTipText = PXItem(pos).ItemName
        'Store the Quanity of each item
       '' picPlayerItem(pos).Tag = PXItem(pos).Quanity
       '' PlayerItem(pos).Number = pos
        
        PlayerItem(I).Name = PXItem(I).ItemName
        PlayerItem(I).ScrollPos = pos
        PlayerItem(I).Value = PXItem(I).Value
       picPlayerItem(pos).Picture = LoadPicture(App.Path & "\monsters\" & PXItem(I).ItemName & ".gif")
       picPlayerItem(pos).ToolTipText = PXItem(I).ItemName
        'Store the Quanity of each item
        picPlayerItem(pos).Tag = PXItem(I).Quanity
        PlayerItem(I).Number = I
        Exit For
        
          End If
        Else
            PlayerItem(pos).Name = PXItem(pos).ItemName
       PlayerItem(pos).ScrollPos = pos
      PlayerItem(pos).Value = PXItem(pos).Value
      PlayerItem(pos).Number = pos
      
        'PlayerItem(I).Name = PXItem(I).ItemName
        'PlayerItem(I).ScrollPos = pos
        'PlayerItem(I).Value = PXItem(I).Value
        'PlayerItem(I).Number = I
     '   MsgBox PlayerItem(pos).Name & " " & pos
        End If
    
      End If
    Next pos

Next I
Call DrawQty
'set the scrollbars max values
'vsPlayer.Max = MaxItems \ 5

vsPlayer.Max = MaxItems

End Sub
Private Sub DrawQty()
For I = 0 To 14
    picPlayerItem(I).CurrentX = 10
    picPlayerItem(I).CurrentY = -100
    picPlayerItem(I).Forecolor = vbBlack
    picPlayerItem(I).FontSize = 12
    picPlayerItem(I).FontBold = True
    picPlayerItem(I).AutoRedraw = True
    picPlayerItem(I).Print picPlayerItem(I).Tag
Next

End Sub

Private Sub vsPlayer_Change()
'Me.Caption = vsPlayer.Value & " " & PlayerItem(vsPlayer.Value).Name
Call DrawPlayerItems
End Sub
Private Sub vsPlayer_Scroll()
'Me.Caption = vsPlayer.Value & " " & PlayerItem(vsPlayer.Value).Name
Call DrawPlayerItems
End Sub

Private Sub DrawPlayerItems()
'Called when the scrollbar is scrolled or changed
For I = 0 To 14
'clear all the picture boxes
    picPlayerItem(I).Picture = Nothing
    picPlayerItem(I).Tag = ""
    picPlayerItem(I).ToolTipText = ""
Next

For pos = 0 To UBound(PlayerItem)
    If PlayerItem(pos).ScrollPos = vsPlayer.Value Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(0).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(0).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(0).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 1 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(1).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(1).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(1).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 2 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(2).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(2).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(2).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 3 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(3).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(3).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(3).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 4 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(4).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(4).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(4).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 5 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(5).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(5).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(5).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 6 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(6).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(6).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(6).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 7 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(7).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(7).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(7).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 8 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(8).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(8).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(8).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 9 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(9).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(9).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(9).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 10 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(10).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(10).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(10).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 11 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(11).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(11).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(11).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 12 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(12).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(12).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(12).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 13 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(13).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(13).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(13).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
    If PlayerItem(pos).ScrollPos = vsPlayer.Value + 14 Then
        If PlayerItem(pos).Name = "" Then
        Else
        picPlayerItem(14).Picture = LoadPicture(App.Path & "\monsters\" & PlayerItem(pos).Name & ".gif")
        picPlayerItem(14).ToolTipText = PlayerItem(pos).Name
        'Store the Quanity of each item
        picPlayerItem(14).Tag = PXItem(PlayerItem(pos).Number).Quanity
        End If
    End If
Next pos
'Update Qtys
Call DrawQty


End Sub
Private Sub DrawShopitems()
'Called when the scrollbar is scrolled or changed
For I = 0 To 14
'clear all the picture boxes
    picNPCItem(I).Picture = Nothing
    picNPCItem(I).Tag = ""
    picNPCItem(I).ToolTipText = ""
Next

For pos = 0 To UBound(StoreItem)
    If StoreItem(pos).ScrollPos = vsNPC.Value Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(0).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(0).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(0).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 1 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(1).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(1).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(1).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 2 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(2).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(2).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(2).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 3 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(3).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(3).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(3).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 4 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(4).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(4).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(4).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 5 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(5).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(5).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(5).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 6 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(6).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(6).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(6).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 7 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(7).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(7).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(7).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 8 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(8).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(8).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(8).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 9 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(9).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(9).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(9).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 10 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(10).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(10).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(10).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 11 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(11).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(11).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(11).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 12 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(12).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(12).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(12).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 13 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(13).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(13).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(13).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
    If StoreItem(pos).ScrollPos = vsNPC.Value + 14 Then
        If StoreItem(pos).Name = "" Then
        Else
        picNPCItem(14).Picture = LoadPicture(App.Path & "\monsters\" & StoreItem(pos).Name & ".gif")
        picNPCItem(14).ToolTipText = StoreItem(pos).Name
        'Store the Quanity of each item
        picNPCItem(14).Tag = PXItem(StoreItem(pos).Number).Quanity
        End If
    End If
Next pos



End Sub



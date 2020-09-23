VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmShortcuts 
   BackColor       =   &H80000005&
   Caption         =   "Shortcuts"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShortcuts.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1725
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":2394
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":3070
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":4A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":6394
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":7D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":96B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":A392
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":B06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":BD46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":CA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":D6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":DFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":ECB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":F992
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1066E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":10F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":11C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1250A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":131E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":14B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcuts.frx":1650E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMenu 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   6376
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MousePointer    =   99
      MouseIcon       =   "frmShortcuts.frx":16DEA
      OLEDragMode     =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub CommandPass(ByVal srcPerformWhat As String)
    Select Case srcPerformWhat
        Case "New"
            '
        Case "Edit"
            frmAbout.show vbModal
    End Select
End Sub


Private Sub Active()
    HighlightInWin Me.Name: MAIN.ShowTBButton "ttfffff"
    With MAIN
        .tbMenu.Buttons(3).Caption = "User's Guide"
        .tbMenu.Buttons(3).Image = 10
        
        .tbMenu.Buttons(4).Caption = "About"
        .tbMenu.Buttons(4).Image = 11
        
        .mnuRACN.Caption = "User's Guide"
        .mnuRAES.Caption = "About"
    End With
End Sub

Private Sub Deactive()
    MAIN.HideTBButton "", True
    With MAIN
        .tbMenu.Buttons(3).Caption = "New"
        .tbMenu.Buttons(3).Image = 1
        
        .tbMenu.Buttons(4).Caption = "Edit"
        .tbMenu.Buttons(4).Image = 2
    
        .mnuRACN.Caption = "Create New"
        .mnuRAES.Caption = "Edit Selected"
    End With
End Sub

Private Sub Form_Activate()
    Active
    HighlightInWin Name
End Sub

Private Sub Form_Deactivate()
    Deactive
End Sub

Private Sub Form_Load()
    
    With lvMenu
        Set .SmallIcons = ImageList1
        Set .Icons = ImageList1
        'For Clients
        .ListItems.Add , "frmClients", "Manage Clients", 1, 1
        
        'For Vendors
        .ListItems.Add , "frmVendors", "Manage Vendors", 4, 4
        
        'For Stocks
        .ListItems.Add , "frmStocks", "Stocks", 6, 6
        
        'For Sales
        .ListItems.Add , "frmCashSales", "Cash Sales", 14, 14
        .ListItems.Add , "frmSalesOrder", "Sales Order", 14, 14
        .ListItems.Add , "frmSalesOrderDelivery", "Sales Order Delivery", 14, 14
        
        'For Purchases
        .ListItems.Add , "frmLocalPurchase", "Local Purchase", 6, 6
        .ListItems.Add , "frmPurchaseOrder", "Purchase Order", 6, 6
        .ListItems.Add , "frmPurchaseOrderReceive", "Receive Stock", 6, 6

    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Beep: Cancel = 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvMenu.Width = ScaleWidth
    lvMenu.Height = ScaleHeight
End Sub

Private Sub lvMenu_DblClick()
    Select Case lvMenu.SelectedItem.Key
        'For Clients
        Case "frmClients": LoadForm frmClients
        
        'For Vendors
        Case "frmVendors": LoadForm frmVendors
        
        'For Stocks
        Case "frmStocks": LoadForm frmStocks
        
        'For Sales
        Case "frmCashSales": frmCashSalesAE.show 1
        Case "frmSalesOrder": LoadForm frmSalesOrder
        Case "frmSalesOrderDelivery": LoadForm frmSalesOrderDelivery
        
        'For Purchases
        Case "frmLocalPurchase": LoadForm frmLocalPurchase
        Case "frmPurchaseOrder": LoadForm frmPurchaseOrder
        Case "frmPurchaseOrderReceive": LoadForm frmPurchaseOrderReceive
    End Select

End Sub

Private Sub lvMenu_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then lvMenu_DblClick
End Sub

Private Sub lvMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu MAIN.mnuSO
End Sub

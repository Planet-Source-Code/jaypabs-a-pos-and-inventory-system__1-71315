VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Lookup Table "
   ClientHeight    =   5055
   ClientLeft      =   1050
   ClientTop       =   3495
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6135
   Begin MSComctlLib.ListView lvList 
      Height          =   4185
      Left            =   60
      TabIndex        =   2
      Top             =   570
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7382
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Barcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Qty-OnHand"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.TextBox txtItemDesc 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   180
      Width           =   4905
   End
   Begin VB.Label Label2 
      Caption         =   "Esc (Exit)"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   4800
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Item Desc"
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   1065
   End
End
Attribute VB_Name = "frmItemLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  RetreiveProducts txtItemDesc.Text
End Sub

Private Sub lvList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then Unload Me
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    frmCashSalesAE.txtEntry(7).Text = lvList.ListItems(lvList.SelectedItem.Index).SubItems(1)
    txtItemDesc.Text = lvList.ListItems(lvList.SelectedItem.Index).SubItems(1)
    Unload Me
  End If
End Sub

Private Sub txtItemDesc_Change()
  RetreiveProducts txtItemDesc.Text
End Sub

Private Sub txtItemDesc_GotFocus()
  HLText txtItemDesc
End Sub

Private Sub txtItemDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then Unload Me
End Sub

Private Sub RetreiveProducts(ByVal Product As String)
  Dim sql As String
  Dim rstemp As New Recordset
  
  sql = "SELECT Stocks.Barcode, Stocks.Stock, Stocks.OnHand " _
  & "From Stocks " _
  & "Where (((Stocks.Stock) Like '%" & Replace(Product, "'", "''") & "%')) " _
  & "ORDER BY Stocks.Stock"
  
  rstemp.Open sql, CN, adOpenDynamic, adLockOptimistic
  lvList.ListItems.Clear
  
  Do While Not rstemp.EOF
    lvList.ListItems.Add
    lvList.ListItems(lvList.ListItems.Count).SubItems(1) = rstemp!Barcode
    lvList.ListItems(lvList.ListItems.Count).SubItems(2) = rstemp!Stock
    lvList.ListItems(lvList.ListItems.Count).SubItems(3) = rstemp!OnHand
    rstemp.MoveNext
  Loop
  
  
  rstemp.Close
  Set rstemp = Nothing
End Sub

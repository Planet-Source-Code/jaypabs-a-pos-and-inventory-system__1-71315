VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStocksAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Entry"
   ClientHeight    =   6945
   ClientLeft      =   1425
   ClientTop       =   4725
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBegInv 
      Caption         =   "Beg. Inv."
      Height          =   315
      Left            =   5580
      TabIndex        =   10
      ToolTipText     =   "Beginning Inventory"
      Top             =   6510
      Width           =   1065
   End
   Begin VB.ComboBox cboPieces 
      Height          =   315
      ItemData        =   "frmStocksAE.frx":0000
      Left            =   7200
      List            =   "frmStocksAE.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1200
      Width           =   675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Barcode"
      Height          =   315
      Left            =   7950
      TabIndex        =   20
      Top             =   1200
      Width           =   1125
   End
   Begin VB.PictureBox picBarcode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   6780
      ScaleHeight     =   645
      ScaleWidth      =   1785
      TabIndex        =   19
      Top             =   330
      Width           =   1785
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   5
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1860
      Width           =   660
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   6510
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   4
      Left            =   1410
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1515
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   3
      Left            =   1410
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1170
      Width           =   3165
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   2
      Left            =   1410
      MaxLength       =   200
      TabIndex        =   2
      Top             =   825
      Width           =   3165
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   1
      Left            =   1410
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   480
      Width           =   4005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8205
      TabIndex        =   9
      Top             =   6510
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   6765
      TabIndex        =   8
      Top             =   6510
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1410
      MaxLength       =   13
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1965
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   11
      Top             =   6405
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcCategory 
      Height          =   315
      Left            =   1410
      TabIndex        =   6
      Top             =   2220
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   90
      TabIndex        =   23
      Top             =   2730
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Product Measures"
      TabPicture(0)   =   "frmStocksAE.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "nsdUnit"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Grid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtSalesPrice"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btnRemove"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtOrder"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtOnHand"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAdd"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtQty"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSupplierPrice"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txtSupplierPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3870
         MaxLength       =   10
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   690
         Width           =   1200
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         MaxLength       =   10
         TabIndex        =   29
         Top             =   690
         Width           =   540
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   5880
         TabIndex        =   28
         Top             =   690
         Width           =   495
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5100
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "0"
         Top             =   690
         Width           =   720
      End
      Begin VB.TextBox txtOrder 
         Height          =   315
         Left            =   150
         TabIndex        =   26
         Top             =   690
         Width           =   585
      End
      Begin VB.CommandButton btnRemove 
         Height          =   275
         Left            =   210
         Picture         =   "frmStocksAE.frx":0020
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Remove"
         Top             =   1200
         Visible         =   0   'False
         Width           =   275
      End
      Begin VB.TextBox txtSalesPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         MaxLength       =   10
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   690
         Width           =   960
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   2220
         Left            =   150
         TabIndex        =   31
         Top             =   1080
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   275
         ForeColorFixed  =   -2147483640
         BackColorSel    =   1091552
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdUnit 
         Height          =   315
         Left            =   1380
         TabIndex        =   32
         Top             =   690
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Supplier Price"
         Height          =   255
         Left            =   3870
         TabIndex        =   38
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Qty"
         Height          =   285
         Left            =   810
         TabIndex        =   37
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "Unit"
         Height          =   285
         Left            =   1410
         TabIndex        =   36
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "On-Hand"
         Height          =   255
         Left            =   5070
         TabIndex        =   35
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "Order"
         Height          =   255
         Left            =   180
         TabIndex        =   34
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Price"
         Height          =   255
         Left            =   2850
         TabIndex        =   33
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Pieces:"
      Height          =   255
      Left            =   6270
      TabIndex        =   21
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Reorder Point"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   11
      Left            =   60
      TabIndex        =   17
      Top             =   2220
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "ICode"
      Height          =   240
      Index           =   4
      Left            =   -90
      TabIndex        =   16
      Top             =   1515
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Short2"
      Height          =   240
      Index           =   3
      Left            =   -90
      TabIndex        =   15
      Top             =   1170
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Short1"
      Height          =   240
      Index           =   2
      Left            =   -90
      TabIndex        =   14
      Top             =   825
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stock"
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Barcode"
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   135
      Width           =   915
   End
End
Attribute VB_Name = "frmStocksAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode
Public srcTextAdd           As TextBox 'Used in pop-up mode -> Display the customer address
Public srcTextCP            As TextBox 'Used in pop-up mode -> Display the customer contact person
Public srcTextDisc          As Object  'Used in pop-up mode -> Display the customer Discount (can be combo or textbox)

Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset
Dim RSStockUnit             As New Recordset

Dim DocumentDB As DAO.Database
Dim sFileName As String
'Dim m_sBarcode As String, m_lBarcodeLength As Long
Private WithEvents HO As clsBarcode
Attribute HO.VB_VarHelpID = -1

Private Sub DisplayForEditing()
On Error GoTo err
    
    With rs
      txtEntry(0).Text = .Fields("Barcode")
      txtEntry(1).Text = .Fields("Stock")
      txtEntry(2).Text = .Fields("Short1")
      txtEntry(3).Text = .Fields("Short2")
      txtEntry(4).Text = .Fields("ICode")
      txtEntry(5).Text = toNumber(.Fields("ReorderPoint"))
      dcCategory.BoundText = .Fields![CategoryID]
    End With
    
    'Display the details
    Dim RSStockUnit As New Recordset

    cIRowCount = 0
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID=" & PK & " Order by [Order] ASC", CN, adOpenStatic, adLockOptimistic
    
    If RSStockUnit.RecordCount > 0 Then
        RSStockUnit.MoveFirst
        While Not RSStockUnit.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 7) = "" Then
                    .TextMatrix(1, 1) = RSStockUnit![Order]
                    .TextMatrix(1, 2) = RSStockUnit![Qty]
                    .TextMatrix(1, 3) = RSStockUnit![Unit]
                    .TextMatrix(1, 4) = toMoney(RSStockUnit![SalesPrice])
                    .TextMatrix(1, 5) = toMoney(RSStockUnit![SupplierPrice])
                    .TextMatrix(1, 6) = RSStockUnit![Onhand]
                    .TextMatrix(1, 7) = RSStockUnit![UnitID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSStockUnit![Order]
                    .TextMatrix(.Rows - 1, 2) = RSStockUnit![Qty]
                    .TextMatrix(.Rows - 1, 3) = RSStockUnit![Unit]
                    .TextMatrix(.Rows - 1, 4) = toMoney(RSStockUnit![SalesPrice])
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSStockUnit![SupplierPrice])
                    .TextMatrix(.Rows - 1, 6) = RSStockUnit![Onhand]
                    .TextMatrix(.Rows - 1, 7) = RSStockUnit![UnitID]
                End If
            End With
            RSStockUnit.MoveNext
        Wend
        Grid.Row = 1
        Grid.Colsel = 6
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    RSStockUnit.Close
    'Clear variables
    Set RSStockUnit = Nothing
    
    Exit Sub
    
err:
    If err.Number = 94 Then Resume Next
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.Rowsel)
    End With

    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtOrder.Text) = "" Or Trim(txtQty.Text) = "" Or Trim(nsdUnit.Text) = "" Then Exit Sub

    Dim CurrRow As Integer
    Dim intUnitID As Integer
    
    CurrRow = getFlexPos(Grid, 7, nsdUnit.Tag)

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 7) = "" Then
                .TextMatrix(1, 1) = txtOrder.Text
                .TextMatrix(1, 2) = txtQty.Text
                .TextMatrix(1, 3) = nsdUnit.Text
                .TextMatrix(1, 4) = toMoney(txtSalesPrice.Text)
                .TextMatrix(1, 5) = toMoney(txtSupplierPrice.Text)
                .TextMatrix(1, 6) = txtOnHand.Text
                .TextMatrix(1, 7) = nsdUnit.Tag 'intUnitID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = txtOrder.Text
                .TextMatrix(.Rows - 1, 2) = txtQty.Text
                .TextMatrix(.Rows - 1, 3) = nsdUnit.Text
                .TextMatrix(.Rows - 1, 4) = toMoney(txtSalesPrice.Text)
                .TextMatrix(.Rows - 1, 5) = toMoney(txtSupplierPrice.Text)
                .TextMatrix(.Rows - 1, 6) = txtOnHand.Text
                .TextMatrix(.Rows - 1, 7) = nsdUnit.Tag 'intUnitID

                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = txtOrder.Text
                .TextMatrix(CurrRow, 2) = txtQty.Text
                .TextMatrix(CurrRow, 3) = nsdUnit.Text
                .TextMatrix(CurrRow, 4) = toMoney(txtSalesPrice.Text)
                .TextMatrix(CurrRow, 5) = toMoney(txtSupplierPrice.Text)
                .TextMatrix(CurrRow, 6) = txtOnHand.Text
                .TextMatrix(CurrRow, 7) = nsdUnit.Tag 'intUnitID
            Else
                Exit Sub
            End If
        End If
        
        'Highlight the current row's column
        .Colsel = 6
        'Display a remove button
        Grid_Click
    End With
End Sub

Private Sub cmdBegInv_Click()
    With frmBegInv
        .intStockID = PK
        
        .show 1
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    txtEntry(0).SetFocus
End Sub

Private Sub cmdPrint_Click()
    Dim i As Integer
    Dim intPieces As Integer
    
    CN.Execute "DELETE * FROM Barcode_Temp"
    
    intPieces = CInt(cboPieces.Text) / 5 + 1
    
    For i = 1 To intPieces
        CN.Execute "INSERT INTO Barcode_Temp " _
                & "SELECT qry_Barcodes.* " _
                & "FROM qry_Barcodes " _
                & "WHERE StockID = " & PK
    Next i
    
    DoEvents

    With frmReports
        .strReport = "Print Barcode"
        .strWhere = "{Barcode_Temp.StockID} = " & PK
        
        frmReports.show vbModal
    End With
End Sub

Private Sub cmdSave_Click()
On Error GoTo err
  Dim rsStockUnits As New Recordset
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    
    If Len(txtEntry(0).Text) < 13 Then
        MsgBox "Barcode must be 13 digit", vbInformation
        Exit Sub
    End If
    
    CN.BeginTrans
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("StockID") = PK
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With rs
      .Fields("Barcode") = txtEntry(0).Text
      .Fields("Stock") = txtEntry(1).Text
      .Fields("Short1") = txtEntry(2).Text
      .Fields("Short2") = txtEntry(3).Text
      .Fields("ICode") = txtEntry(4).Text
      .Fields("ReorderPoint") = toNumber(txtEntry(5).Text)
      .Fields("CategoryID") = dcCategory.BoundText
      .Update
    End With
    
    Dim RSStockUnit As New Recordset

    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM Stock_Unit WHERE StockID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                RSStockUnit.AddNew

                RSStockUnit![StockID] = PK
                RSStockUnit![Order] = toNumber(.TextMatrix(c, 1))
                RSStockUnit![UnitID] = toNumber(.TextMatrix(c, 7))
                RSStockUnit![Qty] = toNumber(.TextMatrix(c, 2))
                RSStockUnit![SalesPrice] = toNumber(.TextMatrix(c, 4))
                RSStockUnit![SupplierPrice] = toNumber(.TextMatrix(c, 5))
                RSStockUnit![Onhand] = toNumber(.TextMatrix(c, 6))

                RSStockUnit.Update
            ElseIf State = adStateEditMode Then
                RSStockUnit.Filter = "UnitID = " & toNumber(.TextMatrix(c, 7))
            
                If RSStockUnit.RecordCount = 0 Then GoTo AddNew

                RSStockUnit![Order] = toNumber(.TextMatrix(c, 1))
                RSStockUnit![UnitID] = toNumber(.TextMatrix(c, 7))
                RSStockUnit![Qty] = toNumber(.TextMatrix(c, 2))
                RSStockUnit![SalesPrice] = toNumber(.TextMatrix(c, 4))
                RSStockUnit![SupplierPrice] = toNumber(.TextMatrix(c, 5))
                RSStockUnit![Onhand] = toNumber(.TextMatrix(c, 6))

                RSStockUnit.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set RSStockUnit = Nothing
    
    CN.CommitTrans
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
    
  Exit Sub
err:
  CN.RollbackTrans
  MsgBox "Error: " & err.Description, vbExclamation
  'If err.Number = -2147217887 Then Resume Next
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(rs.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Stocks")
End Sub

Private Sub Form_Load()
    InitGrid
    InitNSD

    Set DocumentDB = DAO.Workspaces(0).OpenDatabase(App.Path & "\Data\db.mdb", False, False, ";PWD=jaypee")
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Stocks WHERE StockID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Stocks_Category ORDER BY Category ASC", "Category", dcCategory, "CategoryID", True
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        GeneratePK
        
        CreateBarcode
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        LoadBarcode
    End If
    
    Dim i As Integer
    
    For i = 5 To 55 Step 5
        cboPieces.AddItem i
    Next i
    cboPieces.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmStocks.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmStocksAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        txtOrder.Text = .TextMatrix(.Rowsel, 1)
        txtQty.Text = .TextMatrix(.Rowsel, 2)
        nsdUnit.Text = .TextMatrix(.Rowsel, 3)
        nsdUnit.Tag = .TextMatrix(.Rowsel, 7) 'Add tag coz boundtext is empty
        txtSalesPrice.Text = .TextMatrix(.Rowsel, 4)
        txtSupplierPrice.Text = .TextMatrix(.Rowsel, 5)
        txtOnHand.Text = .TextMatrix(.Rowsel, 6)
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 7) = "" Then '7 = StockID
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub nsdUnit_Change()
    nsdUnit.Tag = nsdUnit.BoundText
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub

Private Sub CreateBarcode()
    Dim Pos As Long
    Dim tempEven As Long
    Dim tempOdd As Long
    Dim tempTotal
    Dim newBarCode As String
    Dim Checksum As Integer
    Randomize
    
    'newBarCode = Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10)
    
    If State = adStateAddMode Then
        newBarCode = Format(getIndex("Barcode"), "000000000000")
    
        For Pos = 2 To 12 Step 2
            tempEven = tempEven + Val(Mid(newBarCode, Pos, 1))
        Next
    
        For Pos = 1 To 11 Step 2
            tempOdd = tempOdd + Val(Mid(newBarCode, Pos, 1))
        Next
    
        tempEven = tempEven * 3
        tempTotal = tempOdd + tempEven
        Checksum = tempTotal Mod 10
        
        If Checksum > 0 Then
            Checksum = 10 - Checksum
        End If
        
        newBarCode = newBarCode & Checksum
        If Checksum <> Mid(newBarCode, 13, 1) Then
            MsgBox "Invalid number"
            Exit Sub
        End If
        
        txtEntry(0).Text = newBarCode
    End If

'    Picture2.Left = Picture1.Width / 3 - 100
    Call Barcoder(txtEntry(0).Text, picBarcode)
    
    SavePicture picBarcode.Image, App.Path & "\Barcodes\" & txtEntry(0).Text & ".jpg"
    sFileName = App.Path & "\Barcodes\" & txtEntry(0).Text & ".jpg"
    SaveBinaryObject
End Sub

Private Sub SaveBinaryObject()
    Dim FieldNames(2) As Variant           'names of the other fields to return
    Dim FieldData(2) As Variant            'names of the other fields to return
    Dim RD() As Variant                    'store for the returned data, not the binary field
    Dim FN As String                       'Binary file name to use as storage
    Dim i As Integer

    If sFileName = "" Then
        Exit Sub
    End If

    Set HO = New clsBarcode            'create the new bd object

    FieldNames(0) = "ID"               'return the ID field
    FieldNames(1) = "FileName"         'return the filename
    FieldNames(2) = "Barcode"
    
    
    FieldData(0) = PK 'Null                  'return the ID field
    FieldData(1) = sFileName           'return the filename
    FieldData(2) = txtEntry(0).Text
    
    With HO
        .KillFile = False                      'kill the filename if it exists
        Set .DB = DocumentDB                   'pass the database
        .ObjectKeyFieldName = "ID"             'the key/index field is
        If State = adStateAddMode Then
            .ObjectKey = -1
        Else
            .ObjectKey = PK                    'the value to search for is
        End If
        .ObjectFieldName = "OLEModule"         'name of the field that contains the binary file
        .ObjectTableName = "Barcode"           'table that contains the binary files
        .SubFieldNames = FieldNames            'pass in the field names to return
        .SubFieldData = FieldData
        .FileName = sFileName                  'file name to use
        .SaveObject                            'get the file from the database
        .ReturnData RD()                       'return any aditional data
        FN = .FileName                         'actual file name used - if default was used
    End With
    Set HO = Nothing

    For i = 0 To UBound(RD)
        Debug.Print RD(i)                      'print aditional info returned
    Next
End Sub

Private Sub LoadBarcode()
On Error GoTo err_GetObject
    Dim FN As String
    
    FN = App.Path & "\Barcodes\" & txtEntry(0).Text & ".jpg"
    picBarcode.Picture = LoadPicture(FN)
    
    Exit Sub
    
err_GetObject:
    If err.Number = 53 Then
        CreateBarcode
    Else
        MsgBox err.Number & " " & err.Description, vbInformation
    End If
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 8
        .Colsel = 7
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 900
        .ColWidth(5) = 1200
        .ColWidth(6) = 900
        .ColWidth(7) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Order"
        .TextMatrix(0, 2) = "Qty"
        .TextMatrix(0, 3) = "Unit"
        .TextMatrix(0, 4) = "Sales Price"
        .TextMatrix(0, 5) = "Supplier Price"
        .TextMatrix(0, 6) = "On Hand"
        .TextMatrix(0, 7) = "Unit ID"
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
'        .ColAlignment(1) = vbLeftJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = flexAlignGeneral
'        .ColAlignment(4) = flexAlignGeneral
'        .ColAlignment(5) = vbRightJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
'        .ColAlignment(8) = vbRightJustify
    End With
End Sub

Private Sub InitNSD()
    'For Vendor
    With nsdUnit
        .ClearColumn
        .AddColumn "Unit ID", 1794.89
        .AddColumn "Unit", 2264.88
        .Connection = CN.ConnectionString
        
        '.sqlFields = "VendorID, Company, Location"
        .sqlFields = "UnitID, Unit"
        .sqlTables = "Unit"
        .sqlSortOrder = "Unit ASC"
        
        .BoundField = "UnitID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Units Record"
    End With
    
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim rsUnit As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    rsUnit.CursorLocation = adUseClient
    rsUnit.Open "SELECT * FROM Stock_Unit WHERE StockID=" & PK, CN, adOpenStatic, adLockOptimistic
    If rsUnit.RecordCount > 0 Then
        rsUnit.MoveFirst
        While Not rsUnit.EOF
            CurrRow = getFlexPos(Grid, 7, rsUnit!UnitID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Stock_Unit", "StockUnitID", "", True, rsUnit!StockUnitID
                End If
            End With
            rsUnit.MoveNext
        Wend
    End If
End Sub

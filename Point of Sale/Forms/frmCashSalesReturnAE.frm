VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCashSalesReturnAE 
   BorderStyle     =   0  'None
   Caption         =   "CashSales"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   360
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   210
      ScaleHeight     =   630
      ScaleWidth      =   10740
      TabIndex        =   34
      Top             =   2250
      Width           =   10740
      Begin VB.TextBox txtStock 
         Height          =   255
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   2685
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   3480
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   5775
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox txtUnitPrice 
         Height          =   285
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   2775
         TabIndex        =   38
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9000
         TabIndex        =   37
         Top             =   225
         Width           =   840
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDisc 
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   5775
         TabIndex        =   47
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Qty"
         Height          =   240
         Index           =   10
         Left            =   2775
         TabIndex        =   46
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   240
         Index           =   9
         Left            =   4500
         TabIndex        =   45
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Items/Stocks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   43
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   8040
         TabIndex        =   42
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   7140
         TabIndex        =   41
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6900
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9540
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1425
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9540
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7500
      Width           =   1425
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1050
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtVoucherNo 
      Height          =   285
      Left            =   6090
      TabIndex        =   25
      Top             =   720
      Width           =   2475
   End
   Begin VB.TextBox txtRemarks 
      Height          =   1335
      Left            =   225
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Tag             =   "Remarks"
      Top             =   6870
      Width           =   5910
   End
   Begin VB.TextBox txtGross 
      BackColor       =   &H00E6FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmCashSalesReturnAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Remove"
      Top             =   3390
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   8160
      TabIndex        =   7
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9615
      TabIndex        =   6
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   5
      Top             =   8490
      Width           =   1755
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7950
      Width           =   1425
   End
   Begin VB.TextBox txtSoldTo 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1065
      Width           =   2475
   End
   Begin VB.TextBox txtInvoiceNo 
      Height          =   285
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   2475
   End
   Begin VB.TextBox txtBusinessName 
      Height          =   285
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1380
      Width           =   2475
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1710
      Width           =   2475
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   11
      Top             =   8340
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   3270
      Left            =   210
      TabIndex        =   12
      Top             =   3240
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   5768
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   6090
      TabIndex        =   24
      Top             =   1050
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   94961667
      CurrentDate     =   38207
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Purchase Return Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   240
      TabIndex        =   18
      Top             =   2970
      Width           =   4365
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7440
      TabIndex        =   33
      Top             =   6930
      Width           =   2040
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Tax Base"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7440
      TabIndex        =   32
      Top             =   7230
      Width           =   2040
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Vat(0.12)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7440
      TabIndex        =   31
      Top             =   7530
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Sales Return"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   13
      Top             =   150
      Width           =   4905
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Voucher No.:"
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   750
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      Height          =   8235
      Left            =   120
      Top             =   630
      Width           =   10935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   8895
      Left            =   60
      Top             =   60
      Width           =   11085
   End
   Begin VB.Label Labels 
      Caption         =   "Remarks"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   6600
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7440
      TabIndex        =   22
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   10935
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   10935
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7440
      TabIndex        =   21
      Top             =   7980
      Width           =   2040
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   20
      Top             =   1050
      Width           =   1485
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Sold To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   19
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   2940
      Width           =   10740
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   9210
      X2              =   10920
      Y1              =   7890
      Y2              =   7890
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Invoice No.:"
      Height          =   255
      Left            =   210
      TabIndex        =   17
      Top             =   750
      Width           =   1485
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5100
      TabIndex        =   16
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   255
      Left            =   210
      TabIndex        =   15
      Top             =   1740
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Business Name:"
      Height          =   255
      Left            =   210
      TabIndex        =   14
      Top             =   1380
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frmCashSalesReturnAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice
Dim intQtyOld               As Integer 'Allowed value for return qty

Private Sub btnUpdate_Click()
    Dim CurrRow As Integer

    CurrRow = getFlexPos(Grid, 10, Grid.TextMatrix(Grid.Rowsel, 10))

    'Add to grid
    With Grid
        .Row = CurrRow
                
        'Restore back the invoice amount and discount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.Rowsel, 7))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.Rowsel, 9))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        cDAmount = cDAmount - toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(Grid.TextMatrix(.Rowsel, 4)) * toNumber(txtUnitPrice.Text)))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")

        .TextMatrix(CurrRow, 4) = txtQty.Text
        .TextMatrix(CurrRow, 5) = cboUnit.Text
        .TextMatrix(CurrRow, 7) = toMoney(txtGross(1).Text)
        .TextMatrix(CurrRow, 9) = toMoney(toNumber(txtNetAmount.Text))
        
        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtGross(1).Text)
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount + toNumber(txtNetAmount.Text)
        cDAmount = cDAmount + toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text)))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        'Highlight the current row's column
        .Colsel = 10
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update grooss to current purchase amount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.Rowsel, 7))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.Rowsel, 9))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        cDAmount = cDAmount - toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(Grid.TextMatrix(.Rowsel, 4)) * toNumber(Grid.TextMatrix(.Rowsel, 6))))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)

        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.Rowsel)
    End With

    btnRemove.Visible = False
    Grid_Click
    
End Sub

Private Sub txtdisc_Change()
    txtQty_Change
End Sub

Private Sub txtdisc_Click()
    txtQty_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtDisc_GotFocus()
    HLText txtDisc
End Sub

Private Sub txtdisc_Validate(Cancel As Boolean)
    txtDisc.Text = toNumber(txtDisc.Text)
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
  If Trim(txtVoucherNo.Text) = "" Then
    MsgBox "Please enter voucher number before saving this record.", vbExclamation
    Exit Sub
  End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to return before saving this record.", vbExclamation
        Exit Sub
    End If
   
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    'Connection for Cash_Sales_Return
    Dim RSReturn As New Recordset

    RSReturn.CursorLocation = adUseClient
    RSReturn.Open "Cash_Sales_Return", CN, adOpenDynamic, adLockOptimistic, adCmdTable

    'Connection for Cash_Sales_Return_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "Cash_Sales_Return_Detail", CN, adOpenDynamic, adLockOptimistic, adCmdTable

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    'Save the record
    With RSReturn
        .AddNew
        Dim VoucherPK As Integer
        
        VoucherPK = getIndex("Cash_Sales_Return")
        ![VoucherID] = VoucherPK
        ![CashSalesID] = PK
        ![VoucherNo] = txtVoucherNo.Text
        ![Date] = dtpDate.Value
        ![Remarks] = txtRemarks.Text
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)
        
        ![DateAdded] = Now
        ![AddedByFK] = CurrUser.USER_PK
                
        .Update
    End With
   
    With Grid
        'Save to stock card
        Dim RSStockCard As New Recordset
    
        RSStockCard.CursorLocation = adUseClient
        RSStockCard.Open "Stock_Card", CN, , adLockOptimistic, adCmdTable
        
        'Save to stocks table
        Dim RSStocks As New Recordset
    
        RSStocks.CursorLocation = adUseClient
        RSStocks.Open "Stocks", CN, , adLockOptimistic
        
        'Save to Local Purchase Details
        Dim RSCashSalesDetails As New Recordset
    
        RSCashSalesDetails.CursorLocation = adUseClient
        RSCashSalesDetails.Open "SELECT * From Cash_Sales_Detail Where CashSalesID = " & PK, CN, , adLockOptimistic
        
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            RSDetails.AddNew

            RSDetails![VoucherID] = VoucherPK
            RSDetails![StockID] = toNumber(.TextMatrix(c, 10))
            RSDetails![Qty] = toNumber(.TextMatrix(c, 4))
            RSDetails![Unit] = GetUnitID(.TextMatrix(c, 5))
            RSDetails![Price] = toNumber(.TextMatrix(c, 6))
            RSDetails![Discount] = toNumber(.TextMatrix(c, 8)) / 100

            RSDetails.Update

            'Add record to stock card
            RSStockCard.AddNew
                
            RSStockCard!Type = "CPR"
            RSStockCard!RefNo2 = txtInvoiceNo.Text
            RSStockCard!Pieces2 = "-" & toNumber(.TextMatrix(c, 4))
            RSStockCard!Cost = toNumber(.TextMatrix(c, 6))
            RSStockCard!StockID = toNumber(.TextMatrix(c, 10))
                
            RSStockCard.Update

            'deduct qty returned in stocks
            RSStocks.Find "[StockID] = " & toNumber(.TextMatrix(c, 10)), , adSearchForward, 1
            RSStocks!Onhand = toNumber(RSStocks!Onhand) - toNumber(.TextMatrix(c, 4))
            
            RSStocks.Update
            
            'add qty returned in Local Purchase Details
            RSCashSalesDetails.Find "[StockID] = " & toNumber(.TextMatrix(c, 10)), , adSearchForward, 1
            RSCashSalesDetails!QtyReturned = toNumber(RSCashSalesDetails!QtyReturned) + toNumber(.TextMatrix(c, 4))
            
            RSCashSalesDetails.Update
        Next c
    End With

    'Clear variables
    c = 0
    Set RSDetails = Nothing

    CN.CommitTrans

    HaveAction = True
    Screen.MousePointer = vbDefault

    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    CN.RollbackTrans
    Prompt_Err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tUser1 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: n/a" & vbCrLf & _
           "Modified By: n/a", vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tUser1 = vbNullString
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    Else
        txtInvoiceNo.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid

    loadUnit

    Screen.MousePointer = vbHourglass
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        'Set the recordset
        rs.Open "SELECT * FROM Cash_Sales WHERE CashSalesID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpDate.Value = Date
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
                   
        DisplayForEditing
    Else
        'Set the recordset
        rs.Open "SELECT * FROM qry_Cash_Sales_Return WHERE VoucherID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        cmdCancel.Caption = "Close"
        DisplayForViewing
        
        If ForCusAcc = True Then
            'Me.Icon = frmLocalPurchaseReturn.Icon
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
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
        .Cols = 11
        .Colsel = 10
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 2025
        .ColWidth(2) = 2505
        .ColWidth(3) = 1545
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 900
        .ColWidth(7) = 900
        .ColWidth(8) = 900
        .ColWidth(9) = 1545
        .ColWidth(10) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Barcode"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "ICode"
        .TextMatrix(0, 4) = "Unit Qty"
        .TextMatrix(0, 5) = "Unit"
        .TextMatrix(0, 6) = "Sales Price"
        .TextMatrix(0, 7) = "Gross"
        .TextMatrix(0, 8) = "Discount(%)"
        .TextMatrix(0, 9) = "Net Amount"
        .TextMatrix(0, 10) = "Stock ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbRightJustify
        .ColAlignment(5) = vbLeftJustify
        .ColAlignment(6) = vbRightJustify
        .ColAlignment(7) = vbRightJustify
        .ColAlignment(8) = vbRightJustify
        .ColAlignment(9) = vbRightJustify
    End With
End Sub

Private Sub ResetEntry()
    'nsdStock.ResetValue
    txtUnitPrice.Tag = 0
    txtUnitPrice.Text = "0.00"
    txtQty.Text = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCashSalesReturnAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        txtStock.Text = .TextMatrix(.Rowsel, 2)
        txtQty = .TextMatrix(.Rowsel, 4)
        cboFindList cboUnit, .TextMatrix(.Rowsel, 5)
        txtUnitPrice = toMoney(.TextMatrix(.Rowsel, 6))
        txtGross(1) = toMoney(.TextMatrix(.Rowsel, 7))
        txtDisc = toMoney(.TextMatrix(.Rowsel, 8))
        txtNetAmount = toMoney(.TextMatrix(.Rowsel, 9))
        
        If State = adStateEditMode Then Exit Sub
        If Grid.Rows = 2 And Grid.TextMatrix(1, 10) = "" Then
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtQty_LostFocus()
    If txtQty > intQtyOld Then
        MsgBox "Overreturn for " & txtStock.Text & ".", vbInformation
        txtQty.Text = intQtyOld
    End If
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtUnitPrice_Change()
    txtQty_Change
End Sub

Private Sub txtUnitPrice_Validate(Cancel As Boolean)
    txtUnitPrice.Text = toMoney(toNumber(txtUnitPrice.Text))
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnUpdate.Enabled = False
    Else
        btnUpdate.Enabled = True
    End If
    
    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text)))
    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text)) - ((toNumber(txtDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text))))
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
    
    intQtyOld = txtQty.Text
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Used to edit record
Private Sub DisplayForEditing()
    On Error GoTo err
    txtInvoiceNo.Text = rs![InvoiceNo]
    txtSoldTo.Text = rs![SoldTo]
    txtDate.Text = Format$(rs![Date], "MMM-dd-yyyy")
    txtBusinessName.Text = rs![BusinessName]
    txtAddress.Text = rs![Address]
    
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtTaxBase.Text = toMoney(rs![TaxBase])
    txtVat.Text = toMoney(rs![Vat])
    txtNet.Text = toMoney(rs![NetAmount])
    txtRemarks.Text = rs![Remarks]
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Cash_Sales_Detail WHERE CashSalesID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![ICode]
                    .TextMatrix(1, 4) = RSDetails![QtyDue]
                    .TextMatrix(1, 5) = RSDetails![Unit]
                    .TextMatrix(1, 6) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 7) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 10) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![ICode]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![QtyDue]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Price]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Gross]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 10) = RSDetails![StockID]
                End If
                cIRowCount = cIRowCount + 1
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.Colsel = 10
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    dtpDate.Visible = True
    txtDate.Visible = False

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo err
    
    txtInvoiceNo.Text = rs![InvoiceNo]
    txtSoldTo.Text = rs![SoldTo]
    txtBusinessName.Text = rs![BusinessName]
    txtAddress.Text = rs![Address]
    txtVoucherNo.Text = rs![VoucherNo]
    txtDate.Text = Format$(rs![Date], "MMM-dd-yyyy")
    
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtTaxBase.Text = toMoney(rs![TaxBase])
    txtVat.Text = toMoney(rs![Vat])
    txtNet.Text = toMoney(rs![NetAmount])
    'txtRemarks.Text = rs![Remarks]
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Cash_Sales_Return_Detail WHERE VoucherID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![ICode]
                    .TextMatrix(1, 4) = RSDetails![Qty]
                    .TextMatrix(1, 5) = RSDetails![Unit]
                    .TextMatrix(1, 6) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 7) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 10) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![ICode]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Price]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Gross]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 10) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.Colsel = 10
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    'Disable commands
    LockInput Me, True

    dtpDate.Visible = True
    txtDate.Visible = False
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnUpdate.Visible = False

    'Resize and reposition the controls
    Shape3.Top = 2100
    Label11.Top = 2100
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 2400
    Grid.Height = 4100
    
    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub txtUnitPrice_GotFocus()
    HLText txtUnitPrice
End Sub

Private Sub loadUnit()
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  
  SQL = "SELECT Unit From Unit ORDER BY Unit asc"
  
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  
  cboUnit.Clear
  
  Do While Not rs.EOF
    cboUnit.AddItem rs!Unit
    rs.MoveNext
  Loop
    
  rs.Close
  Set rs = Nothing
End Sub

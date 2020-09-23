VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmSalesOrderReturnAE 
   BorderStyle     =   0  'None
   Caption         =   "Sales Order Return"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDate 
      Height          =   314
      Left            =   8850
      TabIndex        =   68
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtCreditMemoDate 
      Height          =   314
      Left            =   1650
      TabIndex        =   67
      Top             =   1920
      Width           =   1905
   End
   Begin VB.TextBox txtApprovedBy 
      Height          =   314
      Left            =   5280
      TabIndex        =   54
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtPostedBy 
      Height          =   314
      Left            =   5280
      TabIndex        =   53
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtReceivedBy 
      Height          =   314
      Left            =   5280
      TabIndex        =   52
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtPreparedBy 
      Height          =   314
      Left            =   5280
      TabIndex        =   51
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtReason 
      Height          =   314
      Left            =   1650
      TabIndex        =   50
      Top             =   2280
      Width           =   1905
   End
   Begin VB.TextBox txtCreditMemoNo 
      Height          =   314
      Left            =   1650
      TabIndex        =   49
      Top             =   1560
      Width           =   1905
   End
   Begin VB.TextBox txtReference 
      Height          =   314
      Left            =   1650
      TabIndex        =   48
      Top             =   2640
      Width           =   1905
   End
   Begin VB.TextBox txtAcknowledgedBy 
      Height          =   314
      Left            =   8850
      TabIndex        =   47
      Top             =   1920
      Width           =   1935
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   210
      ScaleHeight     =   630
      ScaleWidth      =   10740
      TabIndex        =   13
      Top             =   3360
      Width           =   10740
      Begin VB.TextBox txtStock 
         Height          =   255
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   2685
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   4950
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   7035
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txtUnitPrice 
         Height          =   285
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   4245
         TabIndex        =   18
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9840
         TabIndex        =   17
         Top             =   225
         Width           =   840
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   285
         Left            =   8910
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtDisc 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtICode 
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   7035
         TabIndex        =   30
         Top             =   0
         Width           =   1050
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Qty"
         Height          =   240
         Index           =   10
         Left            =   4245
         TabIndex        =   29
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   240
         Index           =   9
         Left            =   5940
         TabIndex        =   28
         Top             =   0
         Width           =   1050
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
         TabIndex        =   27
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   4950
         TabIndex        =   26
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   8910
         TabIndex        =   25
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   8100
         TabIndex        =   24
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label19 
         Caption         =   "ICode"
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         Top             =   0
         Width           =   1425
      End
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6900
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9540
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1425
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9540
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7500
      Width           =   1425
   End
   Begin VB.TextBox txtRemarks 
      Height          =   1335
      Left            =   225
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
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
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   270
      Picture         =   "frmSalesOrderReturnAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Remove"
      Top             =   4200
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9615
      TabIndex        =   5
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   4
      Top             =   8490
      Width           =   1755
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7950
      Width           =   1425
   End
   Begin VB.TextBox txtSONo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   690
      Width           =   3315
   End
   Begin VB.TextBox txtDeliveryNo 
      Height          =   314
      Left            =   6570
      TabIndex        =   1
      Top             =   660
      Width           =   1905
   End
   Begin VB.TextBox txtDeliveryDate 
      Height          =   314
      Left            =   6570
      TabIndex        =   0
      Top             =   1020
      Width           =   1905
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   31
      Top             =   8340
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2490
      Left            =   180
      TabIndex        =   32
      Top             =   4050
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   4392
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
   Begin ctrlNSDataCombo.NSDataCombo nsdClient 
      Height          =   315
      Left            =   1500
      TabIndex        =   33
      Top             =   1050
      Width           =   3690
      _ExtentX        =   6509
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
   Begin MSComCtl2.DTPicker dtpCreditMemoDate 
      Height          =   315
      Left            =   1650
      TabIndex        =   55
      Top             =   1920
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd MMM, yyyy"
      Format          =   94830595
      CurrentDate     =   38989
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   8850
      TabIndex        =   56
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd MMM, yyyy"
      Format          =   94830595
      CurrentDate     =   38989
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "Approved By:"
      Height          =   255
      Left            =   4080
      TabIndex        =   66
      Top             =   2670
      Width           =   1155
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Posted By:"
      Height          =   255
      Left            =   4080
      TabIndex        =   65
      Top             =   2310
      Width           =   1155
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Received By:"
      Height          =   255
      Left            =   4080
      TabIndex        =   64
      Top             =   1950
      Width           =   1155
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Prepared By:"
      Height          =   255
      Left            =   4080
      TabIndex        =   63
      Top             =   1590
      Width           =   1155
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   255
      Left            =   8100
      TabIndex        =   62
      Top             =   1590
      Width           =   705
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Reason"
      Height          =   255
      Left            =   210
      TabIndex        =   61
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Memo Date:"
      Height          =   255
      Left            =   210
      TabIndex        =   60
      Top             =   1950
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Memo No.:"
      Height          =   255
      Left            =   210
      TabIndex        =   59
      Top             =   1590
      Width           =   1395
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Reference"
      Height          =   255
      Left            =   210
      TabIndex        =   58
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Acknowledged By:"
      Height          =   255
      Left            =   7440
      TabIndex        =   57
      Top             =   1950
      Width           =   1365
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Return Details"
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
      Left            =   210
      TabIndex        =   34
      Top             =   3150
      Width           =   4365
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Return"
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
      TabIndex        =   39
      Top             =   150
      Width           =   4905
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
      Top             =   7530
      Width           =   2040
   End
   Begin VB.Shape Shape1 
      Height          =   8235
      Left            =   120
      Top             =   600
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
      TabIndex        =   43
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
      TabIndex        =   42
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   10935
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   10935
      Y1              =   3090
      Y2              =   3090
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
      TabIndex        =   41
      Top             =   7980
      Width           =   2040
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   3150
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
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5100
      TabIndex        =   40
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendor"
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
      Left            =   180
      TabIndex        =   38
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "SO No."
      Height          =   255
      Left            =   180
      TabIndex        =   37
      Top             =   690
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   10905
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   180
      X2              =   10905
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label Label17 
      Caption         =   "Delivery Date"
      Height          =   255
      Left            =   5250
      TabIndex        =   36
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label18 
      Caption         =   "Delivery No."
      Height          =   255
      Left            =   5250
      TabIndex        =   35
      Top             =   690
      Width           =   1275
   End
End
Attribute VB_Name = "frmSalesOrderReturnAE"
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
  If Trim(txtCreditMemoNo.Text) = "" Then
    MsgBox "Please enter credit memo number before saving this record.", vbExclamation
    Exit Sub
  End If
   
  If cIRowCount < 1 Then
      MsgBox "Please enter item to return before saving this record.", vbExclamation
      Exit Sub
  End If
   
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    'Connection for Local_Purchase_Return
    Dim RSReturn As New Recordset

    RSReturn.CursorLocation = adUseClient
    RSReturn.Open "Sales_Order_Return", CN, adOpenDynamic, adLockOptimistic, adCmdTable

    'Connection for Sales_Order_Return_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "Sales_Order_Return_Detail", CN, adOpenDynamic, adLockOptimistic, adCmdTable

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    'Save the record
    With RSReturn
        .AddNew
        Dim SalesOrderReturnPK As Integer
        
        SalesOrderReturnPK = getIndex("Sales_Order_Return")
        ![SalesOrderReturnID] = SalesOrderReturnPK
        ![DeliveryID] = PK
        ![CreditMemoNo] = txtCreditMemoNo.Text
        ![CreditMemoDate] = dtpCreditMemoDate.Value
        ![Reason] = txtReason.Text
        ![Reference] = txtReference.Text
        ![PreparedBy] = txtPreparedBy.Text
        ![ReceivedBy] = txtReceivedBy.Text
        ![PostedBy] = txtPostedBy.Text
        ![ApprovedBy] = txtApprovedBy.Text
        ![Date] = dtpDate.Value
        ![AcknowledgedBy] = txtAcknowledgedBy.Text
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
        
        'Save to Purchase Order Details
        Dim RSSalesOrderReturnDetails As New Recordset
    
        RSSalesOrderReturnDetails.CursorLocation = adUseClient
        RSSalesOrderReturnDetails.Open "SELECT * From Sales_Order_Delivery_Detail Where DeliveryID = " & PK, CN, , adLockOptimistic
        
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            RSDetails.AddNew

            RSDetails![SalesOrderReturnID] = SalesOrderReturnPK
            RSDetails![StockID] = toNumber(.TextMatrix(c, 10))
            RSDetails![Qty] = toNumber(.TextMatrix(c, 4))
            RSDetails![Unit] = GetUnitID(.TextMatrix(c, 5))
            RSDetails![Price] = toNumber(.TextMatrix(c, 6))
            RSDetails![Discount] = toNumber(.TextMatrix(c, 8)) / 100

            RSDetails.Update

            'Add record to stock card
            RSStockCard.AddNew
                
            RSStockCard!Type = "SR"
            RSStockCard!RefNo2 = txtCreditMemoNo.Text
            RSStockCard!Pieces2 = "-" & toNumber(.TextMatrix(c, 4))
            RSStockCard!Cost = toNumber(.TextMatrix(c, 6))
            RSStockCard!StockID = toNumber(.TextMatrix(c, 10))
                
            RSStockCard.Update

            'Add qty received in stocks
            RSStocks.Find "[StockID] = " & toNumber(.TextMatrix(c, 10)), , adSearchForward, 1
            RSStocks!Onhand = toNumber(RSStocks!Onhand) + toNumber(.TextMatrix(c, 4))
            
            RSStocks.Update
            
            'add qty returned in Purchase Order Return Details
            RSSalesOrderReturnDetails.Find "[StockID] = " & toNumber(.TextMatrix(c, 10)), , adSearchForward, 1
            RSSalesOrderReturnDetails!QtyReturned = toNumber(RSSalesOrderReturnDetails!QtyReturned) + toNumber(.TextMatrix(c, 4))
            
            RSSalesOrderReturnDetails.Update
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
        txtCreditMemoDate.SetFocus
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
        rs.Open "SELECT * FROM qry_Sales_Order_Delivery WHERE DeliveryID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpCreditMemoDate.Value = Date
        dtpDate.Value = Date
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
                   
        DisplayForEditing
    Else
        'Set the recordset
        rs.Open "SELECT * FROM qry_Sales_Order_Return WHERE SalesOrderReturnID=" & PK, CN, adOpenStatic, adLockOptimistic
        
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
    If HaveAction = True Then
        frmSalesOrderDelivery.RefreshRecords
    End If
    
    Set frmSalesOrderReturnAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        txtStock.Text = .TextMatrix(.Rowsel, 2)
        txtICode = .TextMatrix(.Rowsel, 3)
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
    nsdClient.DisableDropdown = True
    nsdClient.TextReadOnly = True
    nsdClient.Text = rs!company
    txtSONo.Text = rs!SONo
    txtDeliveryNo.Text = rs!DeliveryNo
    txtDeliveryDate.Text = rs!DeliveryDate
    
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
    RSDetails.Open "SELECT * FROM qry_Sales_Order_Delivery_Detail WHERE DeliveryID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
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
  
    dtpCreditMemoDate.Visible = True
    txtCreditMemoNo.Visible = False
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
    nsdClient.DisableDropdown = True
    nsdClient.TextReadOnly = True
    nsdClient.Text = rs!company
    txtSONo.Text = rs!SONo
    
    txtDeliveryNo.Text = rs![DeliveryNo]
    txtDeliveryDate.Text = rs![DeliveryDate]
    
    txtCreditMemoNo.Text = rs![CreditMemoNo]
    txtCreditMemoNo.Text = rs![CreditMemoNo]
    txtReason.Text = rs![Reason]
    txtReference.Text = rs![Reference]
    txtPreparedBy.Text = rs![PreparedBy]
    txtReceivedBy.Text = rs![ReceivedBy]
    txtPostedBy.Text = rs![PostedBy]
    txtApprovedBy.Text = rs![ApprovedBy]
    txtDate.Text = rs![Date]
    txtAcknowledgedBy.Text = rs![AcknowledgedBy]
    txtRemarks.Text = rs![Remarks]
        
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
    RSDetails.Open "SELECT * FROM qry_Sales_Order_Return_Detail WHERE SalesOrderReturnID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
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
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![Gross])
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

    dtpCreditMemoDate.Visible = False
    txtCreditMemoNo.Visible = True
    dtpDate.Visible = False
    txtDate.Visible = True
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnUpdate.Visible = False

    'Resize and reposition the controls
    'Shape3.Top = 4800
    'Label11.Top = 4800
    'Line1(1).Visible = False
    'Line2(1).Visible = False
    Grid.Top = 3460
    Grid.Height = 3050
    
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





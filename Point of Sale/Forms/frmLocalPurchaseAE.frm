VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmLocalPurchaseAE 
   BorderStyle     =   0  'None
   Caption         =   "View Record"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintPreview 
      Caption         =   "Print Preview"
      Height          =   315
      Left            =   6780
      TabIndex        =   63
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton CmdReturn 
      Caption         =   "Return"
      Height          =   315
      Left            =   5340
      TabIndex        =   62
      Top             =   8490
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtCanvass 
      Height          =   285
      Left            =   1740
      TabIndex        =   5
      Top             =   2040
      Width           =   2475
   End
   Begin VB.ComboBox cmbPaymentType 
      Height          =   315
      ItemData        =   "frmLocalPurchaseAE.frx":0000
      Left            =   5940
      List            =   "frmLocalPurchaseAE.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   2475
   End
   Begin VB.TextBox txtCash 
      Height          =   285
      Left            =   5940
      TabIndex        =   7
      Top             =   1080
      Width           =   2475
   End
   Begin VB.TextBox txtPurchaseRequest 
      Height          =   285
      Left            =   1740
      TabIndex        =   4
      Top             =   1710
      Width           =   2475
   End
   Begin VB.TextBox txtPurchaseFrom 
      Height          =   285
      Left            =   1740
      TabIndex        =   0
      Top             =   720
      Width           =   2475
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9540
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7500
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9540
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1425
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1380
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.TextBox txtInvoiceNo 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1740
      TabIndex        =   1
      Top             =   1065
      Width           =   2475
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7950
      Width           =   1425
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   29
      Top             =   8490
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9615
      TabIndex        =   28
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   8190
      TabIndex        =   27
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmLocalPurchaseAE.frx":001A
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Remove"
      Top             =   4050
      Visible         =   0   'False
      Width           =   275
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
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.TextBox txtRemarks 
      Height          =   1335
      Left            =   225
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Tag             =   "Remarks"
      Top             =   6870
      Width           =   5910
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6900
      Width           =   1425
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   30
      Top             =   8340
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2610
      Left            =   210
      TabIndex        =   37
      Top             =   3900
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   4604
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
      Left            =   1740
      TabIndex        =   2
      Top             =   1380
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   94961667
      CurrentDate     =   38207
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   240
      ScaleHeight     =   630
      ScaleWidth      =   10740
      TabIndex        =   31
      Top             =   2910
      Width           =   10740
      Begin VB.TextBox txtDisc 
         Height          =   285
         Left            =   7200
         TabIndex        =   18
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9000
         TabIndex        =   20
         Top             =   225
         Width           =   840
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   2775
         TabIndex        =   14
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtUnitPrice 
         Height          =   285
         Left            =   4470
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   5775
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1290
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   225
         Width           =   2640
         _ExtentX        =   4657
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
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   7140
         TabIndex        =   49
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   8040
         TabIndex        =   46
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   45
         Top             =   0
         Width           =   900
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
         TabIndex        =   35
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   240
         Index           =   9
         Left            =   4500
         TabIndex        =   34
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Qty"
         Height          =   240
         Index           =   10
         Left            =   2775
         TabIndex        =   33
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   5775
         TabIndex        =   32
         Top             =   0
         Width           =   1260
      End
   End
   Begin VB.Frame frmPayment 
      Caption         =   "Bank Form"
      Height          =   1635
      Left            =   4440
      TabIndex        =   8
      Top             =   1080
      Width           =   6465
      Begin VB.TextBox txtBankDate 
         Height          =   315
         Left            =   2130
         TabIndex        =   10
         Top             =   510
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txtBank 
         Height          =   285
         Left            =   2130
         TabIndex        =   9
         Top             =   180
         Width           =   4035
      End
      Begin VB.TextBox txtCheckAmount 
         Height          =   285
         Left            =   2130
         TabIndex        =   12
         Top             =   1200
         Width           =   1905
      End
      Begin VB.TextBox txtCheck 
         Height          =   285
         Left            =   2130
         TabIndex        =   11
         Top             =   870
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpBankDate 
         Height          =   315
         Left            =   2130
         TabIndex        =   54
         Top             =   510
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   94961667
         CurrentDate     =   38932
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Bank:"
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   56
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   1
         Left            =   510
         TabIndex        =   55
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Amount:"
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   53
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Check No.:"
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   52
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Purchase"
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
      TabIndex        =   61
      Top             =   150
      Width           =   4905
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label lblCash 
      Alignment       =   1  'Right Justify
      Caption         =   "Cash:"
      Height          =   255
      Left            =   4410
      TabIndex        =   59
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment Type:"
      Height          =   255
      Left            =   4410
      TabIndex        =   60
      Top             =   750
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchase Req. No.:"
      Height          =   255
      Left            =   210
      TabIndex        =   58
      Top             =   1710
      Width           =   1485
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Canvass No.:"
      Height          =   255
      Left            =   210
      TabIndex        =   57
      Top             =   2070
      Width           =   1485
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5100
      TabIndex        =   51
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchase from:"
      Height          =   255
      Left            =   210
      TabIndex        =   50
      Top             =   750
      Width           =   1485
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Purchase Details"
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
      Left            =   300
      TabIndex        =   41
      Top             =   3600
      Width           =   4365
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   9210
      X2              =   10920
      Y1              =   7890
      Y2              =   7890
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
      TabIndex        =   48
      Top             =   7530
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
      TabIndex        =   47
      Top             =   7230
      Width           =   2040
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   225
      Top             =   3600
      Width           =   10740
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Invoice No.:"
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
      TabIndex        =   44
      Top             =   1065
      Width           =   1485
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   43
      Top             =   1380
      Width           =   1485
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
      TabIndex        =   42
      Top             =   7980
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   10935
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   10935
      Y1              =   2790
      Y2              =   2790
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
      TabIndex        =   40
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Label Labels 
      Caption         =   "Remarks"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   39
      Top             =   6600
      Width           =   990
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
      TabIndex        =   38
      Top             =   6930
      Width           =   2040
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
End
Attribute VB_Name = "frmLocalPurchaseAE"
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

Private Sub btnAdd_Click()
    If nsdStock.Text = "" Then nsdStock.SetFocus: Exit Sub

    If toNumber(txtUnitPrice.Text) <= 0 Then
        MsgBox "Please enter a valid sales price.", vbExclamation
        txtUnitPrice.SetFocus
        Exit Sub
    End If

    Dim CurrRow As Integer

    CurrRow = getFlexPos(Grid, 10, nsdStock.BoundText)

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                .TextMatrix(1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(1, 2) = nsdStock.Text
                .TextMatrix(1, 3) = nsdStock.getSelValueAt(5)
                .TextMatrix(1, 4) = txtQty.Text
                .TextMatrix(1, 5) = dcUnit.Text
                .TextMatrix(1, 6) = toMoney(txtUnitPrice.Text)
                .TextMatrix(1, 7) = toMoney(txtGross(1).Text)
                .TextMatrix(1, 8) = toNumber(txtDisc.Text)
                .TextMatrix(1, 9) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(1, 10) = nsdStock.BoundText
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(.Rows - 1, 2) = nsdStock.Text
                .TextMatrix(.Rows - 1, 3) = nsdStock.getSelValueAt(5)
                .TextMatrix(.Rows - 1, 4) = txtQty.Text
                .TextMatrix(.Rows - 1, 5) = dcUnit.Text
                .TextMatrix(.Rows - 1, 6) = toMoney(txtUnitPrice.Text)
                .TextMatrix(.Rows - 1, 7) = toMoney(txtGross(1).Text)
                .TextMatrix(.Rows - 1, 8) = toNumber(txtDisc.Text)
                .TextMatrix(.Rows - 1, 9) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(.Rows - 1, 10) = nsdStock.BoundText
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Invoice payment already exist. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                'Restore back the invoice amount and discount
                cIGross = cIGross - toNumber(Grid.TextMatrix(.Rowsel, 7))
                txtGross(2).Text = Format$(cIGross, "#,##0.00")
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.Rowsel, 9))
                txtNet.Text = Format$(cIAmount, "#,##0.00")
                cDAmount = cDAmount - toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(Grid.TextMatrix(.Rowsel, 4)) * toNumber(txtUnitPrice.Text)))
                txtDesc.Text = Format$(cDAmount, "#,##0.00")
                
                .TextMatrix(CurrRow, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(CurrRow, 2) = nsdStock.Text
                .TextMatrix(CurrRow, 3) = nsdStock.getSelValueAt(5)
                .TextMatrix(CurrRow, 4) = txtQty.Text
                .TextMatrix(CurrRow, 5) = dcUnit.Text
                .TextMatrix(CurrRow, 6) = toMoney(txtUnitPrice.Text)
                .TextMatrix(CurrRow, 7) = toMoney(txtGross(1).Text)
                .TextMatrix(CurrRow, 8) = toNumber(txtDisc.Text)
                .TextMatrix(CurrRow, 9) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(CurrRow, 10) = nsdStock.BoundText

            Else
                Exit Sub
            End If
        End If
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

Private Sub cmdPrintPreview_Click()
    With frmReports
        .strReport = "Local Purchase"
        .strWhere = "{Local_Purchase.LocalPurchaseID} =" & PK
        
        frmReports.show vbModal
    End With
End Sub

Private Sub CmdReturn_Click()
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Local_Purchase_Detail WHERE LocalPurchaseID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        With frmLocalPurchaseReturnAE
            .State = adStateAddMode
            .PK = PK
            .show vbModal
        End With
    Else
        MsgBox "All items are already returned.", vbInformation
    End If
End Sub

Private Sub txtCheckAmount_Validate(Cancel As Boolean)
    txtCheckAmount.Text = toMoney(toNumber(txtCheckAmount.Text))
End Sub

Private Sub txtdisc_Change()
    txtQty_Change
End Sub

Private Sub txtdisc_Click()
    txtQty_Change
End Sub

Private Sub cmbPaymentType_Click()
    If cmbPaymentType.ListIndex = 0 Then 'if Cash
        lblCash.Visible = True
        txtCash.Visible = True
        frmPayment.Visible = False
    Else 'if Bank
        frmPayment.Visible = True
        lblCash.Visible = False
        txtCash.Visible = False
    End If
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
    If txtPurchaseFrom.Text = "" Then
        MsgBox "Please enter vendor name/ description.", vbExclamation
        txtPurchaseFrom.SetFocus
        Exit Sub
    End If
    If txtInvoiceNo.Text = "" Then
        MsgBox "Please enter Invoice number", vbExclamation
        txtInvoiceNo.SetFocus
        Exit Sub
    End If
    If txtPurchaseRequest.Text = "" Then
        MsgBox "Please enter purchase request number.", vbExclamation
        txtPurchaseRequest.SetFocus
        Exit Sub
    End If
    If txtCanvass.Text = "" Then
        MsgBox "Please enter canvas number.", vbExclamation
        txtCanvass.SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to purchase before you can save this record.", vbExclamation
        nsdStock.SetFocus
        Exit Sub
    End If
    
    If isRecordExist("Local_Purchase", "InvoiceNo", txtInvoiceNo.Text, True) = True Then
        MsgBox "Invoice already exist. Please change it.", vbExclamation
        txtInvoiceNo.SetFocus
        Exit Sub
    End If
    
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Local_Purchase_Detail WHERE LocalPurchaseID=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    'Save the record
    With rs
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![LocalPurchaseID] = PK
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        Else
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        ![InvoiceNo] = txtInvoiceNo.Text
        ![Date] = dtpDate.Value
        ![PurchaseFrom] = txtPurchaseFrom.Text
        ![PaymentType] = cmbPaymentType.Text
        ![PurchaseRequestNo] = txtPurchaseRequest.Text
        ![CanvasSheetNo] = txtCanvass.Text
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)
        If cmbPaymentType.ListIndex = 0 Then ![Cash] = toMoney(toNumber(txtCash.Text)) 'if cash
        ![Remarks] = txtRemarks.Text
        
        .Update
    End With

    If cmbPaymentType.ListIndex = 1 Then 'if Bank
        Dim RSChecks As New Recordset
    
        RSChecks.CursorLocation = adUseClient
        RSChecks.Open "SELECT * FROM Local_Purchase_Checks WHERE LocalPurchaseID=" & PK, CN, adOpenStatic, adLockOptimistic
   
        With RSChecks
            .AddNew
            ![LocalPurchaseID] = PK
            ![Bank] = txtBank.Text
            ![Checkdate] = dtpBankDate.Value
            ![CheckNo] = txtCheck.Text
            ![CheckAmount] = txtCheckAmount.Text
            
            .Update
        End With
    End If
    
    With Grid
        'Save to stock card
        Dim RSStockCard As New Recordset
    
        RSStockCard.CursorLocation = adUseClient
        RSStockCard.Open "Stock_Card", CN, , adLockOptimistic, adCmdTable
    
        'Add qty ordered to qty onhand
        Dim RSStockUnit As New Recordset
    
        RSStockUnit.CursorLocation = adUseClient
        RSStockUnit.Open "SELECT * From Stock_Unit", CN, adOpenStatic, adLockOptimistic
    
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
            
                'Add qty received in Local Purchase Details
                RSDetails.AddNew

                RSDetails![LocalPurchaseID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 10))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 4))
                RSDetails![Unit] = GetUnitID(.TextMatrix(c, 5))
                RSDetails![Price] = toNumber(.TextMatrix(c, 6))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 8)) / 100

                RSDetails.Update
                
                '-----------------

                'Add qty in stock card
                RSStockCard.AddNew
                
                RSStockCard!Type = "CP"
                RSStockCard!RefNo1 = txtInvoiceNo.Text
                RSStockCard!Pieces1 = toNumber(.TextMatrix(c, 4))
                RSStockCard!Cost = toNumber(.TextMatrix(c, 6))
                RSStockCard!StockID = toNumber(.TextMatrix(c, 10))
                
                RSStockCard.Update
                
                '-----------------
                
                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 10)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 5) & "'", "UnitID")
    
                RSStockUnit!Onhand = RSStockUnit!Onhand + toNumber(.TextMatrix(c, 4))
                RSStockUnit.Update
            End If

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
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            GeneratePK
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in record has been successfully saved.", vbInformation
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
        txtPurchaseFrom.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid
    frmPayment.Visible = False
    
    bind_dc "SELECT * FROM Unit", "Unit", dcUnit, "UnitID", True
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        InitNSD
        
        
        cmdPrintPreview.Visible = False
        
        'Set the recordset
         rs.Open "SELECT * FROM Local_Purchase WHERE LocalPurchaseID=" & PK, CN, adOpenStatic, adLockOptimistic
         dtpDate.Value = Date
         dtpBankDate.Value = Date
         Caption = "Create New Entry"
         cmdUsrHistory.Enabled = False
         GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        cmdPrintPreview.Visible = True
        'Set the recordset
        rs.Open "SELECT * FROM Local_Purchase WHERE LocalPurchaseID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        cmdCancel.Caption = "Close"
        cmdUsrHistory.Enabled = True
               
        DisplayForViewing
        
        If ForCusAcc = True Then
            Me.Icon = frmLocalPurchase.Icon
        End If

        Screen.MousePointer = vbDefault
    End If
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Local_Purchase")
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
    nsdStock.ResetValue
    txtUnitPrice.Tag = 0
    txtUnitPrice.Text = "0.00"
    txtQty.Text = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmLocalPurchase.RefreshRecords
    End If
    
    Set frmLocalPurchaseAE = Nothing
End Sub

Private Sub Grid_Click()
    If State = adStateEditMode Then Exit Sub
    If Grid.Rows = 2 And Grid.TextMatrix(1, 10) = "" Then
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
        btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
        btnRemove.Left = Grid.Left + 50
    End If
End Sub

Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub


Private Sub nsdStock_Change()
On Error Resume Next

    txtQty.Text = "0"
    
    dcUnit.Text = ""
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & nsdStock.BoundText & " ORDER BY qry_Unit.Order ASC", "Unit", dcUnit, "UnitID", True
    
    nsdStock.Tag = nsdStock.BoundText
    
    'txtUnitPrice.Tag = nsdStock.getSelValueAt(3) 'Unit Cost
    txtUnitPrice.Text = toMoney(nsdStock.getSelValueAt(3)) 'Selling Price
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
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
        btnAdd.Enabled = False
    Else
        btnAdd.Enabled = True
    End If
    
    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text)))
    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text)) - ((toNumber(txtDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtUnitPrice.Text))))
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    InitGrid
    ResetEntry
    
    txtPurchaseFrom.Text = ""
    txtInvoiceNo.Text = ""
    dtpDate.Value = Date
    txtPurchaseRequest.Text = ""
    txtCanvass.Text = ""
    txtCash.Text = ""
    txtBank.Text = ""
    txtCheck.Text = ""
    txtCheckAmount.Text = ""
    txtRemarks.Text = ""
    
    txtGross(2).Text = "0.00"
    txtDesc.Text = "0.00"
    txtTaxBase.Text = "0.00"
    txtVat.Text = "0.00"
    txtNet.Text = "0.00"

    cIAmount = 0
    cDAmount = 0

    txtPurchaseFrom.SetFocus
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo err
    txtPurchaseFrom.Text = rs![PurchaseFrom]
    txtInvoiceNo.Text = rs![InvoiceNo]
    txtDate.Text = Format$(rs![Date], "MMM-dd-yyyy")
    txtPurchaseRequest.Text = rs![PurchaseRequestNo]
    txtCanvass.Text = rs![CanvasSheetNo]
    
    'Display payment type
    If rs![PaymentType] = "Cash" Then
        cmbPaymentType.ListIndex = 0
        txtCash.Text = rs!Cash
    ElseIf rs![PaymentType] = "Bank" Then
        cmbPaymentType.ListIndex = 1
        Dim RSChecks As New Recordset
    
        RSChecks.CursorLocation = adUseClient
        RSChecks.Open "SELECT * FROM Local_Purchase_Checks WHERE LocalPurchaseID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        With RSChecks
            txtBank.Text = ![Bank]
            dtpBankDate.Visible = False
            txtBankDate.Visible = True
            txtBankDate.Text = ![Checkdate]
            txtCheck.Text = ![CheckNo]
            txtCheckAmount.Text = ![CheckAmount]
        End With
    End If
    
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtTaxBase.Text = toMoney(rs![TaxBase])
    txtVat.Text = toMoney(rs![Vat])
    txtNet.Text = toMoney(rs![NetAmount])
    txtRemarks.Text = rs![Remarks]
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Local_Purchase_Detail WHERE LocalPurchaseID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
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
            Grid.FixedRows = Grid.Row: Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    'Disable commands
    LockInput Me, True

    dtpDate.Visible = False
    txtDate.Visible = True
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnAdd.Visible = False
    
    CmdReturn.Left = cmdSave.Left
    CmdReturn.Visible = True

    'Resize and reposition the controls
    Shape3.Top = 2800
    Label11.Top = 2800
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 3100
    Grid.Height = 3380

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub InitNSD()
    'For Stock
    With nsdStock
        .ClearColumn
        .AddColumn "Barcode", 2064.882
        .AddColumn "Stock", 4085.26
        .AddColumn "Supplier Price", 1500
        .AddColumn "ICode", 1500
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Barcode,Stock,SupplierPrice,ICode,StockID"
        .sqlTables = "qry_Stock_Unit_Purchase"
        .sqlSortOrder = "Stock ASC"
        .BoundField = "StockID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Stocks"
    End With
End Sub

Private Sub txtUnitPrice_GotFocus()
    HLText txtUnitPrice
End Sub

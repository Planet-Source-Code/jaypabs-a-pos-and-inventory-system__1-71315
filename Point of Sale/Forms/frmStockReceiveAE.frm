VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmStockReceiveAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Receive"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvInvoice 
      Height          =   1365
      Left            =   180
      TabIndex        =   54
      Top             =   2100
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   2408
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtSalesman 
      Height          =   285
      Left            =   1380
      TabIndex        =   34
      Top             =   1530
      Width           =   2475
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9465
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7230
      Width           =   1425
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   150
      TabIndex        =   31
      Top             =   7710
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9540
      TabIndex        =   30
      Top             =   7710
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   8085
      TabIndex        =   29
      Top             =   7710
      Width           =   1335
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   210
      Picture         =   "frmStockReceiveAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Remove"
      Top             =   4740
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
      Left            =   9465
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6480
      Width           =   1425
   End
   Begin VB.TextBox txtEntry 
      Height          =   750
      Index           =   8
      Left            =   150
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Tag             =   "Remarks"
      Top             =   6750
      Width           =   5805
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9465
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6780
      Width           =   1425
   End
   Begin VB.CommandButton cmdPH 
      Caption         =   "Payment History"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2025
      TabIndex        =   24
      Top             =   7710
      Width           =   1590
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   105
      ScaleHeight     =   630
      ScaleWidth      =   10740
      TabIndex        =   6
      Top             =   3600
      Width           =   10740
      Begin VB.ComboBox cbDisc 
         Height          =   315
         Left            =   7140
         TabIndex        =   13
         Text            =   "0"
         Top             =   225
         Width           =   765
      End
      Begin VB.TextBox txtNetAmount 
         Height          =   285
         Left            =   8040
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNetPrice 
         Height          =   285
         Left            =   9000
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9900
         TabIndex        =   10
         Top             =   225
         Width           =   840
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   3975
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtSP 
         Height          =   285
         Left            =   2670
         TabIndex        =   8
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
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1290
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   0
         TabIndex        =   14
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
         Left            =   4680
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
         TabIndex        =   23
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   8040
         TabIndex        =   22
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Price"
         Height          =   240
         Index           =   3
         Left            =   9000
         TabIndex        =   21
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   4680
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Price(Each)"
         Height          =   240
         Index           =   9
         Left            =   2700
         TabIndex        =   18
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Qty"
         Height          =   240
         Index           =   10
         Left            =   3975
         TabIndex        =   17
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   5775
         TabIndex        =   16
         Top             =   0
         Width           =   1260
      End
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1380
      TabIndex        =   5
      Top             =   870
      Width           =   3315
   End
   Begin VB.TextBox txtShipping_Instructions 
      Height          =   465
      Left            =   6930
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox txtAdditional_Instructions 
      Height          =   465
      Left            =   6930
      TabIndex        =   3
      Top             =   630
      Width           =   3855
   End
   Begin VB.TextBox txtDeclared_as 
      Height          =   315
      Left            =   6930
      TabIndex        =   2
      Top             =   1140
      Width           =   3855
   End
   Begin VB.TextBox txtDeclared_Value 
      Height          =   315
      Left            =   6930
      TabIndex        =   1
      Top             =   1500
      Width           =   3855
   End
   Begin VB.TextBox txtPONo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   3315
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -150
      TabIndex        =   35
      Top             =   7560
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   1830
      Left            =   120
      TabIndex        =   36
      Top             =   4590
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   3228
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
      Left            =   1380
      TabIndex        =   37
      Top             =   1200
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44564483
      CurrentDate     =   38207
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdVendor 
      Height          =   315
      Left            =   1380
      TabIndex        =   38
      Top             =   510
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
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   4995
      TabIndex        =   53
      Top             =   4980
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Salesman"
      Height          =   225
      Left            =   60
      TabIndex        =   52
      Top             =   1530
      Width           =   1275
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   9135
      X2              =   10845
      Y1              =   7170
      Y2              =   7170
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   4290
      Width           =   10740
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   " Date"
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   51
      Top             =   1185
      Width           =   1275
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
      Left            =   7365
      TabIndex        =   50
      Top             =   7260
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   135
      X2              =   10860
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   135
      X2              =   10860
      Y1              =   1980
      Y2              =   1980
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
      Left            =   7365
      TabIndex        =   49
      Top             =   6510
      Width           =   2040
   End
   Begin VB.Label Labels 
      Caption         =   "Remarks"
      Height          =   240
      Index           =   4
      Left            =   165
      TabIndex        =   48
      Top             =   6480
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
      Left            =   7365
      TabIndex        =   47
      Top             =   6810
      Width           =   2040
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
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   46
      Top             =   510
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   225
      Left            =   60
      TabIndex        =   45
      Top             =   855
      Width           =   1275
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Details"
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
      Left            =   195
      TabIndex        =   44
      Top             =   4290
      Width           =   4365
   End
   Begin VB.Label Label6 
      Caption         =   "Shipping Instructions"
      Height          =   255
      Left            =   5130
      TabIndex        =   43
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label Label12 
      Caption         =   "Additional Instructions"
      Height          =   255
      Left            =   5130
      TabIndex        =   42
      Top             =   630
      Width           =   1785
   End
   Begin VB.Label Label13 
      Caption         =   "Declared As"
      Height          =   255
      Left            =   5130
      TabIndex        =   41
      Top             =   1170
      Width           =   1785
   End
   Begin VB.Label Label14 
      Caption         =   "Declared Value"
      Height          =   255
      Left            =   5130
      TabIndex        =   40
      Top             =   1530
      Width           =   1785
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "PO No."
      Height          =   225
      Left            =   60
      TabIndex        =   39
      Top             =   150
      Width           =   1275
   End
End
Attribute VB_Name = "frmStockReceiveAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public LLFK                 As Long 'Last Loading FK
Public LLVFK                As Long 'Last Loading Van FK
Public LLDate               As String
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim PCase                   As Long 'Pieces per case
Dim PBox                    As Long 'Pieces per box

Dim old_pieces              As Long 'Old pieces value
Dim old_boxes               As Long 'Old boxes value
Dim old_cases               As Long 'Old cases value

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice

Private Sub btnAdd_Click()
    If nsdStock.Text = "" Then nsdStock.SetFocus: Exit Sub

    If toNumber(txtSP.Text) <= 0 Then
        MsgBox "Please enter a valid sales price.", vbExclamation
        txtSP.SetFocus
        Exit Sub
    End If

    Dim CurrRow As Integer

    CurrRow = getFlexPos(Grid, 11, nsdStock.BoundText)

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                .TextMatrix(1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(1, 2) = nsdStock.Text
                .TextMatrix(1, 3) = nsdStock.getSelValueAt(5)
                .TextMatrix(1, 4) = txtQty.Text
                .TextMatrix(1, 5) = dcUnit.Text
                .TextMatrix(1, 6) = txtSP.Text
                .TextMatrix(1, 7) = toMoney(txtGross(1).Text)
                .TextMatrix(1, 8) = toNumber(cbDisc.Text)
                .TextMatrix(1, 9) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(1, 10) = toMoney(toNumber(txtNetPrice.Text))
                .TextMatrix(1, 11) = nsdStock.BoundText
                .TextMatrix(1, 12) = toMoney(toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtSP.Text))
            Else
ADD_NEW_HERE:
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(.Rows - 1, 2) = nsdStock.Text
                .TextMatrix(.Rows - 1, 3) = nsdStock.getSelValueAt(5)
                .TextMatrix(.Rows - 1, 4) = txtQty.Text
                .TextMatrix(.Rows - 1, 5) = dcUnit.Text
                .TextMatrix(.Rows - 1, 6) = txtSP.Text
                .TextMatrix(.Rows - 1, 7) = toMoney(txtGross(1).Text)
                .TextMatrix(.Rows - 1, 8) = toNumber(cbDisc.Text)
                .TextMatrix(.Rows - 1, 9) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(.Rows - 1, 10) = toMoney(toNumber(txtNetPrice.Text))
                .TextMatrix(.Rows - 1, 11) = nsdStock.BoundText
                .TextMatrix(.Rows - 1, 12) = toMoney(toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtSP.Text))
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            'If free option is not equal or discount is not equal or sales price is not equal then add new sold item
            'If .TextMatrix(CurrRow, 10) <> changeYNValue(ckFree.Value) Or toNumber(.TextMatrix(CurrRow, 8)) <> toNumber(cbDisc.Text) Or toNumber(.TextMatrix(CurrRow, 3)) <> toNumber(txtSP.Text) Then
            '    GoTo ADD_NEW_HERE
            'End If
            
            If MsgBox("Invoice payment already exist.Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                'Restore back the invoice amount and discount
                cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 7))
                txtGross(2).Text = Format$(cIGross, "#,##0.00")
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 9))
                txtNet.Text = Format$(cIAmount, "#,##0.00")
                cDAmount = cDAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
                txtDesc.Text = Format$(cDAmount, "#,##0.00")
                
                .TextMatrix(CurrRow, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(CurrRow, 2) = nsdStock.Text
                .TextMatrix(CurrRow, 3) = nsdStock.getSelValueAt(5)
                .TextMatrix(CurrRow, 4) = txtQty.Text
                .TextMatrix(CurrRow, 5) = dcUnit.Text
                .TextMatrix(CurrRow, 6) = txtSP.Text
                .TextMatrix(CurrRow, 7) = toMoney(txtGross(1).Text)
                .TextMatrix(CurrRow, 8) = toNumber(cbDisc.Text)
                .TextMatrix(CurrRow, 9) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(CurrRow, 10) = toMoney(toNumber(txtNetPrice.Text))
                .TextMatrix(CurrRow, 11) = nsdStock.BoundText
                .TextMatrix(CurrRow, 12) = toMoney(toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtSP.Text))

            Else
                Exit Sub
            End If
        End If
        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtGross(1).Text)
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount + toNumber(txtNetAmount.Text)
        cDAmount = cDAmount + toNumber(toNumber(cbDisc.Text) / 100) * (toNumber(toNumber(txtQty.Text) * toNumber(txtSP.Text)))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Highlight the current row's column
        .ColSel = 11
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
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 7))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 9))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        cDAmount = cDAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
    
End Sub

Private Sub cbDisc_Change()
    txtQty_Change
End Sub

Private Sub cbDisc_Click()
    txtQty_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cbDisc_Validate(Cancel As Boolean)
    cbDisc.Text = toNumber(cbDisc.Text)
End Sub

Private Sub cmdPH_Click()
    'frmInvoiceViewerPH.INV_PK = PK
    'frmInvoiceViewerPH.Caption = "Payment History Viewer"
    'frmInvoiceViewerPH.lblTitle.Caption = "Payment History Viewer"
    'frmInvoiceViewerPH.show vbModal
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If nsdVendor.BoundText = "" Then
        MsgBox "Please select a vendor.", vbExclamation
        nsdVendor.SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to purchase before saving this record.", vbExclamation
        nsdStock.SetFocus
        Exit Sub
    End If
       
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM purchase_order_detail WHERE purchase_order_id=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    'Save the record
    With rs
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![purchase_order_id] = PK
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        Else
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        ![vendor_id] = nsdVendor.BoundText
        ![po_no] = txtPONo.Text
        ![Date] = dtpDate.Value
        ![salesman] = txtSalesman.Text
        ![shipping_instructions] = txtShipping_Instructions.Text
        ![additional_instructions] = txtAdditional_Instructions.Text
        ![declared_as] = txtDeclared_as.Text
        ![declared_value] = txtDeclared_Value.Text
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![amount_net] = toNumber(txtNet.Text)
        ![Remarks] = txtEntry(8).Text
        
        .Update
        
    End With
  
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
            
                RSDetails.AddNew

                'RSDetails![PK] = getIndex("tbl_AR_InvoiceDetails")

                RSDetails![purchase_order_id] = PK
                RSDetails![stock_id] = toNumber(.TextMatrix(c, 11))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 4))
                RSDetails![Unit] = toNumber(.TextMatrix(c, 5))
                RSDetails![Price] = toNumber(.TextMatrix(c, 6))
                RSDetails![amount_gross] = toNumber(.TextMatrix(c, 7))
                RSDetails![discount_percent] = toNumber(.TextMatrix(c, 8))
                RSDetails![discount_amount] = toNumber(.TextMatrix(c, 12))
                RSDetails![amount_net] = toNumber(.TextMatrix(c, 9))
                RSDetails![net_price] = toNumber(.TextMatrix(c, 10))
                RSDetails![Date] = dtpDate.Value

                RSDetails.Update

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
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
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
        txtEntry(0).SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid
    
    bind_dc "SELECT * FROM Unit", "Unit", dcUnit, "unit_id", True
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        InitNSD
        
        'Set the recordset
         rs.Open "SELECT * FROM purchase_order WHERE purchase_order_id=" & PK, CN, adOpenStatic, adLockOptimistic
         dtpDate.Value = Date
         Caption = "Create New Entry"
         cmdUsrHistory.Enabled = False
         GeneratePK
         txtPONo.Text = Format(PK, "0000000000")
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_PurchaseOrder WHERE purchase_order_id=" & PK, CN, adOpenStatic, adLockOptimistic
        
        cmdCancel.Caption = "Close"
        cmdUsrHistory.Enabled = True
               
        DisplayForViewing
        
        If ForCusAcc = True Then
            Me.Icon = frmCashPurchase.Icon
        Else
            
            MsgBox "This is use for viewing the record only." & vbCrLf & _
               "You cannot perform any changes in this form." & vbCrLf & vbCrLf & _
               "Note:If you have mistake in adding this record then " & vbCrLf & _
               "void this record and re-enter.", vbExclamation
        End If

        Screen.MousePointer = vbDefault
    End If
    
    'Initialize Graphics
    'With MAIN
        'cmdGenerate.Picture = .i16x16.ListImages(14).Picture
        'cmdNew.Picture = .i16x16.ListImages(10).Picture
        'cmdReset.Picture = .i16x16.ListImages(15).Picture
    'End With
 
    'Fill the discount combo
    cbDisc.AddItem "0.01"
    cbDisc.AddItem "0.02"
    cbDisc.AddItem "0.03"
    cbDisc.AddItem "0.04"
    cbDisc.AddItem "0.05"
    cbDisc.AddItem "0.06"
    cbDisc.AddItem "0.07"
    cbDisc.AddItem "0.08"
    cbDisc.AddItem "0.09"
    cbDisc.AddItem "0.1"
     
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("purchase_order")
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
        .Cols = 13
        .ColSel = 11
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
        .ColWidth(10) = 750
        .ColWidth(11) = 500
        .ColWidth(12) = 500
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
        .TextMatrix(0, 10) = "Net Price"
        .TextMatrix(0, 11) = "Stock ID"
        .TextMatrix(0, 12) = "Disc"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
        .ColAlignment(6) = vbLeftJustify
        .ColAlignment(7) = vbLeftJustify
        .ColAlignment(8) = vbLeftJustify
        .ColAlignment(9) = vbLeftJustify
        .ColAlignment(10) = vbLeftJustify
        .ColAlignment(11) = vbLeftJustify
        .ColAlignment(12) = vbLeftJustify
    End With
End Sub

Private Sub ResetEntry()
    nsdStock.ResetValue
    txtSP.Tag = 0
    txtSP.Text = "0.00"
    txtQty.Text = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmPurchaseOrder.RefreshRecords
    End If
    
    Set frmPurchaseOrderAE = Nothing
End Sub

Private Sub Grid_Click()
    If State = adStateEditMode Then Exit Sub
    If Grid.Rows = 2 And Grid.TextMatrix(1, 11) = "" Then
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
    txtQty.Text = "0"
    
    txtSP.Tag = nsdStock.getSelValueAt(3) 'Unit Cost
    txtSP.Text = nsdStock.getSelValueAt(4) 'Selling Price
End Sub

Private Sub nsdVendor_Change()
    If nsdVendor.DisableDropdown = False Then
        txtAddress.Text = nsdVendor.getSelValueAt(3)
    End If
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index > 1 And Index < 5 Then
        txtQty.Text = (toNumber(txtEntry(2).Text) * PCase) + (toNumber(txtEntry(3).Text) * PBox) + toNumber(txtEntry(4).Text)
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
    If Index = 8 Then
        cmdSave.Default = False
    End If
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 1 And Index < 8 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then
        cmdSave.Default = True
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index > 1 And Index < 8 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub

Private Sub txtSP_Change()
    txtQty_Change
End Sub

Private Sub txtSP_Validate(Cancel As Boolean)
    txtSP.Text = toMoney(toNumber(txtSP.Text))
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnAdd.Enabled = False
    Else
        btnAdd.Enabled = True
    End If
    
    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * toNumber(txtSP.Text)))
    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(txtSP.Text)) - ((toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtSP.Text))))
    If toNumber(txtQty.Text) < 1 Then txtNetPrice.Text = 0: Exit Sub
    txtNetPrice.Text = toMoney(toNumber(txtSP.Text)) - ((toNumber(txtSP.Text) * (toNumber(cbDisc.Text) / 100)))
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
End Sub

Private Sub txtSP_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    InitGrid
    ResetEntry
    
    txtEntry(0).Text = ""
    dtpDate.Value = Date
    txtpurchase_from.Text = ""
        
    txtEntry(8).Text = ""
    
    txtGross(2).Text = "0.00"
    txtDesc.Text = "0.00"
    txtNet.Text = "0.00"

    cIAmount = 0
    cDAmount = 0

    txtEntry(0).SetFocus
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo err
    nsdVendor.DisableDropdown = True
    nsdVendor.TextReadOnly = True
    nsdVendor.Text = rs!company
    txtPONo.Text = rs!po_no
    txtAddress.Text = rs!address
    txtDate.Text = rs![Date]
    txtSalesman.Text = rs![salesman]
    txtShipping_Instructions.Text = rs![shipping_instructions]
    txtAdditional_Instructions.Text = rs![additional_instructions]
    txtDeclared_as.Text = rs![declared_as]
    txtDeclared_Value.Text = rs![declared_value]
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtNet.Text = toMoney(rs![amount_net])
    txtEntry(8).Text = rs![Remarks]
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_purchase_orderDetails WHERE purchase_order_id=" & PK & " ORDER BY stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                    .TextMatrix(1, 1) = RSDetails![barcode]
                    .TextMatrix(1, 2) = RSDetails![stock]
                    .TextMatrix(1, 3) = RSDetails![icode]
                    .TextMatrix(1, 4) = RSDetails![Qty]
                    .TextMatrix(1, 5) = RSDetails![Unit]
                    .TextMatrix(1, 6) = RSDetails![Price]
                    .TextMatrix(1, 7) = RSDetails![amount_gross]
                    .TextMatrix(1, 8) = RSDetails![discount_percent]
                    .TextMatrix(1, 9) = toMoney(RSDetails![amount_net])
                    .TextMatrix(1, 10) = RSDetails![net_price]
                    .TextMatrix(1, 11) = RSDetails![stock_id]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![icode]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Price]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![amount_gross]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![discount_percent]
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![amount_net])
                    .TextMatrix(.Rows - 1, 10) = RSDetails![net_price]
                    .TextMatrix(.Rows - 1, 11) = RSDetails![stock_id]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 11
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
    'txtLess.Locked = True

    'Resize and reposition the controls
    Shape3.Top = 2100 '2850
    Label11.Top = 2100 '2850
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 2400 '3150
    Grid.Height = 3180 '2800

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
    'For Vendor
    With nsdVendor
        .ClearColumn
        .AddColumn "Vendor ID", 1794.89
        .AddColumn "Company", 2264.88
        .AddColumn "Address", 2670.23
        .Connection = CN.ConnectionString
        
        .sqlFields = "vendor_id, company, address"
        .sqlTables = "vendors"
        .sqlSortOrder = "company ASC"
        
        .BoundField = "vendor_id"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Vendors Record"
        
    End With

    'For Stock
    With nsdStock
        .ClearColumn
        .AddColumn "Barcode", 2064.882
        .AddColumn "Stock", 4085.26
        .AddColumn "Cost", 1500
        .AddColumn "Sales Price", 1500
        .AddColumn "ICode", 1500
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "barcode,stock,cost,sales_price,icode,stock_id"
        .sqlTables = "stocks"
        .sqlSortOrder = "stock ASC"
        .BoundField = "stock_id"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Stocks"
        
    End With

End Sub

Private Sub txtSP_GotFocus()
    HLText txtSP
End Sub

Private Sub InitInvoice()
    
End Sub

VERSION 5.00
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmPrintBarcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Barcode(s)"
   ClientHeight    =   1365
   ClientLeft      =   4290
   ClientTop       =   5520
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   315
      Left            =   2730
      TabIndex        =   2
      Top             =   780
      Width           =   1125
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdStock 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   360
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
   Begin VB.Label Label1 
      Caption         =   "Barcode"
      Height          =   165
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmPrintBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    With frmReports
        .strReport = "Print Barcode"
        .strWhere = "{Barcode_Temp.StockID} = " & nsdStock.BoundText
        
        frmReports.show vbModal
    End With
End Sub

Private Sub Form_Load()
  InitNSD
End Sub

Private Sub InitNSD()
  With nsdStock
    .ClearColumn
    .AddColumn "Barcode", 2064.882
    .AddColumn "Stock", 4085.26
    .AddColumn "Short1", 1500
    .AddColumn "Short2", 1500
    .AddColumn "ICode", 1500
    
    .Connection = CN.ConnectionString
    
    .sqlFields = "Barcode,Stock,Short1,Short2,ICode,StockID"
    .sqlTables = "qry_Barcodes1"
    .sqlSortOrder = "Stock ASC"
    .BoundField = "StockID"
    .PageBy = 25
    .DisplayCol = 2
    
    .setDropWindowSize 6800, 4000
    .TextReadOnly = True
    .SetDropDownTitle = "Stocks"
  End With
End Sub

Private Sub nsdStock_Change()
  nsdStock.Text = nsdStock.getSelValueAt(1)
End Sub

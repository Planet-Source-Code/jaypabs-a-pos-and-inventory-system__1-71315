VERSION 5.00
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmStockCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Card"
   ClientHeight    =   1545
   ClientLeft      =   2250
   ClientTop       =   4260
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   930
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   315
      Left            =   1410
      TabIndex        =   1
      Top             =   930
      Width           =   1095
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdStock 
      Height          =   315
      Left            =   330
      TabIndex        =   3
      Top             =   510
      Width           =   3330
      _ExtentX        =   5874
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
      Caption         =   "Product:"
      Height          =   285
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   1125
   End
End
Attribute VB_Name = "frmStockCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
    With frmReports
        .strReport = "Stock Card"
        
        If nsdStock.Text <> "" Then
            .strWhere = "{qry_Stock_Card.StockIDAlias} = " & nsdStock.BoundText
        Else
            .strWhere = ""
        End If
        
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
    .AddColumn "Sales Price", 1500
    .AddColumn "Cost", 1500
    .AddColumn "ICode", 1500
    
    .Connection = CN.ConnectionString
    
    .sqlFields = "Barcode,Stock,SalesPrice,Cost,ICode,StockID"
    .sqlTables = "Stocks"
    .sqlSortOrder = "Stock ASC"
    .BoundField = "StockID"
    .PageBy = 25
    .DisplayCol = 2
    
    .setDropWindowSize 6800, 4000
    .TextReadOnly = True
    .SetDropDownTitle = "Stocks"
  End With
End Sub


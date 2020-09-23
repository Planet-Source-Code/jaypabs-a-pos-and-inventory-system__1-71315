VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBegInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beginning Inventory"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo dcUnit 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   1470
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2130
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2340
      TabIndex        =   6
      Top             =   2130
      Width           =   1035
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
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1080
      Width           =   960
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   690
      Width           =   945
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   60
      TabIndex        =   8
      Top             =   2010
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   53
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Sales Price"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1110
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit"
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   1470
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Qty"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   690
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Beginning Inventory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4365
   End
End
Attribute VB_Name = "frmBegInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intStockID       As Integer

Dim rs As New Recordset

Private Sub cmdCancel_Click()
    rs.Close
    Unload Me
End Sub

Private Sub cmdSave_Click()
    With rs
        If .RecordCount > 0 Then
            rs!Pieces1 = txtQty.Text
            rs!Cost = txtSalesPrice
            rs!StockID = nsdUnit.Text
            
            .Update
        Else
            .AddNew
            
            rs!StockID = intStockID
            rs!Type = "BI"
            rs!Pieces1 = txtQty.Text
            rs!Cost = txtSalesPrice
            rs!UnitID = dcUnit.BoundText
            
            .Update
        End If
    End With
    
    rs.Close
    Unload Me
End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM Unit", "Unit", dcUnit, "UnitID", True
    dcUnit.Text = ""
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT StockID, UnitID, DateInsert, Type, Pieces1, Cost FROM Stock_Card WHERE StockID=" & intStockID & " AND [Type] = 'BI' AND format(DateInsert,'yyyy') =" & Year(Date), CN, adOpenStatic, adLockOptimistic
    Debug.Print "SELECT StockID, UnitID, DateInsert, Type, Pieces1, Cost FROM Stock_Card WHERE StockID=" & intStockID & " AND [Type] = 'BI' AND format(DateInsert,'yyyy') =" & Year(Date)

    If rs.RecordCount > 0 Then
        txtQty.Text = rs!Pieces1
        txtSalesPrice = rs!Cost
        dcUnit.BoundText = rs!UnitID
    End If
End Sub

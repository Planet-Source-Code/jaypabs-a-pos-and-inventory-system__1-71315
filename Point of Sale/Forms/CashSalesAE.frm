VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CashSalesAE 
   ClientHeight    =   6435
   ClientLeft      =   2565
   ClientTop       =   2325
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8730
   Begin MSComctlLib.ListView lvList 
      Height          =   2055
      Left            =   270
      TabIndex        =   17
      Top             =   3060
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   3625
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Seq"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Disc"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   7
      Left            =   1020
      TabIndex        =   15
      Top             =   2700
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   180
      TabIndex        =   14
      Top             =   2550
      Width           =   8325
   End
   Begin VB.TextBox txtEntry 
      Height          =   765
      Index           =   6
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1740
      Width           =   3525
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   5
      Left            =   1020
      TabIndex        =   11
      Text            =   "ABS-CBN (Sabado Barkada)"
      Top             =   1380
      Width           =   3525
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   4
      Left            =   1020
      TabIndex        =   9
      Text            =   "00001"
      Top             =   1020
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   3
      Left            =   3210
      TabIndex        =   7
      Text            =   "Zenny"
      Top             =   630
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   2
      Left            =   1020
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   660
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   1
      Left            =   3210
      TabIndex        =   3
      Text            =   "4:09:12"
      Top             =   270
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   1
      Text            =   "Dec-9-06"
      Top             =   300
      Width           =   1335
   End
   Begin VB.Label Lables 
      Caption         =   "Barcode"
      Height          =   255
      Index           =   7
      Left            =   270
      TabIndex        =   16
      Top             =   2730
      Width           =   735
   End
   Begin VB.Label Lables 
      Caption         =   "Remarks"
      Height          =   255
      Index           =   6
      Left            =   270
      TabIndex        =   13
      Top             =   1770
      Width           =   735
   End
   Begin VB.Label Lables 
      Caption         =   "Customer"
      Height          =   255
      Index           =   5
      Left            =   270
      TabIndex        =   10
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label Lables 
      Caption         =   "CSI#"
      Height          =   255
      Index           =   4
      Left            =   270
      TabIndex        =   8
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Lables 
      Caption         =   "Cashier"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Lables 
      Caption         =   "Terminal"
      Height          =   255
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Lables 
      Caption         =   "Time"
      Height          =   255
      Index           =   1
      Left            =   2670
      TabIndex        =   2
      Top             =   300
      Width           =   735
   End
   Begin VB.Label Lables 
      Caption         =   "Date"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   330
      Width           =   735
   End
End
Attribute VB_Name = "CashSalesAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


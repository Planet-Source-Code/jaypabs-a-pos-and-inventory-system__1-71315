VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCashSalesAE 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Edit Entry"
   ClientHeight    =   8175
   ClientLeft      =   2580
   ClientTop       =   1365
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9075
      Left            =   -1080
      Picture         =   "frmCashSalesAE.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   12795
      TabIndex        =   13
      Top             =   0
      Width           =   12795
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   9660
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   2100
         Width           =   825
      End
      Begin VB.TextBox txtEntry 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   10
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   5940
         Width           =   4425
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   405
         Index           =   9
         Left            =   10890
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   7020
         Width           =   1515
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2100
         Width           =   1005
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   8790
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   2100
         Width           =   825
      End
      Begin VB.TextBox txtEntry 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2100
         Width           =   4905
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   8130
         TabIndex        =   4
         Text            =   "1"
         Top             =   2100
         Width           =   615
      End
      Begin VB.TextBox txtEntry 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   10530
         TabIndex        =   7
         Text            =   "0"
         Top             =   2100
         Width           =   825
      End
      Begin VB.TextBox txtEntry 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Top             =   2100
         Width           =   1575
      End
      Begin VB.TextBox txtEntry 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2430
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1425
         Width           =   3135
      End
      Begin VB.TextBox txtEntry 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2430
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1635
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   3240
         Left            =   1500
         TabIndex        =   12
         Top             =   2400
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   5715
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   275
         ForeColorFixed  =   -2147483640
         BackColorSel    =   1091552
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin InvtySystem.vButtons_H cmdPrint 
         Height          =   345
         Left            =   3030
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   6870
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         Caption         =   "(F12) Re-Print"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   4194304
         cGradient       =   4194304
         Gradient        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin InvtySystem.vButtons_H cmdSettle 
         Height          =   345
         Left            =   1590
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6870
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         Caption         =   "(F11) Settle"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   4194304
         cGradient       =   4194304
         Gradient        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin InvtySystem.vButtons_H btnLabel 
         Height          =   345
         Left            =   1620
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   7410
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   609
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388608
         cFHover         =   8388608
         cBhover         =   4194304
         cGradient       =   4194304
         Gradient        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   8
         Left            =   9810
         TabIndex        =   49
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   10710
         TabIndex        =   48
         Top             =   5910
         Width           =   1635
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   8430
         TabIndex        =   47
         Top             =   5910
         Width           =   2040
      End
      Begin VB.Label lblTaxbase 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   10710
         TabIndex        =   46
         Top             =   6210
         Width           =   1635
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax base"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   8430
         TabIndex        =   45
         Top             =   6210
         Width           =   2040
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   7
         Left            =   1560
         TabIndex        =   44
         Top             =   5700
         Width           =   1155
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Due"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   8430
         TabIndex        =   43
         Top             =   6780
         Width           =   2040
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   10710
         TabIndex        =   42
         Top             =   6780
         Width           =   1635
      End
      Begin VB.Label lblVAT12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   10710
         TabIndex        =   41
         Top             =   6510
         Width           =   1635
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VAT (12%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   8430
         TabIndex        =   40
         Top             =   6510
         Width           =   2040
      End
      Begin VB.Label lblChange 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   10740
         TabIndex        =   27
         Top             =   7320
         Width           =   1635
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "(F5) Accept Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   6600
         TabIndex        =   39
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "(F6) Add To List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   5130
         TabIndex        =   37
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(F9) Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   1560
         TabIndex        =   36
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   6
         Left            =   11400
         TabIndex        =   35
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   1
         Left            =   8940
         TabIndex        =   34
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   5
         Left            =   3150
         TabIndex        =   33
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   4
         Left            =   8070
         TabIndex        =   32
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disc in %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   3
         Left            =   10530
         TabIndex        =   31
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   2
         Left            =   1560
         TabIndex        =   30
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   8430
         TabIndex        =   29
         Top             =   5640
         Width           =   2040
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   10710
         TabIndex        =   28
         Top             =   5640
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(Esc) Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   11460
         TabIndex        =   26
         Top             =   660
         Width           =   945
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "(F8) Use Scanner"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   2490
         TabIndex        =   25
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "(F7) Lookup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   4020
         TabIndex        =   24
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "(F4) Void"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   7740
         TabIndex        =   23
         Top             =   660
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(F3) - Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   8670
         TabIndex        =   22
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(F2) + Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   9660
         TabIndex        =   21
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblclose 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   12390
         TabIndex        =   20
         ToolTipText     =   "Close Window"
         Top             =   90
         Width           =   165
      End
      Begin VB.Label lblminimise 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   165
         Left            =   10620
         TabIndex        =   19
         ToolTipText     =   "Minimize"
         Top             =   240
         Width           =   195
      End
      Begin VB.Image imgcloseactive 
         Height          =   180
         Left            =   12180
         Picture         =   "frmCashSalesAE.frx":183392
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Tendered"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   8580
         TabIndex        =   18
         Top             =   7020
         Width           =   1890
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2055
         TabIndex        =   17
         Top             =   3675
         Width           =   1245
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sold To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Slip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   0
         Left            =   1560
         TabIndex        =   15
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   8430
         TabIndex        =   14
         Top             =   7320
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmCashSalesAE"
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

Public blnPaid              As Boolean 'Use to determine if transaction is already paid using frmAmountReceive form

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
Dim bStart                  As Boolean
Dim bScannerOn              As Boolean
Dim blnSave                 As Boolean
Dim intQtyOld               As Integer 'Old txtQty Value. Hold when editing qty

Dim nqty    As Double
Dim nprice  As Double
Dim ndisc   As Double
Dim namt    As Double
Dim ntotal  As Double

Private Sub AddToGrid()
On Error GoTo err
    
    Dim RSStockUnit As New Recordset
    
    Dim intTotalOnhand          As Integer
    Dim intTotalIncoming        As Integer
    Dim intTotalOnhInc          As Integer 'Total of Onhand + Incoming
    Dim intExcessQty            As Integer
    
    Dim intSuggestedQty         As Integer
    Dim blnAddIncoming          As Boolean
    Dim intQtyOrdered           As Integer 'hold the value of txtQty
    Dim intCount                As Integer
    
    If txtEntry(2).Text = "" Then txtEntry(2).SetFocus: Exit Sub
    
    Dim CurrRow As Integer

    Dim intStockID As Integer
    
    CurrRow = getFlexPos(Grid, 8, txtEntry(2).Tag)
    intStockID = txtEntry(2).Tag
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & intStockID & " ORDER BY Stock_Unit.Order ASC", CN, adOpenStatic, adLockOptimistic
    
    If toNumber(txtEntry(5).Text) <= 0 Then
        MsgBox "Please enter a valid sales price.", vbExclamation
        txtEntry(5).SetFocus
        Exit Sub
    End If
    
    intQtyOrdered = txtEntry(4).Text
              
    RSStockUnit.Find "UnitID = " & txtEntry(3).Tag
          
    If RSStockUnit!Onhand < intQtyOrdered Then GoSub GetOnhand

Continue:
    'Save to stock card
    Dim RSStockCard As New Recordset

    RSStockCard.CursorLocation = adUseClient
    RSStockCard.Open "SELECT * FROM Stock_Card", CN, adOpenStatic, adLockOptimistic

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 8) = "" Then
                .TextMatrix(1, 0) = txtEntry(2).Text
                .TextMatrix(1, 1) = txtEntry(3).Text
                .TextMatrix(1, 2) = intQtyOrdered 'txtentry(4).Text
                .TextMatrix(1, 3) = toMoney(txtEntry(5).Text)
                .TextMatrix(1, 4) = toMoney(txtEntry(6).Text)
                .TextMatrix(1, 5) = toMoney(txtEntry(7).Text)
                .TextMatrix(1, 6) = toMoney(txtEntry(8).Text)
                .TextMatrix(1, 7) = txtEntry(3).Tag
                .TextMatrix(1, 8) = intStockID
            Else
AddIncoming:
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = txtEntry(2).Text
                .TextMatrix(.Rows - 1, 1) = txtEntry(3).Text
                .TextMatrix(.Rows - 1, 2) = intQtyOrdered 'txtentry(4).Text
                .TextMatrix(.Rows - 1, 3) = toMoney(txtEntry(5).Text)
                .TextMatrix(.Rows - 1, 4) = toMoney(txtEntry(6).Text)
                .TextMatrix(.Rows - 1, 5) = toMoney(txtEntry(7).Text)
                .TextMatrix(.Rows - 1, 6) = toMoney(txtEntry(8).Text)
                .TextMatrix(.Rows - 1, 7) = txtEntry(3).Tag
                .TextMatrix(.Rows - 1, 8) = intStockID
                
                .FillStyle = 1

                .Row = .Rows - 1
                .Colsel = 6
                If blnAddIncoming = True And intCount = 2 Then
                    .CellForeColor = vbBlue
                    
                    blnAddIncoming = False
                End If
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already exist. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                'Restore back the invoice amount and discount
                cIGross = cIGross - toNumber(Grid.TextMatrix(.Rowsel, 4))
                lblTotal.Caption = Format$(cIGross, "#,##0.00")
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.Rowsel, 6))
                lblTot.Caption = Format$(cIAmount, "#,##0.00")
                'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
                cDAmount = cDAmount - toNumber(toNumber(txtEntry(7).Text) / 100) * _
                        (toNumber(toNumber(Grid.TextMatrix(.Rowsel, 2))) * _
                        toNumber(txtEntry(5).Text))
                lblDiscount.Caption = Format$(cDAmount, "#,##0.00")
                
                .TextMatrix(CurrRow, 0) = txtEntry(2).Text
                .TextMatrix(CurrRow, 1) = txtEntry(3).Text
                .TextMatrix(CurrRow, 2) = intQtyOrdered 'txtentry(4).Text
                .TextMatrix(CurrRow, 3) = toMoney(txtEntry(5).Text)
                .TextMatrix(CurrRow, 4) = toMoney(txtEntry(6).Text)
                .TextMatrix(CurrRow, 5) = toMoney(txtEntry(7).Text)
                .TextMatrix(CurrRow, 6) = toMoney(txtEntry(8).Text)
                .TextMatrix(CurrRow, 7) = txtEntry(3).Tag
                .TextMatrix(CurrRow, 8) = intStockID
                                               
                'deduct qty from Stock Unit's table
                RSStockUnit.Filter = "UnitID = " & txtEntry(3).Tag  'getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                
                RSStockUnit!Onhand = RSStockUnit!Onhand + intQtyOld
                
                RSStockUnit.Update
            Else
                Exit Sub
            End If
        End If
        
        RSStockCard.Filter = "StockID = " & intStockID & " AND RefNo2 = '" & txtEntry(1).Text & "'"

        If RSStockCard.RecordCount = 0 Then RSStockCard.AddNew
        
        'Deduct qty solt to stock card
        RSStockCard!Type = "S"
        RSStockCard!UnitID = txtEntry(3).Tag
        RSStockCard!RefNo2 = txtEntry(1).Text
        RSStockCard!Pieces2 = intQtyOrdered
        'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
        RSStockCard!SalesPrice = txtEntry(5).Text
        RSStockCard!StockID = intStockID

        RSStockCard.Update
        
        RSStockUnit.Find "UnitID = " & txtEntry(3).Tag

        'Deduct qty from highest unit breakdown if Onhand is less than qty ordered
        If RSStockUnit!Onhand < intQtyOrdered Then
            DeductOnhand intQtyOrdered, RSStockUnit!Order, True, RSStockUnit
        End If
        
        'deduct qty from Stock Unit's table
        RSStockUnit.Find "UnitID = " & txtEntry(3).Tag
        
        RSStockUnit!Onhand = RSStockUnit!Onhand - intQtyOrdered
        
        RSStockUnit.Update
            
        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtEntry(6).Text)
        lblTotal.Caption = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount + toNumber(txtEntry(8).Text)
        'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
        cDAmount = cDAmount + toNumber(toNumber(txtEntry(7).Text) / 100) * _
                (toNumber(intQtyOrdered * _
                toNumber(txtEntry(5).Text)))
        lblDiscount.Caption = Format$(cDAmount, "#,##0.00")
        lblTot.Caption = Format$(cIAmount, "#,##0.00")
        lblTaxbase.Caption = toMoney(lblTot.Caption / 1.12)
        lblVAT12.Caption = toMoney(lblTot.Caption - lblTaxbase.Caption)
        'Highlight the current row's column
        .Colsel = 6
        'Display a remove button
        If blnAddIncoming = True Then
            intQtyOrdered = intSuggestedQty
            intCount = 2
            GoSub AddIncoming
            
'            blnAddIncoming = False
        End If
        
        'Reset the entry fields
        ResetEntry
    End With
    
    Exit Sub
    
GetOnhand:

    intTotalOnhand = GetTotalQty("Onhand", RSStockUnit!Order, RSStockUnit!Onhand, RSStockUnit)
    If intTotalOnhand < 0 Then
        MsgBox "Insufficient qty.", vbInformation
    Else
        GoSub Continue
    End If
    
    Exit Sub
    
err:
    Prompt_Err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Function DeductOnhand(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, rs As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    rs.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If rs!Onhand >= QtyNeededTemp Then
            If blnDeduct = False Then
                DeductOnhand = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - rs!Onhand) / rs!Qty
            
            rs.MoveFirst
            
            rs.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            rs!Onhand = rs!Onhand - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            rs.MoveFirst
            rs.Find "Order = " & OrderTemp
            
            rs!Onhand = rs!Onhand + (QtyNeededTemp * rs!Qty)
            
            rs.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                DeductOnhand = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        DeductOnhand = False
    End If
End Function

'Get the total Qty onhand, incoming and total of onhand and incoming
Private Function GetTotalQty(strField As String, Order As Integer, intOnhand As Integer, rs As Recordset) As Integer
    Dim strFieldValue As Integer
    Dim intOrder As Integer
    
    GetTotalQty = intOnhand
    
    intOrder = Order - 1
    
    Do Until intOrder < 1
        rs.MoveFirst
        rs.Find "Order = " & intOrder
        
        If strField = "Onhand" Then
            strFieldValue = rs!Onhand
        ElseIf strField = "Incoming" Then
            strFieldValue = rs!Incoming
        Else
            strFieldValue = rs!TotalQty
        End If
        
        GetTotalQty = GetTotalQty + GetTotalUnitQty(Order, intOrder, strFieldValue, rs)
        intOrder = intOrder - 1
    Loop
End Function

'This function is called by GetTotalQty Function
Private Function GetTotalUnitQty(Order As Integer, ByVal Ordertmp As Integer, intOnhand As Integer, rs As Recordset)
    GetTotalUnitQty = 1
    Do Until Order = Ordertmp
        Ordertmp = Ordertmp + 1
        
        rs.MoveNext
        
        GetTotalUnitQty = GetTotalUnitQty * rs!Qty
    Loop
    GetTotalUnitQty = intOnhand * GetTotalUnitQty
End Function

Private Function GetIncoming(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, rs As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    rs.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If rs!Incoming >= QtyNeededTemp Then
            If blnDeduct = False Then
                GetIncoming = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - rs!Incoming) / rs!Qty
            
            rs.MoveFirst
            
            rs.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            rs!Incoming = rs!Incoming - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            rs.MoveFirst
            rs.Find "Order = " & OrderTemp
            
            rs!Incoming = rs!Incoming + (QtyNeededTemp * rs!Qty)
            
            rs.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                GetIncoming = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        GetIncoming = False
    End If
End Function

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.Rowsel)
    End With
End Sub

Private Function GetUnitID(ByVal sUnit As String) As Long
  Dim rs As New ADODB.Recordset
  Dim SQL As String
  
  
  SQL = "SELECT Unit.unit_id From unit WHERE (((Unit.unit)='" & Replace(sUnit, "'", "''") & "'))"
  
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    GetUnitID = rs!unit_id
  Else
    GetUnitID = 0
  End If
  
  
  rs.Close
  Set rs = Nothing
End Function

Private Function getunitid1(ByVal Stock As String) As Long
  Dim rs As New ADODB.Recordset
  Dim SQL As String
  
  
  SQL = "SELECT Unit.unit_id " _
  & "FROM Stocks LEFT JOIN Unit ON Stocks.unit_id = Unit.unit_id " _
  & "WHERE (((Stocks.stock)='" & Replace(Stock, "'", "''") & "'))"

  
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    getunitid1 = IIf(IsNull(rs!unit_id), 5, rs!unit_id)
  Else
    getunitid1 = 0
  End If
  
  
  rs.Close
  Set rs = Nothing
End Function


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

Private Sub cmdPrint_Click()
    If blnSave = False Then
        btnLabel.Caption = "Settle a transaction first before printing a receipt."
        Exit Sub
    End If
    
    PrintInvoice txtEntry(0).Text
End Sub

Private Sub PrintInvoice(InvoiceNo As String)
    With frmReports
        .strReport = "Receipt"
        .strWhere = "{Cash_Sales.InvoiceNo} ='" & InvoiceNo & "'"
        
        frmReports.show vbModal
    End With
End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Function makecashslip() As String
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  Dim Temp As String
  
  SQL = "SELECT Last(Cash_Sales.cash_sales_id) AS LastOfcash_sales_id " _
  & "From Cash_Sales " _
  & "ORDER BY Last(Cash_Sales.cash_sales_id)"
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    Temp = IIf(IsNull(rs!LastOfcash_sales_id), 0, rs!LastOfcash_sales_id) + 1
  Else
    Temp = 1
  End If

  makecashslip = Format(Temp, "0000000000")
  rs.Close
  Set rs = Nothing
End Function

Private Sub cmdSettle_Click()
    Dim rs As New Recordset
        
    If blnPaid = False Then
        btnLabel.Caption = "Transaction not yet paid."
        Exit Sub
    End If
    
    rs.Open "SELECT * FROM Cash_Sales WHERE CashSalesID=" & PK, CN, adOpenStatic, adLockOptimistic

    'Verify the entries
    If txtEntry(0).Text = "" Then
        MsgBox "Please enter Cash Slip.", vbExclamation
        txtEntry(0).SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to purchase before you can save this record.", vbExclamation
        txtEntry(2).SetFocus
        Exit Sub
    End If
    
    If isRecordExist("Cash_Sales", "InvoiceNo", txtEntry(0).Text, True) = True Then
'        MsgBox "Cash slip already exist. Please change it.", vbExclamation
        btnLabel.Caption = "Transaction are already save."
        
        txtEntry(0).SetFocus
        Exit Sub
    End If

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Cash_Sales_Detail WHERE CashSalesID=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    'Save the record
    With rs
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![CashSalesID] = PK
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        Else
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        ![InvoiceNo] = txtEntry(0).Text
        ![Date] = Now()
        ![SoldTo] = txtEntry(1).Text
        ![Gross] = toNumber(lblTotal.Caption)
        ![Discount] = lblDiscount.Caption
        ![TaxBase] = lblTaxbase.Caption
        ![Vat] = lblVAT12.Caption
        ![NetAmount] = lblTot.Caption
        ![Tendered] = txtEntry(9).Text
        ![Change] = lblChange.Caption
        ![Remarks] = txtEntry(10).Text

        .Update

    End With

    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then

                RSDetails.AddNew

                'RSDetails![PK] = getIndex("tbl_AR_InvoiceDetails")

                RSDetails![CashSalesID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 8))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 2))
                RSDetails![UnitID] = toNumber(.TextMatrix(c, 7))
                RSDetails![Price] = toNumber(.TextMatrix(c, 3))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 5))
                
                RSDetails.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set RSDetails = Nothing

    CN.CommitTrans

    blnSave = True

    HaveAction = True
    Screen.MousePointer = vbDefault

    btnLabel.Caption = "Changes in record has been successfully saved."
    
    PrintInvoice txtEntry(0).Text
    
    Exit Sub
err:
    blnSave = False

'    CN.RollbackTrans
    Prompt_Err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Function getcashsalesid() As Long
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  
  SQL = "SELECT Cash_Sales.cash_sales_id " _
  & "From Cash_Sales " _
  & "WHERE (((Cash_Sales.cash_slip)='" & Replace(txtEntry(0).Text, "'", "''") & "') AND " _
  & "((Cash_Sales.sold_to)='" & Replace(txtEntry(1).Text, "'", "''") & "'))"
  
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    getcashsalesid = rs!Cash_Sales_ID
  Else
    getcashsalesid = 0
  End If
  
  rs.Close
  Set rs = Nothing
End Function

Private Function GetStockID(ByVal Stock As String) As Long
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  
  SQL = "SELECT Stocks.stock_id " _
  & "From Stocks " _
  & "WHERE (((Stocks.stock)='" & Replace(Stock, "'", "''") & "'))"
  rs.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    GetStockID = rs!stock_id
  Else
    GetStockID = 0
  End If
  
  rs.Close
  Set rs = Nothing
End Function

Private Function saveslipdetail(ByVal b As Boolean) As Boolean
On Error GoTo errhandler
  Dim i As Long
  Dim SQL As String
  
  
  With Grid
    For i = 1 To .Rows - 1
      If b Then
        SQL = "insert into cash_sales_detail(cash_sales_id, stock_id, qty, unit, price, " _
        & "amount_gross, discount_percent, discount_amount, amount_net, net_price, [date]) " _
        & "values(" & getcashsalesid & ", " _
        & GetStockID(.TextMatrix(i, 1)) & ", " _
        & .TextMatrix(i, 2) & ", " _
        & getunitid1(.TextMatrix(i, 1)) & ", " _
        & CDbl(.TextMatrix(i, 3)) & ", " _
        & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) & ", " _
        & .TextMatrix(i, 4) & ", " _
        & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100) & ", " _
        & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) - (CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100)) & ", " _
        & "0, #" _
        & Format(Date, "mm/dd/yy") & "#)"
      Else
        'sql = "update cash_sales_detail set " _
        & "cash_sales_id=" & getcashsalesid & ", " _
        & "stock_id=" & getstockid(.TextMatrix(i, 1)) & ", " _
        & "qty=" & .TextMatrix(i, 2) & ", " _
        & "unit=0, " _
        & "price=" & CDbl(.TextMatrix(i, 3)) & ", " _
        & "amount_gross=" & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) & ", " _
        & "discount_percent=" & .TextMatrix(i, 4) & ", " _
        & "discount_amount=" & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100) & ", " _
        & "amount_net=" & CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) - (CDbl(.TextMatrix(i, 2) * .TextMatrix(i, 3)) * (.TextMatrix(i, 4) / 100)) & ", " _
        & "[date]=#" & Format(Date, "mm/dd/yy") & "#"
      End If
      CN.Execute SQL
    Next
  End With
  saveslipdetail = True
  
  Exit Function
errhandler:
  MsgBox "Error: " & err.Description & vbCr _
  & "Form: frmCashSalesAE" & vbCr _
  & "Function: saveslipdetail", vbExclamation
  saveslipdetail = False
End Function

Private Function saveslip(ByVal b As Boolean) As Boolean
On Error GoTo errhandler
  Dim SQL As String
  
  If b Then
    SQL = "insert into cash_sales(cash_slip, [date], sold_to, payment_type, gross, " _
    & "discount, tax_base, vat, amount_net, cash, remarks) " _
    & "values('" & Replace(txtEntry(0).Text, "'", "''") & "', #" _
    & Format(Date & " " & Time, "mm/dd/yy hh:mm:ss") & "#, '" _
    & Replace(txtEntry(1).Text, "'", "''") & "', '', " _
    & CDbl(lblTotal.Caption) & ", 0, 0, " _
    & CDbl(lblVAT12.Caption) & ", " _
    & CDbl(lblTot.Caption) & ", " _
    & CDbl(txtEntry(8).Text) & ", '" _
    & Replace(txtEntry(9).Text, "'", "''") & "')"
  Else
    'sql = "update cash_sales set " _
    & "cash_slip='" & Replace(txtEntry(0).Text, "'", "''") & "', " _
    & "[date]=#" & Format(Date, "mm/dd/yy") & "#, " _
    & "sold_to='" & Replace(txtEntry(1).Text, "'", "''") & "', " _
    & "payment_type='', " _
    & "gross=" & CDbl(lblTotal.Caption) & ", " _
    & "discount=0, " _
    & "tax_base=0, " _
    & "vat=" & CDbl(lblVAT12.Caption) & ", " _
    & "amount_net=" & CDbl(lblTot.Caption) & ", " _
    & "cash=" & CDbl(txtEntry(8).Text) & ", '" _
    & "remarks='" & Replace(txtEntry(9).Text, "'", "''") & "'"
  End If
  
  CN.Execute SQL
  saveslip = True
  
  Exit Function
  
errhandler:
  MsgBox "Error: " & err.Description & vbCr _
  & "Form: frmCashSalesAE" & vbCr _
  & "Function: saveslip", vbExclamation
  saveslip = False
End Function



Private Sub cmdSettle_KeyDown(KeyCode As Integer, Shift As Integer)
'  Form_KeyDown KeyCode, Shift
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    End If
End Sub

Private Sub addtolist()
  
  With Grid
    If .Rows = 2 And .TextMatrix(1, 0) = "" Then
      .TextMatrix(1, 0) = txtEntry(2).Text
      .TextMatrix(1, 1) = txtEntry(3).Text
      .TextMatrix(1, 2) = txtEntry(4).Text
      .TextMatrix(1, 3) = Format(txtEntry(5).Text, "#,##0.00")
      .TextMatrix(1, 4) = txtEntry(6).Text
      .TextMatrix(1, 5) = Format(txtEntry(7).Text, "#,##0.00")
    Else
      .Rows = .Rows + 1
      .TextMatrix(.Rows - 1, 0) = txtEntry(2).Text
      .TextMatrix(.Rows - 1, 1) = txtEntry(3).Text
      .TextMatrix(.Rows - 1, 2) = txtEntry(4).Text
      .TextMatrix(.Rows - 1, 3) = Format(txtEntry(5).Text, "#,##0.00")
      .TextMatrix(.Rows - 1, 4) = txtEntry(6).Text
      .TextMatrix(.Rows - 1, 5) = Format(txtEntry(7).Text, "#,##0.00")
    End If
  End With
  txtEntry(2).Text = ""
  txtEntry(3).Text = ""
  txtEntry(4).Text = "1"
  txtEntry(5).Text = "0.00"
  txtEntry(6).Text = "0"
  txtEntry(7).Text = "0.00"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If State = adStateEditMode Or State = adStateAddMode Then Exit Sub
    
    Select Case KeyCode
      Case vbKeyEscape
        blnSave = False
        
        lblclose_Click
      
      Case vbKeyF4
        Form_KeyDown vbKeyF9, Shift
      Case vbKeyF5
        frmAmountReceive.txtAmt.Text = Format(lblTot.Caption, "#,##0.00")
        frmAmountReceive.show 1
      Case vbKeyF6

        txtEntry_KeyPress 7, 13
        
      Case vbKeyF7        'lookup table
        If Not bStart Then
          btnLabel.Caption = "Please start transaction first by pressing (F9) key."
          Exit Sub
        End If
        frmLookup.show 1
        txtEntry(2).SetFocus
      Case vbKeyF8        'ready scanner for scanning
        If Not bStart Then
          MsgBox "Please start transaction first by pressing (F9) key.", vbInformation
          Exit Sub
        End If
        
        txtEntry(2).SetFocus
      Case vbKeyF9
        bStart = True
        blnSave = False
        blnPaid = False
        
        txtEntry(0).Text = ""
        txtEntry(1).Text = ""
        
        ResetEntry
        
        txtEntry(0).SetFocus
        ntotal = 0
        
        InitGrid
        
        CN.BeginTrans
        
        GeneratePK
        
        txtEntry(0).Text = Format(PK, "0000000000")
        
        txtEntry(1).SetFocus
        
        cIGross = 0
        cIAmount = 0
        lblTotal.Caption = "0.00"
        lblDiscount.Caption = "0.00"
        lblTaxbase.Caption = "0.00"
        lblVAT12.Caption = "0.00"
        lblTot.Caption = "0.00"
        lblChange.Caption = "0.00"
        
      Case vbKeyF11
        cmdSettle_Click
      Case vbKeyF12
        cmdPrint_Click
      
      
      Case vbKeyEnd
      Case vbKeyHome
      Case vbKeyUp, vbKeyPageUp
        'If Shift = vbCtrlMask Then
        '  cmdFirst_Click
        'Else
        '  cmdPrevious_Click
        'End If
      Case vbKeyDown, vbKeyPageDown
        'If Shift = vbCtrlMask Then
        '  cmdLast_Click
        'Else
        '  cmdNext_Click
        'End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid
    
    bStart = False
    imgcloseactive.Left = 12385: imgcloseactive.Top = 105
         
    Picture1.Left = (Me.ScaleWidth / 2) - (Picture1.ScaleWidth / 2)
    Picture1.Top = (Me.ScaleHeight / 2) - (Picture1.ScaleHeight / 2)
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Cash_Sales")
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        '.FixedCols = 1
        .Cols = 9 '13
        .Colsel = 7
        'Initialize the column size
        .ColWidth(0) = 1455
        .ColWidth(1) = 5460
        .ColWidth(2) = 660
        .ColWidth(3) = 870
        .ColWidth(4) = 870
        .ColWidth(5) = 885
        .ColWidth(6) = 1545
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = "Barcode"
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "Qty"
        .TextMatrix(0, 3) = "Price"
        .TextMatrix(0, 4) = "Gross"
        .TextMatrix(0, 5) = "Discount"
        .TextMatrix(0, 6) = "Net Amount"
        .TextMatrix(0, 7) = "Unit ID"
        .TextMatrix(0, 8) = "Stock ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbRightJustify
        .ColAlignment(3) = vbRightJustify
        .ColAlignment(4) = vbRightJustify
        .ColAlignment(5) = vbRightJustify
        .ColAlignment(6) = vbRightJustify
    End With
End Sub

Private Sub ResetEntry()
    txtEntry(2).Text = ""
    txtEntry(3).Text = ""
    txtEntry(4).Text = "1"
    txtEntry(5).Text = "0.00"
    txtEntry(6).Text = "0"
    txtEntry(7).Text = "0.00"
    txtEntry(8).Text = "0.00"
    txtEntry(9).Text = "0.00"
'    lblTotal.Caption = "0.00"
'    lblDiscount.Caption = "0.00"
'    lblTaxbase.Caption = "0.00"
'    lblVAT12.Caption = "0.00"
'    lblTot.Caption = "0.00"
'    lblChange.Caption = "0.00"
    
'    SendKeys ("{tab}")
'    SendKeys ("{tab}")

    txtEntry(2).SetFocus
End Sub

Private Sub Form_Resize()
  Picture1.Left = (Me.ScaleWidth / 2) - (Picture1.ScaleWidth / 2)
  Picture1.Top = (Me.ScaleHeight / 2) - (Picture1.ScaleHeight / 2)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCashSalesAE = Nothing
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub lblclose_Click()
On Error Resume Next

    If blnSave = False Then CN.RollbackTrans
    Unload Me
End Sub

Private Sub lblclose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgcloseactive.Visible = True
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
  Form_KeyDown KeyCode, Shift
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgcloseactive.Visible = False
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index = 2 And Len(txtEntry(2).Text) = 13 Then
        rs.Open "SELECT StockID, Barcode, Stock, SalesPrice, UnitID FROM qry_Stock_Unit WHERE Barcode = '" & txtEntry(2).Text & "' ORDER BY [Order] DESC", CN, adOpenStatic, adLockOptimistic

        With rs
            If .RecordCount = 0 Then
                MsgBox "No available stock.", vbInformation
                
                .Close
                Exit Sub
            End If
            
            txtEntry(2).Tag = !StockID
            txtEntry(3).Tag = !UnitID 'Add Unit ID to Stock control to avoid adding of dummy control
            txtEntry(3).Text = !Stock
            txtEntry(5).Text = toMoney(!SalesPrice)
            
            .Close
        End With
    End If
    
    If Index = 4 Or Index = 5 Or Index = 7 Then
        txtEntry(6).Text = toMoney((toNumber(txtEntry(4).Text) * toNumber(txtEntry(5).Text)))
        txtEntry(8).Text = toMoney((toNumber(txtEntry(4).Text) * _
                toNumber(txtEntry(5).Text)) - _
                ((toNumber(txtEntry(7).Text) / 100) * _
                toNumber(toNumber(txtEntry(4).Text) * _
                toNumber(txtEntry(5).Text))))
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
    If Index = 8 Then
'        cmdSave.Default = False
    End If
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 7 Then
        AddToGrid
    End If
    
    If Index > 1 And Index < 7 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index = 5 Then
        txtEntry(Index).Text = toMoney(txtEntry(Index).Text)
    End If
End Sub

VERSION 5.00
Begin VB.Form frmAmountReceive 
   ClientHeight    =   1350
   ClientLeft      =   3555
   ClientTop       =   4530
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "frmAmountReceive.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   390
      Width           =   4425
   End
   Begin VB.Label Label2 
      Caption         =   "(Enter) Accept        (Esc) Cancel"
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   990
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   2715
   End
End
Attribute VB_Name = "frmAmountReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtAmt_GotFocus()
  HLText txtAmt
End Sub

Private Sub txtAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyEscape
    frmCashSalesAE.blnPaid = False
    Unload Me
  Case 13
    If CDbl(txtAmt) < CDbl(frmCashSalesAE.lblTot.Caption) Then
      MsgBox "Amount entered is less than total sales price.", vbExclamation
      Exit Sub
    End If
    
    frmCashSalesAE.blnPaid = True
    
    frmCashSalesAE.txtEntry(9).Text = Format(txtAmt, "#,##0.00")
    frmCashSalesAE.lblChange.Caption = Format(frmCashSalesAE.txtEntry(9).Text - frmCashSalesAE.lblTot, "#,##0.00")
    Unload Me
  End Select
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
  KeyAscii = AllowOnlyNumbers(KeyAscii, txtAmt)
End Sub

Private Sub txtAmt_LostFocus()
  If Trim(txtAmt) = "" Then txtAmt = "0.00"
  
  txtAmt = Format(txtAmt, "#,##0.00")
End Sub

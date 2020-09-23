VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preview"
   ClientHeight    =   7080
   ClientLeft      =   2595
   ClientTop       =   2460
   ClientWidth     =   10425
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5580
      TabIndex        =   6
      Top             =   420
      Width           =   1065
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   315
      Left            =   4470
      TabIndex        =   5
      Top             =   420
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "As Of"
      Height          =   1035
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM dd, yyyy"
         Format          =   67895299
         CurrentDate     =   39036
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2100
         TabIndex        =   4
         Top             =   540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM dd, yyyy"
         Format          =   67895299
         CurrentDate     =   39036
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2100
         TabIndex        =   3
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   525
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   3915
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   6405
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mTest As CRAXDRT.Application
Dim mReport As CRAXDRT.Report
Dim SubReport As CRAXDRT.Report
Dim mParam As CRAXDRT.ParameterFieldDefinitions

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
    Set mTest = New CRAXDRT.Application
    Set mReport = New CRAXDRT.Report

    Select Case ReportType
    Case "Periodic Cash Sales Itemized"
        Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPeriodicCashSalesItemized.rpt")
        
        mReport.RecordSelectionFormula = "{Cash_Sales_Itemized.Date} IN #" & dtpFrom.Value & "# TO #" & dtpTo.Value & "#"
                
        Set mParam = mReport.ParameterFields
        
        mParam.Item(1).AddCurrentValue "PERIODIC CHARGE SALES From " & Format(dtpFrom.Value, "mm/dd/yy") & " to " & Format(dtpTo.Value, "mm/dd/yy")
    Case "Periodic Cash Sales Return"
        Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPeriodicCashSalesReturn.rpt")
        
        mReport.RecordSelectionFormula = "{Cash_Sales_Return.Date} IN #" & dtpFrom.Value & "# TO #" & dtpTo.Value & "#"
        
        Set mParam = mReport.ParameterFields
        
        mParam.Item(1).AddCurrentValue "PERIODIC CASH SALES RETURN From " & Format(dtpFrom.Value, "mm/dd/yy") & " to " & Format(dtpTo.Value, "mm/dd/yy")
    Case "Periodic Cash Sales Return (Itemized)"
        Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPeriodicCashSalesReturnItemized.rpt")
        
        mReport.RecordSelectionFormula = "{Cash_Sales_Return.Date} IN #" & dtpFrom.Value & "# TO #" & dtpTo.Value & "#"

        Set mParam = mReport.ParameterFields
        
        mParam.Item(1).AddCurrentValue "PERIODIC CASH SALES RETURN ITEMIZED From " & Format(dtpFrom.Value, "mm/dd/yy") & " to " & Format(dtpTo.Value, "mm/dd/yy")
    Case "Periodic Charge Sales"
        Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPeriodicChargeSales.rpt")
        
        mReport.RecordSelectionFormula = "{Sales_Order.Date} IN #" & dtpFrom.Value & "# TO #" & dtpTo.Value & "#"

        Set mParam = mReport.ParameterFields
        
        mParam.Item(1).AddCurrentValue "PERIODIC CHARGE SALES From " & Format(dtpFrom.Value, "mm/dd/yy") & " to " & Format(dtpTo.Value, "mm/dd/yy")
    Case "Periodic Charge Sales (Itemized)"
        Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPeriodicChargeSalesItemized.rpt")
        
        mReport.RecordSelectionFormula = "{Sales_Order.Date} IN #" & dtpFrom.Value & "# TO #" & dtpTo.Value & "#"
        
        Set mParam = mReport.ParameterFields
        
        mParam.Item(1).AddCurrentValue "PERIODIC CHARGE SALES ITEMIZED From " & Format(dtpFrom.Value, "mm/dd/yy") & " to " & Format(dtpTo.Value, "mm/dd/yy")
    Case "Periodic Charge Sales Returns"
        Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPeriodicChargeSalesReturn.rpt")
        
        mReport.RecordSelectionFormula = "{Sales_Order_Return.CreditMemoDate} IN #" & dtpFrom.Value & "# TO #" & dtpTo.Value & "#"
        
        Set mParam = mReport.ParameterFields
        
        mParam.Item(1).AddCurrentValue "PERIODIC CHARGE SALES RETURNS From " & Format(dtpFrom.Value, "mm/dd/yy") & " to " & Format(dtpTo.Value, "mm/dd/yy")
    Case "Periodic Charge Sales Returns (Itemized)"
        Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPeriodicChargeSalesReturnItemized.rpt")
        
        mReport.RecordSelectionFormula = "{Sales_Order_Return.CreditMemoDate} IN #" & dtpFrom.Value & "# TO #" & dtpTo.Value & "#"
        
        Set mParam = mReport.ParameterFields
        
        mParam.Item(1).AddCurrentValue "PERIODIC CHARGE SALES RETURNS (ITEMIZED) From " & Format(dtpFrom.Value, "mm/dd/yy") & " to " & Format(dtpTo.Value, "mm/dd/yy")
    End Select

    Screen.MousePointer = vbHourglass
    CR.ReportSource = mReport
    CR.ViewReport
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  dtpFrom.Value = Date   'Month(Date) & "/1/" & Year(Date)
  dtpTo.Value = Date  ' Month(Date) & "/1/" & Year(Date)
End Sub

Private Sub Form_Resize()
    With CR
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReports 
   Caption         =   "Report"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   3915
      Left            =   840
      TabIndex        =   0
      Top             =   720
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
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strReport        As String
Public PK               As String
Public strYear          As String
Public blnPaid          As Boolean
Public strWhere         As String

Dim mTest As CRAXDRT.Application
Dim mReport As CRAXDRT.Report
Dim SubReport As CRAXDRT.Report
Dim mParam As CRAXDRT.ParameterFieldDefinitions

Public Sub CommandPass(ByVal srcPerformWhat As String)
    Select Case srcPerformWhat
        Case "Close"
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load
    Dim mSubRep
    
    Set mTest = New CRAXDRT.Application
    Set mReport = New CRAXDRT.Report
    
    Select Case strReport
        Case "Print Barcode"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\Barcode.rpt")
            
            mReport.RecordSelectionFormula = strWhere
        Case "Receipt"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptCashInvoice.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Sales Order"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptSalesOrder.rpt")
            
            mReport.RecordSelectionFormula = strWhere
        Case "Local Purchase"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptLocalPurchase.rpt")
            
            mReport.RecordSelectionFormula = strWhere
        Case "Purchase Order"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPurchaseOrder.rpt")
            
            mReport.RecordSelectionFormula = strWhere
        Case "Receiving Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptReceivingReport.rpt")
            
            mReport.RecordSelectionFormula = strWhere
        Case "Local Purchase Return"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptLocalPurchaseReturn.rpt")
            
            mReport.RecordSelectionFormula = strWhere
        Case "Purchase Order Return"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptPurchaseOrderReturn.rpt")
            
            mReport.RecordSelectionFormula = strWhere
        Case "Stock Card"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rptStockCard.rpt")
            
            If strWhere <> "" Then _
                mReport.RecordSelectionFormula = strWhere
        
        End Select
    
    Screen.MousePointer = vbHourglass
    CR.ReportSource = mReport
    CR.ViewReport
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err_Form_Load:
    Prompt_Err err, Name, "Form_Load"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CR
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReports = Nothing
End Sub



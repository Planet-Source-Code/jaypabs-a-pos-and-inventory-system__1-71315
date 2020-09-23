VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UOM"
   ClientHeight    =   5730
   ClientLeft      =   615
   ClientTop       =   1620
   ClientWidth     =   2040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   2040
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvList 
      Height          =   5655
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
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
         Text            =   "Packaging"
         Object.Width           =   2434
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmSelectUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim rs As New Recordset
  rs.CursorLocation = adUseClient
  rs.Open "SELECT * FROM Unit ORDER BY Unit", CN, adOpenStatic, adLockOptimistic
  
  Do While Not rs.EOF
    lvList.ListItems.Add
    lvList.ListItems(lvList.ListItems.Count).SubItems(1) = lvList.ListItems.Count
    lvList.ListItems(lvList.ListItems.Count).SubItems(2) = rs!Unit
    lvList.ListItems(lvList.ListItems.Count).SubItems(3) = rs!UnitID
        
    rs.MoveNext
  Loop
End Sub

Private Sub lvList_KeyPress(KeyAscii As Integer)
  If lvList.ListItems.Count > 1 Then
    If KeyAscii = 13 Then
      frmCashSalesAE.Unit = lvList.ListItems(lvList.SelectedItem.Index).SubItems(2)
      Unload Me
    ElseIf KeyAscii = 27 Then
      Unload Me
    End If
  End If
End Sub

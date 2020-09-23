VERSION 5.00
Begin VB.Form frmClientsCategoryAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Entry"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Left            =   1515
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "Category"
      Top             =   480
      Width           =   3390
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2175
      TabIndex        =   3
      Top             =   1395
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3615
      TabIndex        =   2
      Top             =   1395
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   1395
      Width           =   1680
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
      _ExtentX        =   10213
      _ExtentY        =   53
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmClientsCategoryAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With rs
        txtEntry.Text = .Fields("Category")
    End With
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    txtEntry.SetFocus
End Sub

Private Sub cmdSave_Click()
    If State = adStateAddMode Then
        rs.AddNew
        rs.Fields("DateAdded") = Now
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    'Phill 2:12
    With rs
        .Fields("category") = txtEntry.Text
    
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(rs.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Clients_Category WHERE CategoryID = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmClientsCategory.RefreshRecords
        End If
    End If
    
    Set frmClientsCategoryAE = Nothing
End Sub


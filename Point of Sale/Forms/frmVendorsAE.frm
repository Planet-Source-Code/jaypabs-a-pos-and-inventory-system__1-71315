VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVendorsAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendors Record"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbGender 
      Height          =   315
      ItemData        =   "frmVendorsAE.frx":0000
      Left            =   1620
      List            =   "frmVendorsAE.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2775
      Width           =   1290
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   330
      TabIndex        =   16
      Top             =   3990
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   15
      Top             =   2415
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   8
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   14
      Top             =   510
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   6165
      MaxLength       =   10
      TabIndex        =   13
      Top             =   120
      Width           =   1290
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   12
      Top             =   2025
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   11
      Top             =   1650
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1620
      MaxLength       =   200
      TabIndex        =   10
      Top             =   1275
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   9
      Tag             =   "Name"
      Top             =   540
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6945
      TabIndex        =   8
      Top             =   3990
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   5505
      TabIndex        =   7
      Top             =   3990
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1620
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   165
      Width           =   1965
   End
   Begin VB.TextBox txtEntry 
      Height          =   615
      Index           =   6
      Left            =   1620
      TabIndex        =   5
      Top             =   3165
      Width           =   3855
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   9
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   4
      Top             =   885
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   10
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1245
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   11
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1605
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   12
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1965
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   13
      Left            =   6165
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2325
      Width           =   2490
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   300
      TabIndex        =   18
      Top             =   3885
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcCategory 
      Height          =   315
      Left            =   1620
      TabIndex        =   19
      Top             =   870
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   11
      Left            =   270
      TabIndex        =   35
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Gender"
      Height          =   240
      Index           =   10
      Left            =   420
      TabIndex        =   34
      Top             =   2775
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Landline"
      Height          =   240
      Index           =   7
      Left            =   4665
      TabIndex        =   33
      Top             =   510
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Middlename"
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   2415
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Mobile"
      Height          =   240
      Index           =   5
      Left            =   4665
      TabIndex        =   31
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Firstname"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   30
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Lastname"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   29
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "TIN"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company"
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   27
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account No."
      Height          =   240
      Index           =   0
      Left            =   570
      TabIndex        =   26
      Top             =   165
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   12
      Left            =   420
      TabIndex        =   25
      Top             =   3165
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax"
      Height          =   240
      Index           =   9
      Left            =   4665
      TabIndex        =   24
      Top             =   885
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Email"
      Height          =   240
      Index           =   13
      Left            =   4665
      TabIndex        =   23
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Website"
      Height          =   240
      Index           =   14
      Left            =   4665
      TabIndex        =   22
      Top             =   1605
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting Balance"
      Height          =   240
      Index           =   8
      Left            =   4665
      TabIndex        =   21
      Top             =   1965
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Limit"
      Height          =   240
      Index           =   15
      Left            =   4665
      TabIndex        =   20
      Top             =   2325
      Width           =   1365
   End
End
Attribute VB_Name = "frmVendorsAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode
Public srcTextAdd           As TextBox 'Used in pop-up mode -> Display the customer address
Public srcTextCP            As TextBox 'Used in pop-up mode -> Display the customer contact person
Public srcTextDisc          As Object  'Used in pop-up mode -> Display the customer Discount (can be combo or textbox)

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With rs
        txtEntry(0).Text = .Fields("AccountNo")
        txtEntry(1).Text = .Fields("Company")
        dcCategory.BoundText = .Fields![CategoryID]
        txtEntry(2).Text = .Fields("Tin")
        txtEntry(3).Text = .Fields("Lastname")
        txtEntry(4).Text = .Fields("Firstname")
        txtEntry(5).Text = .Fields("Middlename")
        cmbGender.Text = .Fields("Gender")
        txtEntry(6).Text = .Fields("Address")
        txtEntry(7).Text = .Fields("Mobile")
        txtEntry(8).Text = .Fields("Landline")
        txtEntry(9).Text = .Fields("Fax")
        txtEntry(10).Text = .Fields("Email")
        txtEntry(11).Text = .Fields("Website")
        txtEntry(12).Text = .Fields("StartingBalance")
        txtEntry(13).Text = .Fields("CreditLimit")
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
    
    txtEntry(0).SetFocus
End Sub


Private Sub cmdSave_Click()
On Error GoTo err

    If Trim(txtEntry(0).Text) = "" Then
      MsgBox "Please provide account number.", vbExclamation
      Exit Sub
    End If
    If Trim(txtEntry(1).Text) = "" Then
      MsgBox "Please provide company name.", vbExclamation
      Exit Sub
    End If
        
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("AccountNo") = txtEntry(0).Text
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("Company") = txtEntry(1).Text
        .Fields("CategoryID") = dcCategory.BoundText
        .Fields("tin") = txtEntry(2).Text
        .Fields("Lastname") = txtEntry(3).Text
        .Fields("Firstname") = txtEntry(4).Text
        .Fields("Middlename") = txtEntry(5).Text
        .Fields("Gender") = cmbGender.Text
        .Fields("Address") = txtEntry(6).Text
        .Fields("Mobile") = txtEntry(7).Text
        .Fields("Landline") = txtEntry(8).Text
        .Fields("Fax") = txtEntry(9).Text
        .Fields("Email") = txtEntry(10).Text
        .Fields("Website") = txtEntry(11).Text
        .Fields("StartingBalance") = toNumber(txtEntry(12).Text)
        .Fields("CreditLimit") = toNumber(txtEntry(13).Text)

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
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
    
    
    
    
    
    Exit Sub
err:
  MsgBox err.Description
        If err.Number = -2147217887 Then Resume Next
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
    rs.Open "SELECT * FROM Vendors WHERE VendorID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Vendors_Category", "Category", dcCategory, "CategoryID", True
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmVendors.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmVendorsAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub





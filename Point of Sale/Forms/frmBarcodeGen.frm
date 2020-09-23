VERSION 5.00
Begin VB.Form frmBarcodeGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   1455
   ClientTop       =   3480
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7380
   Begin VB.PictureBox picEan 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   270
      ScaleHeight     =   37.607
      ScaleMode       =   0  'User
      ScaleWidth      =   108.91
      TabIndex        =   2
      Top             =   270
      Width           =   3225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Make Barcode"
      Height          =   345
      Left            =   2460
      TabIndex        =   1
      Top             =   1290
      Width           =   1215
   End
   Begin VB.TextBox txtBarcode 
      Height          =   345
      Left            =   270
      TabIndex        =   0
      Text            =   "0000000002301"
      Top             =   1290
      Width           =   2145
   End
End
Attribute VB_Name = "frmBarcodeGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_sBarcode As String, m_lBarcodeLength As Long
Private WithEvents HO As Class1
Attribute HO.VB_VarHelpID = -1
Private sFileName As String
Private DocumentDB As Database
Dim Short1  As String
Dim Short2  As String
Dim ICode   As String

Private Sub GetItemInfo(ByVal Barcode As String)
  Dim sql As String
  Dim rstemp As Recordset
  
  
End Sub

Private Sub CreateBarcode()
  On Error GoTo errHandler                            'Error Handling function
  
  Dim bytBarcodeType As Byte, sTemp As String         'Initiate variables
  With txtBarcode
    Select Case Len(.Text)
    Case 0 To 6:
      'Alert "Enter 7+ numbers into the text box": Exit Sub    '6 or less numbers entered
    Case 7 To 11:
      bytBarcodeType = 7                                      'EAN 8 barcode
      m_lBarcodeLength = 8
    Case 12 To 20:
      bytBarcodeType = 12                                     'EAN 13 barcode
      m_lBarcodeLength = 13
    End Select
  
    m_sBarcode = MakeBarcode(Left(.Text, bytBarcodeType))           'Puts correct checkdigit on barcode root.
    .Text = m_sBarcode                                              'Full EAN code
    DrawEan                                                         'Draw the barcode!
  
  End With
  Exit Sub
errHandler:
  'Select Case err.Number
  'Case 13: Resume Next 'Alert "Enter only numbers into text box!"   'In case someone puts other characters then numbers into textbox
  'Case Else: Alert "Error occurred: " & err.Description   'Any other error, die nicely
  'End Select
  Resume Next
End Sub

Private Sub DrawEan()
  Dim bytCentreDigit As Byte, lPositionX As Long, i As Integer, j As Integer
  Dim lCurrNumber As Long, lFirstNumber As Long, iModule As Integer

  bytCentreDigit = IIf(m_lBarcodeLength = 8, 5, 8)     'Where to put the middle bars? EAN8: 5 digit, EAN13: 8th digit (just before each)
  With picEan
    .Cls                  'Clear
    .BackColor = vbWhite  'Set colour
    .FontSize = 6        'Set font size
    .DrawWidth = 1.5     'Set draw width
    
    '.FontUnderline = True
    '.CurrentX = (picEan.ScaleWidth / 2) - ((Len("ALEN MARKETING") * 4) / 2)
    '.CurrentY = 2
    'picEan.Print "ALEN MARKETING"
        
    '.FontUnderline = False
    '.CurrentX = (picEan.ScaleWidth / 2) - ((Len("NICEPACK STATIONERY TAPE") * 3.7) / 2)
    '.CurrentY = 11
    'picEan.Print "NICEPACK STATIONERY TAPE"
    
    '.CurrentX = (picEan.ScaleWidth / 2) - ((Len("8.5 X 11 S-20 ONE REAM") * 3.2) / 2)
    '.CurrentY = 19
    'picEan.Print "8.5 X 11 S-20 ONE REAM"
    
    .FontSize = 6
    lPositionX = 11       'X position (11 =must be 11 modules [1 module = usually 0.33 millimeters, in my case picEan.ScaleWidth <bar width> / 113] 11 on left side, 7 on right side

    For i = 1 To m_lBarcodeLength     '8 or 13 digit code
      lCurrNumber = CLng(Mid(m_sBarcode, i, 1)) 'Current n°
      
      
      If i = 1 Then
        GuardBar lPositionX         'Draw double lines at current X position
        lFirstNumber = lCurrNumber  'This
        .CurrentX = 2
        .CurrentY = 14 ' 66
        picEan.Print IIf(m_lBarcodeLength = 8, "<", lFirstNumber) 'If EAN8, draw "<", else draw number
      End If
      
      
      If i <> 1 Or m_lBarcodeLength = 8 Then
        If i < bytCentreDigit Then                'On the left side, there are modules 1 or 2 (A, B) depending on the 1st digit = [Mdl(0 - 9, 0 or 1)]...
          Select Case m_lBarcodeLength
            Case 8: iModule = 0                   'For EAN 8, always use module 0 (if doesnt work, see start for email addy. Please inform!
            Case 13: iModule = MidInt(MdlLeft(lFirstNumber), i - 1)
          End Select
        Else: iModule = 2                     '...on the right side always module 2 (C) = [Mdl(0 - 9, 2)]
        End If
        
        If i = bytCentreDigit Then                       'Draw the centre pattern
          lPositionX = lPositionX + 2
          GuardBar lPositionX
          lPositionX = lPositionX + 1
        End If
        
        For j = 1 To 7                  '7 modules for each n° (System of 7 black or white sprites)
          If MidInt(Mdl(iModule)(lCurrNumber), j) = 1 Then DrawLine lPositionX, 0  'Draw modules(sprites) for each n°
          lPositionX = lPositionX + 1
        Next j
        .CurrentX = lPositionX - 8
        .CurrentY = 23 ' 66
        picEan.Print lCurrNumber                  'Print n°s
      End If
    Next i
    
    .CurrentX = lPositionX + 8
    .CurrentY = 23 '66
    
    If m_lBarcodeLength = 8 Then picEan.Print ">"
    
    GuardBar lPositionX
    
    'draw ICode
    '.FontSize = 9
    '.CurrentX = (picEan.ScaleWidth / 2) - ((Len("A3SPF-06-06  041") * 2.6) / 2)
    '.CurrentY = 66
    'picEan.Print "A3SPF-06-06  041"
  End With
End Sub

Private Sub Command1_Click()
  CreateBarcode
  'SavePicture picEan.Image, App.Path & "\Barcodes\" & txtBarcode.Text & ".jpg"
  'sFileName = App.Path & "\Barcodes\" & txtBarcode.Text & ".jpg"
  'SaveBinaryObject
End Sub

Private Sub GuardBar(r_lPositionX As Long)
  DrawLine r_lPositionX, 10          '1st guardbar length
  DrawLine r_lPositionX + 2, 10     'last guardbar length
  r_lPositionX = r_lPositionX + 3
End Sub

Private Sub DrawLine(r_lPositionX As Long, r_bytExtension As Byte)
  picEan.Line (r_lPositionX, 3)-(r_lPositionX, 22 + r_bytExtension)
End Sub

Private Sub Form_Load()
  Init
  OpenDB DocumentDB, True
End Sub

Private Sub SaveBinaryObject()
Dim FieldNames(6) As Variant           'names of the other fields to return
Dim FieldData(6) As Variant            'names of the other fields to return

Dim RD() As Variant                    'store for the returned data, not the binary field
Dim FN As String                       'Binary file name to use as storage
Dim i As Integer

    If sFileName = "" Then
        Exit Sub
    End If

    Set HO = New Class1       'create the new bd object

    FieldNames(0) = "ID"               'return the ID field
    FieldNames(1) = "FileName"         'return the filename
    FieldNames(3) = "Barcode"
    FieldNames(4) = "Short1"
    FieldNames(5) = "Short2"
    FieldNames(6) = "ICode"
    
    FieldData(0) = Null                  'return the ID field
    FieldData(1) = sFileName           'return the filename

    With HO
        .KillFile = False                       'kill the filename if it exists
        Set .DB = DocumentDB                   'pass the database
        .ObjectKeyFieldName = "ID"             'the key/index field is
        .ObjectKey = -1                        'the value to search for is
        .ObjectFieldName = "OLEModule"         'name of the field that contains the binary file
        .ObjectTableName = "tblFileObject"     'table that contains the binary files
        .SubFieldNames = FieldNames            'pass in the field names to return
        .SubFieldData = FieldData
        .FileName = sFileName                  'file name to use
        .SaveObject                            'get the file from the database
        .ReturnData RD()                       'return any aditional data
        FN = .FileName                         'actual file name used - if default was used
    End With
    Set HO = Nothing

   ' LoadListBox

    For i = 0 To UBound(RD)
        Debug.Print RD(i)                      'print aditional info returned
    Next
End Sub

Public Sub OpenDB(MyDB As Database, Optional OpenMDB As Boolean = True)


    If OpenMDB Then
        '/* Password protected database file */
        
        Set MyDB = Workspaces(0).OpenDatabase(App.Path & "\Data\Temp.mdb", False, False, "")
        
    Else
        MyDB.Close
        Set MyDB = Nothing
    End If
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************
' Name:         IsFileExist
' Description:  This class saves a binary file to
'               a database of the users choice
'
' By:           Tom Brennfleck
'
' Inputs:
'                Let KillFile(ByVal vData As Boolean)           - Kill the file if its present
'                Let ObjectKeyFieldName(ByVal vData As String)  - key field name to the database
'                Let ObjectTableName(ByVal vData As String)     - table name holding the binary object
'                Let ObjectKey(ByVal vData As Long)             - binary object key to extract
'                Let SubFieldData(ByVal vData As Variant)       - other field data to extract/save
'                Let SubFieldNames(ByVal vData As Variant)      - other field names to extract/save
'                Let ObjectFieldName(ByVal vData As String)     - name of field holding the binary object
'                Set DB(ByVal vData As Database)                - database holding the binary object table
'                Let BlockSize(ByVal vData As Long)             - block size to use
'                Let FileName(ByVal vData As String)            - filename to extract to or import from
'
'
' Returns:
'                Get FileName() As String                       - return file name if it was changed
'                                                                 eg: temp file was used
'                ReturnData(vData() As Variant)                 - returns a variant array of the aditional
'                                                                 database fields that were requested
'
' Methods:       GetObject()                                    - gets the binary object from the database
'                SaveObject()                                   - saves the binary object from the database
'
' Events:        Status(ID As Long, Msg As String)              - returns progress status
'                Error(ID As Long, Msg As String)               - returns error messages
'
'
' Database:
'                TableName      - tblBinaryObject
'                Field          1. ID           -   AutoNumber  - Primary Key
'                               2. FileName     -   Text 255
'                               3. OLEModule    -   OLE Object
'
'   Copyright T. Brennfleck SamSa Consulting Group Pty. Ltd. 2001
'   This code is copyrighted and has limited warranties.
'   Terms of Agreement:
'   By using this code, you agree to the following terms...
'   1) You may use this code in your own programs
'      (and may compile it into a program and distribute it in compiled
'      format for langauges that allow it) freely and with no charge.
'   2) You MAY NOT redistribute this code (for example to a web site, or code library, or book)
'      without written permission from the original author.
'      Failure to do so is a violation of copyright laws.
'   3) If you use this code for a commercial product I would like to have an email
'      stating the product that it is being used in.
'
'
'
' Usage:
'
'
'Option Explicit
'
'***************     Form Level Declaration
'Private WithEvents HO As cBinaryDBObject
'Private sFileName As String
'Private DocumentDB As Database
'
'
'
'Private Sub Command2_Click()
'
'    GetBinaryObject
'
'End Sub
'
'Private Sub Command3_Click()
'
'    SaveBinaryObject
'
'End Sub
'
'***************     Form Level Procedures
'Private Sub Form_Load()
'
'    'open the database
'    OpenDB DocumentDB, True
'    txtFileName = ""
'    LoadListBox
'
'End Sub
'
'Private Sub LoadListBox()
'Dim rs As Recordset
'Dim SQL As String
'
'    List1.Clear
'
'    ' open the table.
'    SQL = "SELECT * FROM tblFileObject"
'    Set rs = DocumentDB.OpenRecordset(SQL)
'    With rs
'        If .EOF And .BOF Then
'            .Close
'            GoTo LoadListBoxExit
'        End If
'        .MoveFirst
'        Do
'            List1.AddItem !ID
'            .MoveNext
'        Loop Until .EOF
'    End With
'
'LoadListBoxExit:
'
'End Sub
'
'Public Sub OpenDB(MyDB As Database, Optional OpenMDB As Boolean = True)
'
'    If OpenMDB Then
'        '/* Password protected database file */
'        Set MyDB = Workspaces(0).OpenDatabase(App.Path & "\db1.mdb", False, False, "")
'    Else
'        MyDB.Close
'        Set MyDB = Nothing
'    End If
'
'End Sub
'
'Private Sub HO_Error(ID As Long, Msg As String)
'
'    MsgBox ID & ":  " & Msg
'
'End Sub
'
'Private Sub HO_Status(ID As Long, Msg As String)
'
'    lblStatus = CStr(ID) & ":  " & Msg
'
'End Sub
'
'Private Sub GetBinaryObject()
'Dim FieldNames(1) As Variant           'names of the other fields to return
'Dim RD() As Variant                    'store for the returned data, not the binary field
'Dim FN As String                       'Binary file name to use as storage
'Dim i As Integer
'
'    If List1.SelCount > 0 Then
'        Set HO = New cBinaryDBObject       'create the new bd object
'
'        FieldNames(0) = "ID"               'return the ID field
'        FieldNames(1) = "FileName"         'return the filename
'
'        With HO
'            .KillFile = True                            'kill the filename if it exists
'            Set .DB = DocumentDB                        'pass the database
'            .ObjectKeyFieldName = "ID"                  'the key/index field is
'            .ObjectKey = List1.List(List1.ListIndex)    'the value to search for is
'            .ObjectFieldName = "OLEModule"              'name of the field that contains the binary file
'            .ObjectTableName = "tblFileObject"          'table that contains the binary files
'            .SubFieldNames = FieldNames                 'pass in the field names to return
'            .FileName = App.Path & "\picture.bmp"       'file name to use
'            .GetObject                                  'get the file from the database
'            .ReturnData RD()                            'return any aditional data
'            FN = .FileName                              'actual file name used - if default was used
'        End With
'        Set HO = Nothing
'
'        pPicture.Picture = LoadPicture(FN)
'
'        For i = 0 To UBound(RD)
'            Debug.Print RD(i)                      'print aditional info returned
'        Next
'
'    End If
'
'End Sub
'
'
'Private Sub SaveBinaryObject()
'Dim FieldNames(1) As Variant           'names of the other fields to return
'Dim FieldData(1) As Variant            'names of the other fields to return
'Dim RD() As Variant                    'store for the returned data, not the binary field
'Dim FN As String                       'Binary file name to use as storage
'Dim i As Integer
'
'    If sFileName = "" Then
'        Exit Sub
'    End If
'
'    Set HO = New cBinaryDBObject       'create the new bd object
'
'    FieldNames(0) = "ID"               'return the ID field
'    FieldNames(1) = "FileName"         'return the filename
'    FieldData(0) = Null                  'return the ID field
'    FieldData(1) = sFileName           'return the filename
'
'    With HO
'        .KillFile = False                       'kill the filename if it exists
'        Set .DB = DocumentDB                   'pass the database
'        .ObjectKeyFieldName = "ID"             'the key/index field is
'        .ObjectKey = -1                        'the value to search for is
'        .ObjectFieldName = "OLEModule"         'name of the field that contains the binary file
'        .ObjectTableName = "tblFileObject"     'table that contains the binary files
'        .SubFieldNames = FieldNames            'pass in the field names to return
'        .SubFieldData = FieldData
'        .FileName = sFileName                  'file name to use
'        .SaveObject                            'get the file from the database
'        .ReturnData RD()                       'return any aditional data
'        FN = .FileName                         'actual file name used - if default was used
'    End With
'    Set HO = Nothing
'
'    LoadListBox
'
'    For i = 0 To UBound(RD)
'        Debug.Print RD(i)                      'print aditional info returned
'    Next
'End Sub
'
'Private Sub cmdGetPicture_Click()
'On Error Resume Next
'
'    With CommonDialog1
'         .FileName = ""
'         .DialogTitle = "Extract to"
'         .CancelError = True
'         .Filter = "BMP (*.BMP)|*.bmp|All (*.*)|*.*"
'         .FileName = sFileName
'         .ShowSave
'         If Err.Number > 0 Then
'              Exit Sub
'         End If
'         sFileName = .FileName
'    End With
'    txtFileName = sFileName
'    pPicture.Picture = LoadPicture(sFileName)
'
'
'End Sub
'
'Private Sub Command1_Click()
'
'    pPicture.Picture = LoadPicture("")
'
'End Sub
'
'**************************************

Private Const defBlockSize = 10000

Private Enum StatusMessages
    StatusOK = 0
    ProgressStart = 1
    ProgressStatus = 2
    ProgressEnd = 3
    BadRecordID = 4
End Enum

Private Enum ErrorMessage
    NoError = 0
    GeneralError = 1
    NoSuchFileError = 2
    FileExistsError = 3
    NoSuchRecord = 4
End Enum

'**************************************
'Windows API/Global Declarations for
'**************************************
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Event Status(id As Long, Msg As String)
Public Event Error(id As Long, Msg As String)

'write only
Private mvarDB                  As DAO.Database  'local copy
Private mvarObjectKeyFieldName  As String    'local copy
Private mvarObjectFieldName     As String    'local copy
Private mvarObjectTableName     As String    'local copy
Private mvarBlockSize           As Long      'local copy
Private mvarKillFile            As Boolean   'local copy
Private mvarSubFieldNames       As Variant   'local copy
Private mvarObjectKey           As Long      'local copy

'read and write
Private mvarFileName            As String    'local copy
Private mvarSubFieldData        As Variant   'local copy

Public Property Let KillFile(ByVal vData As Boolean)
    mvarKillFile = vData
End Property

Public Property Let ObjectKeyFieldName(ByVal vData As String)
    mvarObjectKeyFieldName = vData
End Property

Public Property Let ObjectTableName(ByVal vData As String)
    mvarObjectTableName = vData
End Property

Public Property Let ObjectKey(ByVal vData As Long)
    mvarObjectKey = vData
End Property

Public Property Let SubFieldData(ByVal vData As Variant)
    mvarSubFieldData = vData
End Property

Public Property Let SubFieldNames(ByVal vData As Variant)
    mvarSubFieldNames = vData
End Property

Public Property Let ObjectFieldName(ByVal vData As String)
    mvarObjectFieldName = vData
End Property

Public Property Set DB(ByVal vData As DAO.Database)
    Set mvarDB = vData
End Property

Public Property Let BlockSize(ByVal vData As Long)
    mvarBlockSize = vData
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Sub GetObject()
On Error GoTo GetObjectErr

Dim rs As DAO.Recordset
Dim SQL As String
Dim lFileLength As Long
Dim lBlocks As Long
Dim lRemainder As Long
Dim lCurrBlock As Long
Dim FileNumber As Long
Dim aData() As Byte
Dim MaxArrayIndex As Long
Dim i As Integer

    'check if the file exists the file
    'cannot exist if it is to be extracted from the DB
    If Not FileExists(mvarFileName) Then
        
        RaiseStatus ProgressStart
        
        ' open the table.
        SQL = "SELECT * FROM " & mvarObjectTableName & " WHERE " _
                & mvarObjectKeyFieldName & " = " & mvarObjectKey
        
        Set rs = mvarDB.OpenRecordset(SQL)
        
        With rs
            'if do not have a record raise an error
            ' and status message
            'dont like to goto here
            If .EOF And .BOF Then
                RaiseStatus BadRecordID, CStr(mvarObjectKey)
                RaiseError NoSuchRecord, CStr(mvarObjectKey)
                .Close
                GoTo GetObjectExit
            End If
            
            ' If there's data in the field,
            ' save it to temp file.
            'lFileLength = .Fields(mvarObjectFieldName).FieldSize
            ReDim aData(lFileLength) As Byte
            If lFileLength > 0 Then
                
                'show progress
                RaiseStatus ProgressStatus
                
                ' Extract from database
                ' aData = .Fields(mvarObjectFieldName).GetChunk(0, lFileLength)
                '
                ' Save to temp file - the default temp file is used if no
                ' filename is entered - see class initialise
                FileNumber = FreeFile
                Open mvarFileName For Binary As #FileNumber
                Put #FileNumber, , aData
                Close #FileNumber
                '
                'check if we need to return any other field information
                If Not IsEmpty(mvarSubFieldNames) Then
                    'see if we need to return more than one field
                    MaxArrayIndex = Elements(mvarSubFieldNames)
                    If MaxArrayIndex = 0 Then
                        mvarSubFieldData = .Fields(mvarSubFieldNames)
                    Else
                        'return all fields required
                        ReDim mvarSubFieldData(MaxArrayIndex)
                        For i = 0 To MaxArrayIndex
                            mvarSubFieldData(i) = .Fields(mvarSubFieldNames(i))
                        Next
                    End If
                End If
            
            End If
            .Close
            RaiseStatus ProgressEnd
        End With
    Else
        RaiseError FileExistsError, mvarFileName
    End If


GetObjectExit:
    Exit Sub
    
GetObjectErr:
    RaiseError GeneralError, "In GetObject()"
    Resume GetObjectExit:

End Sub

Public Sub SaveObject()
On Error GoTo SaveObjectErr
Dim rs As DAO.Recordset
Dim SQL As String
Dim lFileLength As Long
Dim lBlocks As Long
Dim lRemainder As Long
Dim lCurrBlock As Long
Dim aData() As Byte
Dim FileNumber As Long
Dim MaxArrayIndex As Long
Dim i As Integer

    'see if the file exists - filename must have been entered
    'otherwise it would use a temp file name which is null
    If FileExists(mvarFileName) Then
        
        RaiseStatus ProgressStart
        
        ' open the table.
        SQL = "SELECT * FROM " & mvarObjectTableName & " WHERE " _
                & mvarObjectKeyFieldName & " = " & mvarObjectKey

        Set rs = mvarDB.OpenRecordset(SQL)

        With rs
            
            ' open the file for binary read
            FileNumber = FreeFile
            Open mvarFileName For Binary As #FileNumber
            lFileLength = LOF(FileNumber)
            RaiseStatus ProgressStatus, " - 0 of " & lFileLength
        
            'check that \we have somthing in the file
            If lFileLength > 0 Then
            
            'check if we need to edit or add the record
            If .EOF And .BOF Then
              .AddNew
            Else
              .Edit
            End If
            
                'see if we need to update aditional fields
                If Not IsEmpty(mvarSubFieldNames) Then
                    MaxArrayIndex = Elements(mvarSubFieldNames)
                    'see how many fields need to be updated
                    If MaxArrayIndex = 0 Then
                        .Fields(mvarSubFieldNames) = mvarSubFieldData
                    Else
                        For i = 0 To MaxArrayIndex
                            'update field only if it contains a value -
                            'need to remember that if we add a field that has
                            'a unique index or is an auto number field
                            'then we should not update it, so we need to set
                            'the field data to null in the calling routine
                            If Not IsNull(mvarSubFieldData(i)) Then
                                .Fields(mvarSubFieldNames(i)) = mvarSubFieldData(i)
                            End If
                        Next
                    End If
                End If
            
                'start the chunk culculations
                lBlocks = Int(lFileLength / mvarBlockSize)
                lRemainder = lFileLength Mod mvarBlockSize
                
                'add the first chunk
                If lRemainder > 0 Then
                    RaiseStatus ProgressStatus, lRemainder & " of " & lFileLength
                    ReDim aData(lRemainder - 1)
                    Get #FileNumber, , aData()
                    .Fields(mvarObjectFieldName).AppendChunk aData()
                End If
                
                'add the remaining chunks
                For lCurrBlock = 1 To lBlocks
                    ReDim aData(mvarBlockSize - 1)
                    Get #FileNumber, , aData()
                    .Fields(mvarObjectFieldName).AppendChunk aData()
                    RaiseStatus ProgressStatus, Format((lCurrBlock / lBlocks) * 100, "0.00") & "% " & .Fields(mvarObjectFieldName).FieldSize & " of " & lFileLength
                Next lCurrBlock
                
                'raise the end progress message
                RaiseStatus ProgressEnd, .Fields(mvarObjectFieldName).FieldSize & " of " & lFileLength
                Close FileNumber
            End If
            .Update
            
            'resample the field data values - brings back any autonumber fields, or fields that
            'have been updated
            If Not IsEmpty(mvarSubFieldNames) Then
                .Bookmark = .LastModified
                If MaxArrayIndex = 0 Then
                    mvarSubFieldData = .Fields(mvarSubFieldNames)
                Else
                    ReDim mvarSubFieldData(MaxArrayIndex)
                    For i = 0 To MaxArrayIndex
                        mvarSubFieldData(i) = .Fields(mvarSubFieldNames(i)).Value
                    Next
                End If
            End If
            
            .Close
        End With
    Else
        RaiseError NoSuchFileError, mvarFileName
    End If

SaveObjectExit:
    Exit Sub
    
SaveObjectErr:
    RaiseError GeneralError, "In SaveObject()"
    Resume SaveObjectExit:
    
End Sub

Private Sub RaiseStatus(id As StatusMessages, Optional Msg As String = "")
Dim StatusMsg As String
    
    Select Case id
        Case StatusOK
            StatusMsg = "OK "
        Case ProgressStart
            StatusMsg = "Starting Save"
        Case ProgressStatus
            StatusMsg = "Save progress "
        Case ProgressEnd
            StatusMsg = "Save progress 100% "
        Case BadRecordID
            StatusMsg = "Record not Valid: - Record ID = "
        Case Else
            StatusMsg = "No Such Status ID"
    End Select
    
    RaiseEvent Status(id, StatusMsg & Msg)

End Sub

Private Sub RaiseError(id As ErrorMessage, Optional Msg As String = "")
Dim Sep As String
Dim ErrorMsg As String

    Sep = ""
    If Msg = "" Then
        Sep = ": - "
    End If
    
    Select Case id
        Case NoError
            ErrorMsg = "No Error"
        Case GeneralError
            ErrorMsg = "Error In"
        Case NoSuchFileError
            ErrorMsg = "No Such File"
        Case FileExistsError
            ErrorMsg = "File Exists"
        Case Else
            ErrorMsg = "Unknown Error"
    End Select
    
    RaiseEvent Error(err, ErrorMsg & Sep & Msg)
    
End Sub

Private Function TempFileName() As String
Dim lpPathName As String
Dim lpPrefixString As String
Dim uUnique As Long
Dim lpTempFileName As String

    lpPathName = TempDir() & Chr(0)
    lpPrefixString = "DM" & Chr(0)
    uUnique = 0
    lpTempFileName = String(254, " ")
    
    GetTempFileName lpPathName, lpPrefixString, uUnique, lpTempFileName
    TempFileName = Left(Trim(lpTempFileName), Len(Trim(lpTempFileName)) - 1)

End Function

Private Function TempDir() As String
    
    TempDir = String(255, " ")
    GetTempPath 255, TempDir
    TempDir = Left(Trim(TempDir), Len(Trim(TempDir)) - 1)

End Function

Private Sub Class_Initialize()

    mvarBlockSize = defBlockSize
    mvarFileName = TempFileName
    
End Sub

Private Function FileExists(ByVal sFileName As String) As Boolean
' Determine if a file exists
Dim Exists As Boolean

    Exists = Not (Dir(sFileName) = "")
    
    If mvarKillFile And Exists Then
        Kill sFileName
        Exists = False
    End If
    
    FileExists = Exists
    
End Function

Private Function Elements(ary As Variant) As Long
Dim max As Long
On Error GoTo ElementsExit

    max = 0
    max = UBound(ary)
    
ElementsExit:
    Elements = max
    
End Function

Sub ReturnData(vData() As Variant)
Dim i As Integer
Dim MaxIndex As Integer

    'have not found a way to return a variant array other than this
    'need to research if there is another or better way
    If Not IsEmpty(mvarSubFieldData) Then
        MaxIndex = UBound(mvarSubFieldData)
        ReDim vData(MaxIndex)
        For i = 0 To MaxIndex
            vData(i) = mvarSubFieldData(i)
        Next
    End If
    
End Sub

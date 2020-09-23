Attribute VB_Name = "modADO"
''*****************************************************************
'' File Name:
'' Purpose:
'' Required Files:
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created:
'' Last Modified:
'' Modified By:
'' Credits: NONE, ALL CODES ARE CODED BY Philip V. Naparan
''*****************************************************************

'Some of the code are being modified in order to fit with my Point of Sale and Inventory System program.
'For more source code please visit my website at http://www.sourcecodester.com/

Option Explicit

Public Function OpenDB() As Integer
  Dim isOpen      As Boolean
  Dim ANS         As VbMsgBoxResult
  
  isOpen = False
  On Error GoTo errhandler
    
    
    
  Do Until isOpen = True
    CN.CursorLocation = adUseClient
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=" & DBPath & ";" _
    & "Persist Security Info=False;" _
    & "Jet OLEDB:Database Password=jaypee"
    
    isOpen = True
  Loop
  OpenDB = isOpen
    
  Exit Function
errhandler:
  ANS = MsgBox("Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, _
  vbCritical + vbRetryCancel)
  If ANS = vbCancel Then
    OpenDB = vbCancel
  ElseIf ANS = vbRetry Then
    OpenDB = vbRetry
  End If
End Function

Public Sub CloseDB()
    'Close the connection
    CN.Close
    Set CN = Nothing
End Sub

'Function that return the current index for a certain table
Public Function getIndex(ByVal srcTable As String) As Long
    On Error GoTo err
    Dim rs As New Recordset
    Dim RI As Long
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM TBL_GENERATOR WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
    
    RI = rs.Fields("NextNo")
    rs.Fields("NextNo") = RI + 1
    rs.Update
    
    getIndex = RI
    
    srcTable = ""
    RI = 0
    Set rs = Nothing
    Exit Function
err:
        ''Error when incounter a null value
        If err.Number = 94 Then
            getIndex = 1
            Resume Next
        Else
            MsgBox err.Description
        End If
End Function

'Function used to get the sum  of fields
Public Function getSumOfFields(ByVal sTable As String, ByVal sField As String, ByRef sCN As ADODB.Connection, Optional inclField As String, Optional sCondition As String) As Double
    On Error GoTo err
    Dim rs As New ADODB.Recordset

    rs.CursorLocation = adUseClient
    If sCondition <> "" Then sCondition = " GROUP BY " & inclField & " HAVING(" & sCondition & ")"
    If inclField <> "" Then inclField = "," & inclField
    rs.Open "SELECT Sum(" & sTable & "." & sField & ") AS fTotal" & inclField & " FROM " & sTable & sCondition, sCN, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            getSumOfFields = getSumOfFields + rs.Fields("fTotal")
            rs.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If
    
    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function

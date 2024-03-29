Attribute VB_Name = "modProcedure"
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

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub LoadForm(ByRef srcForm As Form)
    srcForm.show
    srcForm.WindowState = vbMaximized
    srcForm.SetFocus
End Sub

'Used to locate the key in opened form
Public Sub HighlightInWin(ByVal srcKey As String)
    With MAIN.lvWin
        If .ListItems.Count > 0 Then
            If .SelectedItem.Key <> srcKey Then
                Dim c As Integer
                For c = 1 To .ListItems.Count
                    If .ListItems(c).Key = srcKey Then
                        .ListItems(c).Selected = True
                        .ListItems(c).EnsureVisible
                        Exit For
                    End If
                Next c
            End If
        End If
    End With
End Sub

'Procedure used to custom move the recordset cursor
Public Sub customMove(ByRef sRS As Recordset, ByVal isNum As Boolean, ByVal findStr As String, ByVal sField As String)
    If sRS.RecordCount < 1 Then Exit Sub
    Dim old_pos As Long
    sRS.MoveFirst
    old_pos = sRS.AbsolutePosition
    If isNum = True Then
        sRS.Find sField & " = " & findStr
    Else
        sRS.Find sField & " = '" & findStr & "'"
    End If
    If sRS.EOF Then sRS.AbsolutePosition = old_pos
    old_pos = 0
End Sub
'This code is also available in .NET version with ADO.NET
'Procedure used to fill list view
Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)
    Dim X As Variant
    Dim i As Byte
    On Error Resume Next
    sListView.ListItems.Clear
    If sRecordSource.RecordCount < 1 Then Exit Sub
    sRecordSource.MoveFirst
    Do While Not sRecordSource.EOF
        If with_num = True Then
            Set X = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, sNumIco, sNumIco)
        Else
            Set X = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), sNumIco, sNumIco)
        End If
            If srcHiddenField <> "" Then X.Tag = sRecordSource.Fields(srcHiddenField)
            For i = 1 To sNumOfFields - 1
                If show_first_rec = True Then
                    If with_num = True Then
                        If sRecordSource.Fields(CInt(i) - 1).Type = adDouble Then
                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i) - 1))
                        Else
                            X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) - 1))
                        End If
                    Else
                        If sRecordSource.Fields(CInt(i)).Type = adDouble Then
                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i)))
                        Else
                            X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i)))
                        End If
                    End If
                Else
                    X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) + 1))
                End If
            Next i
        sRecordSource.MoveNext
    Loop
    i = 0
    Set X = Nothing
End Sub

'Procedure used to promp unexpected errors
Public Sub Prompt_Err(ByVal sError As ErrObject, ByVal ModuleName As String, ByVal OccurIn As String)
    MsgBox "Error From: " & ModuleName & vbNewLine & _
           "Occur In: " & OccurIn & vbNewLine & _
           "Error Number: " & sError.Number & vbNewLine & _
           "Description: " & sError.Description, vbCritical, "Application Error"
    'Save the error log (The save error log will be display later on in the program)
    Open App.Path & "\Error.log" For Append As #1
        Print #1, Format(Date, "MMM-dd-yyyy") & "~~~~~" & Time & "~~~~~" & sError.Number & "~~~~~" & sError.Description & "~~~~~" & ModuleName & "~~~~~" & OccurIn
    Close #1
End Sub

'Procedure used to delete record with SQL
Public Sub DelRecwSQL(ByVal sTable As String, ByVal sField As String, ByVal sString As String, ByVal isNumber As Boolean, ByVal snum As Long)
    If isNumber = True Then
        CN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " =" & snum
    Else
        CN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " ='" & sString & "'"
    End If
End Sub

'Procedure used to fill the listview in paging method
Public Sub pageFillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal pos_start As Long, ByVal pos_end As Long, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional match_field As String, Optional match_str As String, Optional match_ico As Byte, Optional srcHiddenField As String)

    Dim X As ListItem
    Dim i As Byte, c As Long, old_pt As Long
    sListView.ListItems.Clear
    If sRecordSource.RecordCount < 1 Then Exit Sub
    sRecordSource.AbsolutePosition = pos_start
    On Error Resume Next
    old_pt = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    DoEvents
    Do
        If match_field = "" Then
            If with_num = True Then
                Set X = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, sNumIco, sNumIco)
            Else
                Set X = sListView.ListItems.Add(, , "" & FormatRS(sRecordSource.Fields(0)), sNumIco, sNumIco)
            End If
        Else
            If sRecordSource.Fields(match_field) = match_str Then
                If with_num = True Then
                    Set X = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, match_ico, match_ico)
                Else
                    Set X = sListView.ListItems.Add(, , "" & FormatRS(sRecordSource.Fields(0)), match_ico, match_ico)
                End If
            Else
                If with_num = True Then
                    Set X = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, sNumIco, sNumIco)
                Else
                    Set X = sListView.ListItems.Add(, , "" & FormatRS(sRecordSource.Fields(0)), sNumIco, sNumIco)
                End If
            End If
        End If
            If srcHiddenField <> "" Then
                X.Tag = sRecordSource.Fields(srcHiddenField) & "*~~~~~*" & c + pos_start
              Else
                X.Tag = c + pos_start
            End If
            For i = 1 To sNumOfFields - 1
                If show_first_rec = True Then
                    If with_num = True Then
                             X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) - 1))
                    Else
                            X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i)))
                    End If
                Else
                        X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) + 1))
                End If
            Next i
            
        If sRecordSource.AbsolutePosition >= pos_end Then
            Exit Do
        Else
            sRecordSource.MoveNext
            c = c + 1
        End If
    Loop
    Screen.MousePointer = old_pt
    i = 0: c = 0: old_pt = 0
    Set X = Nothing
End Sub

'Procedure used to highlight text when focus
Public Sub HLText(ByRef sText)
    On Error Resume Next
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub

'Procedure used to bind data combo
Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, CN, adOpenStatic, adLockOptimistic
    
    With srcDC
        .ListField = srcBindField
        .BoundColumn = srcColBound
        Set .RowSource = rs
        'Display the first record
        If ShowFirstRec = True Then
            If Not rs.RecordCount < 1 Then
                .BoundText = rs.Fields(srcColBound)
                .Tag = rs.RecordCount & "*~~~~~*" & rs.Fields(srcColBound)
            Else
                .Tag = "0*~~~~~*0"
            End If
        End If
    End With
    Set rs = Nothing
End Sub

'Procedure used to bind data list
Public Sub bind_dl(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDL As DataList, Optional srcColBound As String, Optional ShowFirstRec As Boolean)
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, CN, adOpenStatic, adLockOptimistic
    
    With srcDL
        .ListField = srcBindField
        .BoundColumn = srcColBound
        Set .RowSource = rs
        'Display the first record
        If ShowFirstRec = True Then
            If Not rs.RecordCount < 1 Then
                .BoundText = rs.Fields(srcColBound)
                .Tag = rs.RecordCount & "*~~~~~*" & rs.Fields(srcColBound)
            Else
                .Tag = "0*~~~~~*0"
            End If
        End If
    End With
    Set rs = Nothing
End Sub

'Procedure used to clear the text content
Public Sub clearText(ByRef sForm As Form)
    Dim CONTROL As CONTROL
    For Each CONTROL In sForm.Controls
        If (TypeOf CONTROL Is TextBox) Then CONTROL = vbNullString
    Next CONTROL
    Set CONTROL = Nothing
End Sub

'Procedure used to clear the text content
Public Sub LockInput(ByRef sForm As Form, ByVal bolLock As Boolean, Optional bolTabStop As Boolean)
    On Error Resume Next
    Dim CONTROL As CONTROL
    For Each CONTROL In sForm.Controls
       CONTROL.Locked = bolLock
    Next CONTROL
    Set CONTROL = Nothing
End Sub

'Procedure that will change the value at once
Public Sub ChangeValue(ByRef srcCN As Connection, ByVal srcTable As String, ByVal srcField As String, ByVal srcValue As String, Optional isNumber As Boolean, Optional srcCondition As String)
    If srcCondition <> vbNullString Then srcCondition = " " & srcCondition
    If isNumber = True Then
        srcCN.Execute "UPDATE " & srcTable & " SET " & srcField & " =" & srcValue & " " & srcCondition
    Else
        srcCN.Execute "UPDATE " & srcTable & " SET " & srcField & " ='" & srcValue & "'" & " " & srcCondition
    End If
End Sub

Public Sub FillFlex(ByRef srcFlex As MSHFlexGrid, ByVal srcSQL As String, ByVal srcNoOfCol As Integer)
    Dim rs As New Recordset
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, CN, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 1 Then Exit Sub
    rs.MoveFirst
    Dim i As Long, c As Long
    srcFlex.Rows = (srcFlex.Rows + rs.RecordCount) - 1
        For i = 1 To rs.RecordCount
            For c = 0 To srcNoOfCol - 1
                srcFlex.TextMatrix(i, c) = rs.Fields(c)
            Next c
            rs.MoveNext
        Next i
    i = 0
    c = 0
    Set rs = Nothing
End Sub

'Procedure used to search in listview
Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
    Dim tmp_listtview As ListItem
    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
End Sub

'Procedure used to center form
Public Sub centerForm(ByRef sForm As Form, ByVal sHeight As Integer, ByVal sWidth As Integer)
    sForm.Move (sWidth - sForm.Width) / 2, (sHeight - sForm.Height) / 2
End Sub
'Procedure used to center object horizontal
Public Sub center_obj_horizontal(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Left = (sParentObj - sMoveObj.Width) / 2
End Sub
'Procedure used to center vertical
Public Sub center_obj_vertical(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Top = (sParentObj.Height - sMoveObj.Height) / 2
End Sub

Function GetINI(strMain As String, strSub As String) As String
  Dim strBuffer As String
  Dim lngLen As Long
  Dim lngRet As Long
  
  strBuffer = Space(100)
  lngLen = Len(strBuffer)
  lngRet = GetPrivateProfileString(strMain, strSub, vbNullString, strBuffer, lngLen, App.Path & "\config.txt")
  GetINI = Left(strBuffer, lngRet)
End Function

Public Sub SetINI(strMain As String, strSub As String, strvalue As String)
  WritePrivateProfileString strMain, strSub, strvalue, App.Path & "\config.txt"
End Sub


'allow only numbers to be accepted on the specified object obj
Public Function AllowOnlyNumbers(KeyAscii As Integer, obj As CONTROL) As Integer
  If ((KeyAscii <> 8) And (KeyAscii <> vbKeyDelete) And _
  (KeyAscii <> 46)) And ((KeyAscii < 48 Or KeyAscii > 57)) Then
    AllowOnlyNumbers = 0
  Else
    If KeyAscii = 46 Then
      If InStr(obj.Text, ".") Then
        KeyAscii = 0
        Exit Function
      End If
    End If
    AllowOnlyNumbers = KeyAscii
  End If
End Function



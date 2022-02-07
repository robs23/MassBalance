Attribute VB_Name = "ZFOR"
Public Sub formatZFOR()

With ThisWorkbook.Sheets("ZFOR Comp")
    .Range("A1:A2").Merge
    .Range("B1:b2").Merge
    .Range("C1:c2").Merge
    .Range("D1:D2").Merge
    .Range("E1:G1").Merge
    .Range("H1:J1").Merge
    .Range("K1:M1").Merge
    .Range("N1:N2").Merge
    .Range("A1") = "Order No"
    .Range("B1") = "ZFOR"
    .Range("C1") = "Description"
    .Range("D1") = "Beans?"
    .Range("E1") = "Green Coffee [kg]"
    .Range("H1") = "Roasted Coffee [kg]"
    .Range("K1") = "Ground Coffee [kg]"
    .Range("E2") = "SCADA"
    .Range("F2") = "SAP"
    .Range("G2") = "SCADA vs SAP"
    .Range("H2") = "SCADA"
    .Range("I2") = "MES"
    .Range("J2") = "SCADA vs MES"
    .Range("K2") = "SAP"
    .Range("L2") = "MES"
    .Range("M2") = "SAP vs MES"
    .Range("N1") = "Comment"
    .Range("A1:N2").Font.Bold = True
    .Range("A1:N2").HorizontalAlignment = xlCenter
End With

End Sub
Public Sub compareZfors(control As IRibbonControl)
ThisWorkbook.Sheets("ZFOR Comp").Cells.clear
formatZFOR
transferSCADA
transferMES
bringBeans
importSAP
finishMe
transferResults
End Sub


Public Sub transferSCADA()
Dim i As Integer
Dim lastRow As Long
Dim sht As Worksheet
Dim rng As Range

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets("SCADA")

lastRow = sht.Range("J:J").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
Set rng = sht.Range("J2:L" & lastRow)
rng.Copy ThisWorkbook.Sheets("ZFOR Comp").Range("A3:C" & lastRow + 1)
Set rng = sht.Range("O2:O" & lastRow)
rng.Copy ThisWorkbook.Sheets("ZFOR Comp").Range("E3:E" & lastRow + 1)
Set rng = sht.Range("P2:P" & lastRow)
rng.Copy ThisWorkbook.Sheets("ZFOR Comp").Range("H3:H" & lastRow + 1)

exit_here:
Set sht = Nothing
Set rng = Nothing
Exit Sub

err_trap:
MsgBox "Error in transferSCADA. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Sub transferMES()
Dim rng As Range
Dim rng2 As Range
Dim c As Range
Dim mes As Worksheet
Dim dest As Worksheet
Dim lastRow As Long
Dim lastRow2 As Long
Dim theRow As Long

On Error GoTo err_trap

Set mes = ThisWorkbook.Sheets("MES")
Set dest = ThisWorkbook.Sheets("ZFOR Comp")

lastRow = mes.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
lastRow2 = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = mes.Range("A3:A" & lastRow)
Set rng2 = dest.Range("A3:A" & lastRow2)
    
For Each c In rng
    If WorksheetFunction.CountIf(rng2, c.Value) = 0 Then
        dest.Cells(Rows.Count, 1).End(xlUp)(2) = c.Value
    End If
Next c

lastRow = mes.Range("G:G").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
lastRow2 = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = mes.Range("G3:G" & lastRow)
Set rng2 = dest.Range("A3:A" & lastRow2)
    
For Each c In rng
    If WorksheetFunction.CountIf(rng2, c.Value) = 0 Then
        dest.Cells(Rows.Count, 1).End(xlUp)(2) = c.Value
    End If
Next c

' the other way around
lastRow = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
lastRow2 = mes.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = dest.Range("A3:A" & lastRow)
Set rng2 = mes.Range("A3:A" & lastRow2)

For Each c In rng
    dest.Range("M" & c.row).Formula = "=K" & c.row & "-L" & c.row
    dest.Range("G" & c.row).Formula = "=E" & c.row & "-F" & c.row
    dest.Range("J" & c.row).Formula = "=H" & c.row & "-I" & c.row
    dest.Range("M" & c.row).NumberFormat = "0.0"
    dest.Range("J" & c.row).NumberFormat = "0.0"
    dest.Range("G" & c.row).NumberFormat = "0.0"
    theRow = 0
    theRow = rng2.Find(c.Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
    If Not theRow = 0 Then
        dest.Range("I" & c.row) = mes.Range("D" & theRow)
    End If
Next c

Set rng2 = mes.Range("G3:G" & lastRow2)

For Each c In rng
    theRow = 0
    theRow = rng2.Find(c.Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
    If Not theRow = 0 Then
        dest.Range("L" & c.row) = mes.Range("J" & theRow)
    End If
Next c

exit_here:
Set rng = Nothing
Set rng2 = Nothing
Set dest = Nothing
Set mes = Nothing
Exit Sub

err_trap:
If Err.Number = 91 Then
    Resume Next
Else
    MsgBox "Error in transferMES. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub


Public Sub importSAP()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim oCol As String
Dim surCol As String
Dim mielCol As String
Dim aCol As String
Dim mCol As String
Dim kgCol As String
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim prop As String
'Dim zlec As clsZlecenie
Dim ver As String
Dim theCol As Long
Dim rng As Range
Dim lastRow As Long


On Error GoTo err_trap

prop = "rozliczenie file"

If propertyExists(prop) And ThisWorkbook.CustomDocumentProperties(prop) <> "" Then
    If ThisWorkbook.CustomDocumentProperties("import path") <> "" Then
        If FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls"
        ElseIf FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx"
        Else
            MsgBox "Source file """ & ThisWorkbook.CustomDocumentProperties(prop) & """ could not be found in " & ThisWorkbook.CustomDocumentProperties("import path") & "\ . Check in settings if both file name and path are correct.", vbOKOnly + vbExclamation, "Error"
        End If
        
        If path <> "" Then
        If Right(path, 1) = "s" Then
            ver = "8"
        Else
            ver = "12"
        End If
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=0;"";"
        cnn.Open conStr
            For n = 0 To 1
                oCol = ""
                surCol = ""
                mielCol = ""
                If month(ThisWorkbook.CustomDocumentProperties("packingFrom")) + n >= 10 Then
                    Worksheet = "Period " & month(ThisWorkbook.CustomDocumentProperties("packingFrom")) + n & "_FY" & ThisWorkbook.CustomDocumentProperties("yearLoaded") - 2000
                Else
                    Worksheet = "Period 0" & month(ThisWorkbook.CustomDocumentProperties("packingFrom")) + n & "_FY" & ThisWorkbook.CustomDocumentProperties("yearLoaded") - 2000
                End If
                Set rs = importExcelData(path, Worksheet, 1)
                If Not rs.EOF Then
                    rs.MoveFirst
                    Do Until rs.EOF
                        If Len(oCol) = 0 And Len(surCol) = 0 And Len(mielCol) = 0 Then
                            For i = 0 To rs.Fields.Count - 1
    
                                    If InStr(1, rs.Fields(i), "Numer zlecenia", vbTextCompare) > 0 Then
                                    oCol = rs.Fields(i).Name
                                ElseIf InStr(1, rs.Fields(i), "Kawa surowa SAP", vbTextCompare) > 0 Then
                                    surCol = rs.Fields(i).Name
                                ElseIf InStr(1, rs.Fields(i), "prazonej SAP", vbTextCompare) > 0 Then
                                    mielCol = rs.Fields(i).Name
                                End If
                            Next i
                        Else
                            rs.Close
                            Set rs = Nothing
                            Exit Do
                        End If
                        rs.MoveNext
                    Loop

                    'by order number
                    strSQL = "SELECT * FROM [" & Worksheet & "$]  WHERE " & oCol & " > 100;"
                    'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
                    'strSQL = "SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100;"

                    Set rs = New ADODB.Recordset
                    rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
                    If Not rs.EOF Then
                        rs.MoveFirst
                        With ThisWorkbook.Sheets("ZFOR Comp")
                            lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
                            Set rng = .Range("A3:A" & lastRow)
                            Do Until rs.EOF
                                theRow = 0
                                theRow = rng.Find(rs.Fields(oCol).Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
                                If theRow > 0 Then
                                    .Range("F" & theRow) = rs.Fields(surCol).Value
                                    If .Range("D" & theRow) = 0 Then .Range("K" & theRow) = rs.Fields(mielCol).Value
                                End If
                                rs.MoveNext
                            Loop
                        End With
                    End If
                    rs.Close
                    Set rs = Nothing
    '                        If Not orderExists(rs.Fields(pCol).value) Then
    '                            'add new order
    '                            Set zlec = New clsZlecenie
    '                            rs.filter = pCol & " = " & rs.Fields(pCol).value
    '                            With zlec
    '                                .index = rs.Fields(pCol).value
    '                                .Name = rs.Fields(nCol).value
    '                                .Order = rs.Fields(oCol).value
    '                            End With
    '                            zlecenia.Add zlec, CStr(rs.Fields(pCol))
    '                        End If
                End If
            Next n
        End If
    End If
End If

exit_here:
Set rng = Nothing
Set rs = Nothing
Exit Sub

err_trap:
If Err.Number = 91 Then
    Resume Next
Else
    MsgBox "Error in importSAP. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Public Sub finishMe()
Dim rng As Range
Dim lastRow As Long

On Error GoTo err_trap

With ThisWorkbook.Sheets("ZFOR Comp")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A1:N2")
    rng.Interior.ColorIndex = 15
    Set rng = .Range("A1:N" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
End With

exit_here:
Set rng = Nothing
Exit Sub

err_trap:
MsgBox "Error in finishMe. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here
End Sub

Public Sub bringBeans()
Dim rng As Range
Dim lastRow As Long
Dim sSql As String
Dim prodStr As String
Dim zfors() As Variant
Dim i As Integer
Dim c As Range
Dim found As Boolean
Dim rs As ADODB.Recordset
Dim res As VbMsgBoxResult

Set conn = New ADODB.Connection
conn.Open ConnectionString
conn.CommandTimeout = 90

With ThisWorkbook.Sheets("ZFOR Comp")
    lastRow = .Range("B:B").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("B3:B" & lastRow)
    For Each c In rng
        prodStr = prodStr & c.Value & ","
    Next c
    prodStr = Left(prodStr, Len(prodStr) - 1)
    
    sSql = "SELECT zfinIndex, zfinName, [beans?] FROM tbZfinProperties RIGHT JOIN tbZfin on tbZfin.zfinId = tbZfinProperties.zfinId WHERE tbZfin.zfinIndex IN (" & prodStr & ") AND tbZfin.zfinType = 'zfor';"
    Set rs = New ADODB.Recordset
    rs.Open sSql, conn, adOpenStatic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        rs.MoveFirst
        ReDim Preserve zfors(1, 0) As Variant
        zfors(0, 0) = rs.Fields("zfinIndex")
        If IsNull(rs.Fields("beans?")) Then
            res = MsgBox("I can't tell if product " & rs.Fields("zfinIndex") & " " & rs.Fields("zfinName") & " is BEAN or GROUND product. If it's bean product, click YES, otherwise click NO. You can skip this with CANCEL", vbYesNoCancel + vbExclamation, "Create new ZFOR")
            If res = vbYes Then
                'we've got beans
                insertZfinZfor rs.Fields("zfinIndex"), rs.Fields("zfinName"), True, "zfor"
                zfors(1, UBound(zfors, 2)) = True
            ElseIf res = vbNo Then
                'we've got ground
                insertZfinZfor rs.Fields("zfinIndex"), rs.Fields("zfinName"), False, "zfor"
                zfors(1, UBound(zfors, 2)) = False
            End If
        Else
            zfors(1, 0) = Abs(rs.Fields("beans?"))
        End If
        Do Until rs.EOF
            rs.MoveNext
            If Not rs.EOF Then
                ReDim Preserve zfors(1, UBound(zfors, 2) + 1) As Variant
                zfors(0, UBound(zfors, 2)) = rs.Fields("zfinIndex")
                If IsNull(rs.Fields("beans?")) Then
                    res = MsgBox("I can't tell if product " & rs.Fields("zfinIndex") & " " & rs.Fields("zfinName") & " is BEAN or GROUND product. If it's bean product, click YES, otherwise click NO. You can skip this with CANCEL", vbYesNoCancel + vbExclamation, "Create new ZFOR")
                    If res = vbYes Then
                        'we've got beans
                        insertZfinZfor rs.Fields("zfinIndex"), rs.Fields("zfinName"), True, "zfor"
                        zfors(1, UBound(zfors, 2)) = True
                    ElseIf res = vbNo Then
                        'we've got ground
                        insertZfinZfor rs.Fields("zfinIndex"), rs.Fields("zfinName"), False, "zfor"
                        zfors(1, UBound(zfors, 2)) = False
                    End If
                Else
                    zfors(1, UBound(zfors, 2)) = Abs(rs.Fields("beans?"))
                End If
            End If
        Loop
        For Each c In rng
            found = False
            For i = LBound(zfors, 2) To UBound(zfors, 2)
                If c.Value = zfors(0, i) Then
                    .Range("D" & c.row) = zfors(1, i)
                    found = True
                    Exit For
                End If
            Next i
            If found = False Then
                'new ZFOR has been found. Add it to db
                res = MsgBox("New ZFOR has been found! It's " & c.Value & " " & c.Offset(0, 1) & ". I'll add it to database for later use. Please take a moment to determine: is it ""BEAN"" product?", vbYesNoCancel + vbExclamation, "Create new ZFOR")
                If res = vbYes Then
                    'we've got beans
                    insertZfinZfor CLng(c.Value), c.Offset(0, 1), True, "zfor"
                    c.Offset(0, 2) = 1
                ElseIf res = vbNo Then
                    'we've got ground
                    insertZfinZfor CLng(c.Value), c.Offset(0, 1), False, "zfor"
                    c.Offset(0, 2) = 0
                End If
            End If
        Next c
    End If
    rs.Close
    Set rs = Nothing
End With
conn.Close
Set conn = Nothing

End Sub

Public Sub insertZfinZfor(ind As Long, nam As String, bean As Boolean, theType As String)
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim Id As Integer

On Error GoTo err_trap

Set conn = New ADODB.Connection
conn.Open ConnectionString
conn.CommandTimeout = 90
Set rs = New ADODB.Recordset
rs.Open "SELECT zfinId FROM tbZfin WHERE zfinIndex = " & ind, conn, adOpenKeyset, adLockOptimistic
If rs.EOF Then
    rs.Close
    Set rs = Nothing
    sSql = "INSERT INTO tbZfin (zfinIndex, zfinName, zfinType) VALUES (" & ind & ", '" & nam & "', '" & theType & "'); SELECT SCOPE_IDENTITY()"
    Set rs = New ADODB.Recordset
    rs.Open sSql, conn, adOpenKeyset, adLockOptimistic
    'Set rs = rs.NextRecordset()
    Id = rs.Fields(0).Value
    rs.Close
Else
    rs.MoveFirst
    Id = rs.Fields("zfinId")
    rs.Close
End If

Set rs = Nothing
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM tbZfinProperties WHERE zfinId = " & Id, conn, adOpenKeyset, adLockOptimistic
If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    sSql = "INSERT INTO tbZfinProperties (zfinId, [beans?]) VALUES (" & Id & ", " & CInt(bean) & ");"
    rs.Open sSql, conn, adOpenKeyset, adLockOptimistic
Else
    rs.MoveFirst
    rs.Fields("beans?") = bean
    rs.Update
End If

exit_here:
If rs.State = 1 Then
    rs.Close
End If
If conn.State = 1 Then
    conn.Close
End If

Set rs = Nothing
Set conn = Nothing
Exit Sub

err_trap:
MsgBox "Error in insertZfinZfor. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Sub transferResults()
Dim rng As Range
Dim lastRow As Long


With ThisWorkbook.Sheets("ZFOR Comp")
    lastRow = .Range("B:B").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("D3:D" & lastRow)
End With
With ThisWorkbook.Sheets("BM")
    If .Range("F4") = "" Then .Range("F4") = WorksheetFunction.SumIf(rng, 1, ThisWorkbook.Sheets("ZFOR Comp").Range("E3:E" & lastRow))
    If .Range("F6") = "" Then .Range("F6") = WorksheetFunction.SumIf(rng, 0, ThisWorkbook.Sheets("ZFOR Comp").Range("E3:E" & lastRow))
    If .Range("O5") = "" Then .Range("O5") = WorksheetFunction.SumIf(rng, 1, ThisWorkbook.Sheets("ZFOR Comp").Range("H3:H" & lastRow))
    If .Range("O11") = "" Then .Range("O11") = WorksheetFunction.SumIf(rng, 0, ThisWorkbook.Sheets("ZFOR Comp").Range("H3:H" & lastRow))
End With

Set rng = Nothing
End Sub

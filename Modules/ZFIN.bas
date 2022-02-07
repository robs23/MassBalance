Attribute VB_Name = "ZFIN"
Public Sub compareZfins(control As IRibbonControl)
ThisWorkbook.Sheets("ZFIN Comp").Cells.clear
formatZFIN
transferPW
zfinMES
transferIshida
importTIPTOP
importPlan
finishIt
End Sub

Public Sub formatZFIN()
With ThisWorkbook.Sheets("ZFIN Comp")
    .Cells.clear
    .Range("A1:A2").Merge
    .Range("B1:b2").Merge
    .Range("C1:c2").Merge
    .Range("D1:H1").Merge
    .Range("I1:I2").Merge
    .Range("J1:J2").Merge
    .Range("K1:K2").Merge
    .Range("A1") = "ZFIN"
    .Range("B1") = "Description"
    .Range("C1") = "Beans?"
    .Range("D1") = "Production [kg]"
    .Range("D2") = "MES"
    .Range("E2") = "PW"
    .Range("F2") = "ISHIDA"
    .Range("G2") = "SAP"
    .Range("H2") = "Plan"
    .Range("I1") = "PW vs ISHIDA"
    .Range("J1") = "Comment"
    .Range("A1:J2").Font.Bold = True
    .Range("A1:J2").HorizontalAlignment = xlCenter
    
End With
End Sub


Public Sub zfinMES()
Dim rng As Range
Dim rng2 As Range
Dim c As Range
Dim mes As Worksheet
Dim dest As Worksheet
Dim lastRow As Long
Dim lastRow2 As Long
Dim theRow As Long
Dim added() As Long
Dim bool As Boolean
Dim i As Integer

On Error GoTo err_trap

Set mes = ThisWorkbook.Sheets("MES")
Set dest = ThisWorkbook.Sheets("ZFIN Comp")

lastRow = mes.Range("N:N").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
lastRow2 = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = mes.Range("N3:N" & lastRow)
Set rng2 = dest.Range("A3:A" & lastRow2)
    
For Each c In rng
    bool = True
    If WorksheetFunction.CountIf(rng2, c.Value) = 0 Then
        If Not isArrayEmpty(added) Then
            For i = LBound(added) To UBound(added)
                If added(i) = c.Value Then
                    bool = False
                    Exit For
                End If
            Next i
        End If
        If bool Then
            dest.Cells(Rows.Count, 1).End(xlUp)(2) = c.Value
            If isArrayEmpty(added) Then
                ReDim added(0) As Long
                added(0) = c.Value
            Else
                ReDim Preserve added(UBound(added) + 1) As Long
                added(UBound(added)) = c.Value
            End If
        End If
    End If
Next c

' the other way around
lastRow = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
lastRow2 = mes.Range("N:N").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = dest.Range("A3:A" & lastRow)
Set rng2 = mes.Range("N3:N" & lastRow2)

For Each c In rng
'    dest.Range("M" & c.row).Formula = "=K" & c.row & "-L" & c.row
'    dest.Range("G" & c.row).Formula = "=E" & c.row & "-F" & c.row
'    dest.Range("J" & c.row).Formula = "=H" & c.row & "-I" & c.row
'    dest.Range("M" & c.row).NumberFormat = "0.0"
'    dest.Range("J" & c.row).NumberFormat = "0.0"
'    dest.Range("G" & c.row).NumberFormat = "0.0"
    dest.Range("D" & c.row) = WorksheetFunction.SumIf(rng2, c.Value, mes.Range("P3:P" & lastRow2))
Next c

'Set rng2 = mes.Range("G3:G" & lastRow2)
'
'For Each c In rng
'    theRow = 0
'    theRow = rng2.Find(c.value, searchorder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole).row
'    If Not theRow = 0 Then
'        dest.Range("L" & c.row) = mes.Range("J" & theRow)
'    End If
'Next c

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
    MsgBox "Error in zfinMES. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If


End Sub

Public Sub importISHIDA(control As IRibbonControl)
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim linCol As String
Dim pCol As String
Dim grCol As String
Dim aCol As String
Dim mCol As String
Dim avgCol As String
Dim fromCol As String
Dim toCol As String
Dim wasCol As String
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim prop As String
'Dim zlec As clsZlecenie
Dim ver As String
Dim allSheets As Variant
Dim lastRow As Long
Dim rng As Range

prop = "ishida file"


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
            Set rs = importExcelData(path, , 1)
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If Len(linCol) = 0 Or Len(pCol) = 0 Or Len(grCol) = 0 Or Len(aCol) = 0 Or Len(mCol) = 0 Or Len(avgCol) = 0 Or Len(fromCol) = 0 Or Len(toCol) = 0 Or Len(wasCol) = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            Select Case rs.Fields(i)
                            Case Is = "NAZWA"
                                If pCol = "" Then pCol = rs.Fields(i).Name
                            Case Is = "NOMINALNA"
                                grCol = rs.Fields(i).Name
                            Case Is = "ZAAKCEPTOWANE"
                                aCol = rs.Fields(i).Name
                            Case Is = "CAŁKOWITA"
                                mCol = rs.Fields(i).Name
                            Case Is = "S1"
                                avgCol = rs.Fields(i).Name
                            Case Is = "PRODUKCJI"
                                fromCol = rs.Fields(i).Name
                            Case Is = "POMIARU"
                                toCol = rs.Fields(i).Name
                            Case Is = "STRATY"
                                wasCol = rs.Fields(i).Name
                            End Select
                            If InStr(1, rs.Fields(i).Value, "Linia", vbTextCompare) > 0 Then
                                linCol = rs.Fields(i).Name
                            End If
                        Next i
                    Else
                        rs.Close
                        Set rs = Nothing
                        Exit Do
                    End If
                    rs.MoveNext
                Loop
                If Right(path, 1) = "s" Then
                    ver = "8"
                Else
                    ver = "12"
                End If
                allSheets = getExcelSheetName(path)
                Worksheet = allSheets(0)
                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=1;"";"
                'by order number
'                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sub." & nCol & ", sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE " & grCol & " > 20) sub GROUP BY " & oCol & ", " & pCol & ", " & nCol & ";"
                'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
                strSQL = "SELECT * FROM [" & Worksheet & "] WHERE " & grCol & " is not null;"
                'strSQL = "SELECT * FROM [" & Worksheet & "];"
                cnn.Open conStr
                Set rs = New ADODB.Recordset
                rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs.EOF Then
                    rs.filter = pCol & " >= 100 AND " & pCol & " <> 'NAZWA' AND " & pCol & " <> 'PRODUKTU'"
                    If Not rs.EOF Then
                        rs.MoveFirst
                        With ThisWorkbook.Sheets("ISHIDA")
                            .Cells.clear
                            u = 1
                            .Range("A1") = "Line"
                            .Range("B1") = "Product"
                            .Range("C1") = "Unit weight [gr]"
                            .Range("D1") = "Amount [pc]"
                            .Range("E1") = "Amount [kg]"
                            .Range("F1") = "Average [gr]"
                            .Range("G1") = "Start"
                            .Range("H1") = "End"
                            .Range("I1") = "Loss [kg]"
                            Do Until rs.EOF
                                u = u + 1
                                .Range("A" & u) = rs.Fields(linCol).Value
                                .Range("B" & u) = rs.Fields(pCol).Value
                                .Range("C" & u) = rs.Fields(grCol).Value
                                .Range("D" & u) = rs.Fields(aCol).Value
                                .Range("E" & u) = rs.Fields(mCol).Value
                                .Range("F" & u) = rs.Fields(avgCol).Value
                                .Range("G" & u) = rs.Fields(fromCol).Value
                                .Range("H" & u) = rs.Fields(toCol).Value
                                .Range("I" & u) = rs.Fields(wasCol).Value
                                rs.MoveNext
                            Loop
                            lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
                            Set rng = .Range("A1:I" & lastRow)
                            rng.Sort Key1:=.Range("I1"), order1:=xlDescending, header:=xlYes
                        End With
                        summarizeISHIDA
                        formatISHIDA
                    End If
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
        End If
    End If
End If

Set rs = Nothing
End Sub

Public Sub transferPW()
Dim rng As Range
Dim rng2 As Range
Dim c As Range
Dim pw As Worksheet
Dim dest As Worksheet
Dim lastRow As Long
Dim lastRow2 As Long
Dim theRow As Long

On Error GoTo err_trap

Set pw = ThisWorkbook.Sheets("QGUAR")
Set dest = ThisWorkbook.Sheets("ZFIN Comp")

lastRow = pw.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = pw.Range("A3:A" & lastRow)
    
rng.Copy dest.Range("A3")

Set rng = pw.Range("D3:D" & lastRow)

rng.Copy dest.Range("E3")

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
    MsgBox "Error in zfinMES. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Sub summarizeISHIDA()
Dim c As Range
Dim rng As Range
Dim last As Long

On Error GoTo err_trap

With ThisWorkbook.Sheets("ISHIDA")
    .Range("L1") = "Product"
    .Range("M1") = "Unit weight [kg]"
    .Range("N1") = "Amount [kg]"
    .Range("O1") = "Total loss [kg]"
    .Range("P1") = "Loss [%]"
    last = .Range("B:B").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("B2:B" & last)
    For Each c In rng
        last2 = .Range("L:L").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
        If WorksheetFunction.CountIf(.Range("L2:L" & last2), c.Value) = 0 Then
            .Range("L" & Rows.Count).End(xlUp)(2) = c.Value
            .Range("M" & Rows.Count).End(xlUp)(2) = c.Offset(0, 1)
        End If
    Next c
    last = .Range("L:L").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("L2:L" & last)
    last2 = .Range("B:B").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng2 = .Range("B2:B" & last2)
    For Each c In rng
        c.Offset(0, 2) = WorksheetFunction.SumIf(rng2, c.Value, .Range("E2:E" & last2))
        c.Offset(0, 3) = WorksheetFunction.SumIf(rng2, c.Value, .Range("I2:I" & last2))
        c.Offset(0, 4).Formula = "=O" & c.row & "/N" & c.row
        c.Offset(0, 4).NumberFormat = "0.00%"
    Next c
    last2 = .Range("L:L").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    .Range("L1:P" & last2).Sort Key1:=.Range("O1"), order1:=xlDescending, header:=xlYes
End With

exit_here:
Set rng = Nothing
Exit Sub

err_trap:
If Err.Number = 91 Then
    Resume Next
Else
    MsgBox "Error in summarizeISHIDA. Error number: " & Err.Number & ", " & Err.Description
    Resume err_trap
End If
End Sub

Public Sub formatISHIDA()
Dim rng As Range
Dim lastRow As Long

With ThisWorkbook.Sheets("ISHIDA")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A1:I" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("A1:I1")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
    
    lastRow = .Range("L:L").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("L1:P" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("L1:P1")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
End With

Set rng = Nothing

End Sub

Public Sub transferIshida()
Dim rng As Range
Dim rng2 As Range
Dim c As Range
Dim ish As Worksheet
Dim dest As Worksheet
Dim lastRow As Long
Dim lastRow2 As Long
Dim theRow As Long

On Error GoTo err_trap

Set ish = ThisWorkbook.Sheets("ISHIDA")
Set dest = ThisWorkbook.Sheets("ZFIN Comp")

lastRow = ish.Range("L:L").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
lastRow2 = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = ish.Range("L2:L" & lastRow)
Set rng2 = dest.Range("A3:A" & lastRow2)
    
For Each c In rng
    If WorksheetFunction.CountIf(rng2, c.Value) = 0 Then
        dest.Cells(Rows.Count, 1).End(xlUp)(2) = c.Value
    End If
Next c

' the other way around
lastRow = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
lastRow2 = ish.Range("L:L").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = dest.Range("A3:A" & lastRow)
Set rng2 = ish.Range("L2:L" & lastRow2)

For Each c In rng
'    dest.Range("M" & c.row).Formula = "=K" & c.row & "-L" & c.row
'    dest.Range("G" & c.row).Formula = "=E" & c.row & "-F" & c.row
'    dest.Range("J" & c.row).Formula = "=H" & c.row & "-I" & c.row
'    dest.Range("M" & c.row).NumberFormat = "0.0"
'    dest.Range("J" & c.row).NumberFormat = "0.0"
'    dest.Range("G" & c.row).NumberFormat = "0.0"
    dest.Range("F" & c.row) = WorksheetFunction.SumIf(rng2, c.Value, ish.Range("N2:N" & lastRow2))
Next c
If ThisWorkbook.Sheets("BM").Range("G42") = "" Then ThisWorkbook.Sheets("BM").Range("G42").Formula = "=SUM(ISHIDA!O2:O" & lastRow2 & ")"
'Set rng2 = mes.Range("G3:G" & lastRow2)
'
'For Each c In rng
'    theRow = 0
'    theRow = rng2.Find(c.value, searchorder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole).row
'    If Not theRow = 0 Then
'        dest.Range("L" & c.row) = mes.Range("J" & theRow)
'    End If
'Next c

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
    MsgBox "Error in zfinMES. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If


End Sub

Public Sub importTIPTOP()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim matCol As String
Dim descCol As String
Dim grCol As String
Dim aCol As String
Dim mCol As String
Dim avgCol As String
Dim fromCol As String
Dim toCol As String
Dim wasCol As String
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim prop As String
'Dim zlec As clsZlecenie
Dim ver As String
Dim allSheets As Variant
Dim lastRow As Long
Dim rng As Range
Dim theRow As Long

prop = "tiptop sap file"

On Error GoTo err_trap

If propertyExists(prop) And ThisWorkbook.CustomDocumentProperties(prop) <> "" Then
    If ThisWorkbook.CustomDocumentProperties("import path") <> "" Then
        If FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls"
        ElseIf FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx"
        ElseIf FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsm") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsm"
        Else
            MsgBox "Source file """ & ThisWorkbook.CustomDocumentProperties(prop) & """ could not be found in " & ThisWorkbook.CustomDocumentProperties("import path") & "\ . Check in settings if both file name and path are correct.", vbOKOnly + vbExclamation, "Error"
        End If
        
        If path <> "" Then
            Set rs = importExcelData(path, "SAPBW_DOWNLOAD", 1)
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If Len(matCol) = 0 Or Len(descCol) = 0 Or Len(aCol) = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            Select Case rs.Fields(i)
                            Case Is = "Description"
                                descCol = rs.Fields(i).Name
                            Case Is = "Material"
                                matCol = rs.Fields(i).Name
                            End Select
                            If InStr(1, rs.Fields(i).Value, "Actual", vbTextCompare) > 0 Then
                                aCol = rs.Fields(i).Name
                            End If
                        Next i
                    Else
                        rs.Close
                        Set rs = Nothing
                        Exit Do
                    End If
                    rs.MoveNext
                Loop
                If Right(path, 1) = "s" Then
                    ver = "8"
                Else
                    ver = "12"
                End If
'                allSheets = getExcelSheetName(path)
                Worksheet = "SAPBW_DOWNLOAD$"
                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=1;"";"
                'by order number
'                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sub." & nCol & ", sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE " & grCol & " > 20) sub GROUP BY " & oCol & ", " & pCol & ", " & nCol & ";"
                'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
                strSQL = "SELECT * FROM [" & Worksheet & "] WHERE " & descCol & " is not null;"
                'strSQL = "SELECT * FROM [" & Worksheet & "];"
                cnn.Open conStr
                Set rs = New ADODB.Recordset
                rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText

                If Not rs.EOF Then
                    rs.filter = descCol & " <> 'Description' AND " & descCol & " <> 'Result'"
                    If Not rs.EOF Then
                        rs.MoveFirst
                        With ThisWorkbook.Sheets("ZFIN Comp")
                            Do Until rs.EOF
                                lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
                                If WorksheetFunction.CountIf(.Range("A3:A" & lastRow), rs.Fields(matCol).Value) = 0 Then
                                    'add new one
                                    theRow = lastRow + 1
                                    If Not IsNull(rs.Fields(matCol).Value) Then .Range("A" & theRow) = CLng(rs.Fields(matCol).Value)
                                    .Range("B" & theRow) = rs.Fields(descCol).Value
                                    If Not IsNull(rs.Fields(aCol).Value) Then .Range("G" & theRow) = CDbl(rs.Fields(aCol).Value)
                                Else
                                    'find where is it and add figure
                                    theRow = .Range("A3:A" & lastRow).Find(rs.Fields(matCol).Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
                                    .Range("B" & theRow) = rs.Fields(descCol).Value
                                    If Not IsNull(rs.Fields(aCol).Value) Then .Range("G" & theRow) = CDbl(rs.Fields(aCol).Value)
                                End If
                                rs.MoveNext
                            Loop
                        End With
                    End If
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
        End If
    End If
End If

exit_here:
Set rs = Nothing
Exit Sub

err_trap:
MsgBox "Error in importTIPTOP. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Sub finishIt()
Dim rng As Range
Dim lastRow As Long
Dim sSql As String
Dim prodStr As String
Dim rs As ADODB.Recordset
Dim i As Integer
Dim theRow As Long
Dim res As VbMsgBoxResult
Dim bool As Boolean
Dim rs2 As ADODB.Recordset

On Error GoTo err_trap

Set conn = New ADODB.Connection
conn.Open ConnectionString
conn.CommandTimeout = 90

With ThisWorkbook.Sheets("ZFIN Comp")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A3:A" & lastRow)
    For Each c In rng
        If IsNumeric(c.Value) And c.Value > 0 Then
            prodStr = prodStr & c.Value & ","
        End If
    Next c
    prodStr = Left(prodStr, Len(prodStr) - 1)
    
    sSql = "SELECT tbZfin.zfinId, zfinIndex, zfinName, [beans?] FROM tbZfinProperties RIGHT JOIN tbZfin on tbZfin.zfinId = tbZfinProperties.zfinId WHERE tbZfin.zfinIndex IN (" & prodStr & ");"
    Set rs = New ADODB.Recordset
    rs.Open sSql, conn, adOpenStatic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            If WorksheetFunction.CountIf(rng, rs.Fields("zfinIndex")) > 0 Then
                theRow = rng.Find(rs.Fields("zfinIndex").Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
                If Not IsNull(rs.Fields("beans?").Value) Then
                    If Not IsNull(.Range("C" & theRow)) Then .Range("C" & theRow) = Abs(rs.Fields("beans?").Value)
                Else
                    'this zfin has no bean info
                    res = MsgBox("I can't recognize if product " & rs.Fields("zfinIndex").Value & " " & rs.Fields("zfinName").Value & " is bean or ground product. Is it BEAN product?", vbYesNoCancel + vbInformation, "User's input requested")
                    If res <> vbCancel Then
                        If res = vbYes Then
                            bool = True
                        Else
                            bool = False
                        End If
                        Set rs2 = New ADODB.Recordset
                        rs2.Open "SELECT zfinId, [beans?] FROM tbZfinProperties WHERE zfinId = " & rs.Fields("zfinId"), conn, adOpenDynamic, adLockOptimistic
                        If rs2.EOF Then
                            rs2.Close
                            sSql = "INSERT INTO tbZfinProperties (zfinId, [beans?]) VALUES (" & rs.Fields("zfinId") & ", " & CInt(bool) & ");"
                            rs2.Open sSql, conn, adOpenKeyset, adLockOptimistic
                        Else
                            rs2.MoveFirst
                            rs2.Fields("beans?") = bool
                            rs2.Update
                            rs2.Close
                        End If
                        If Not IsNull(.Range("C" & theRow)) Then .Range("C" & theRow) = Abs(CInt(bool))
                    End If
                End If
                If .Range("B" & theRow) = "" Then .Range("B" & theRow) = rs.Fields("zfinName").Value
                .Range("I" & theRow).Formula = "=E" & theRow & "-F" & theRow
                .Range("K" & theRow).Formula = "=ABS(E" & theRow & "-F" & theRow & ")"
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A3:K" & lastRow)
    rng.Sort Key1:=.Range("K1"), order1:=xlDescending, header:=xlNo
    Set rng = .Range("A1:J" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("A1:J2")
    rng.Interior.ColorIndex = 15
End With
conn.Close

exit_here:
Set conn = Nothing
Set rng = Nothing
Set rs = Nothing
Set rs2 = Nothing
Exit Sub

err_trap:
MsgBox "Error in finishIt. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Sub importPlanNO()
Dim y As Integer
Dim fso As New FileSystemObject
Dim ndFile As Object
Dim path As String
Dim zrodlo As Object
Dim w As String
Dim theName As String
Dim rs As ADODB.Recordset
Dim cnn As ADODB.Connection
Dim theDate As Date
Dim theType As Integer
Dim aName As String
Dim pName As String
Dim aCol As String
Dim pCol As String
Dim conn As ADODB.Connection
Dim lastRow As Long
Dim rs2 As ADODB.Recordset
Dim c As Range
Dim rng As Range
Dim val As Double
Dim uni As Double
Dim theRow As Long

    On Error GoTo err_trap

If ThisWorkbook.CustomDocumentProperties("yearLoaded") = year(Date) Then
    path = "K:\Dział Planowania\PLANOWANIE_PPDS\Plan_PPDS_z_Mesa\"
Else
    path = "K:\Dział Planowania\PLANOWANIE_PPDS\Plan_PPDS_z_Mesa\" & yearLoaded & "\"
End If

Set zrodlo = fso.GetFolder(path)
If ThisWorkbook.CustomDocumentProperties("weekLoaded") < 10 Then
    w = "0" & ThisWorkbook.CustomDocumentProperties("weekLoaded")
Else
    w = ThisWorkbook.CustomDocumentProperties("weekLoaded")
End If

For Each ndFile In zrodlo.Files
    If Mid(ndFile.Name, 2, 2) = w Then
        If Len(theName) > 0 Then
            If DateDiff("d", theDate, ndFile.DateLastModified) > 0 Then
                theName = ndFile.Name
                theDate = ndFile.DateLastModified
            End If
        Else
            theName = ndFile.Name
            theDate = ndFile.DateLastModified
        End If
    End If
Next ndFile

path = path & theName

Set rs = importExcelData(path, , 1)
If Not rs.EOF Then
    rs.MoveFirst
    If rs.Fields(0).Value = "GRID_PROD_GRAPH" Then
        theType = 2
    ElseIf rs.Fields(0).Value = "GRID_PROD_CONTROL" Then
        theType = 1
    Else
        theType = 0
    End If
    If theType = 0 Then
        MsgBox "Error in importPlan. Unrecognized report type given as source data.", vbCritical + vbOKOnly
    Else
        If theType = 2 Then
            aName = "Il. plan. [j. art.]"
            pName = "Nr produktu"
        Else
            aName = "Il. plan. [j. art.]"
            pName = "Artykuł"
        End If
        Do Until rs.EOF
            If Len(aCol) = 0 Or Len(pCol) = 0 Then
                For i = 0 To rs.Fields.Count - 1
                    Select Case rs.Fields(i)
                    Case Is = aName
                        aCol = rs.Fields(i).Name
                    Case Is = pName
                        pCol = rs.Fields(i).Name
                    End Select
                Next i
            Else
                rs.Close
                Set rs = Nothing
                Exit Do
            End If
            rs.MoveNext
        Loop
        If Right(path, 1) = "s" Then
            ver = "8"
        Else
            ver = "12"
        End If
'
        allSheets = getExcelSheetName(path)
        Worksheet = allSheets(0)
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=YES;IMEX=0;"";"
        'by order number
'                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sub." & nCol & ", sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE " & grCol & " > 20) sub GROUP BY " & oCol & ", " & pCol & ", " & nCol & ";"
        'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
        strSQL = "SELECT " & pCol & ", SUM(" & aCol & ") as amount FROM [" & Worksheet & "] WHERE " & pCol & " is not null GROUP BY " & pCol & ";"
        'strSQL = "SELECT * FROM [" & Worksheet & "];"
        Set cnn = New ADODB.Connection
        cnn.Open conStr
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            rs.MoveFirst
            rs.filter = pCol & " <> 'Nr produktu'"
            Do Until rs.EOF
                prodStr = prodStr & rs.Fields(pCol).Value & ","
                rs.MoveNext
            Loop
            prodStr = Left(prodStr, Len(prodStr) - 1)
            Set conn = New ADODB.Connection
            conn.Open ConnectionString
            conn.CommandTimeout = 90
            strSQL = "SELECT zfinIndex, unitWeight FROM tbUom JOIN tbZfin on tbZfin.zfinId = tbUom.zfinId WHERE tbZfin.zfinIndex IN (" & prodStr & ");"
            Set rs2 = New ADODB.Recordset
            rs2.Open strSQL, conn, adOpenStatic, adLockOptimistic, adCmdText
            If Not rs2.EOF Then
                lastRow = ThisWorkbook.Sheets("ZFIN Comp").Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
                Set rng = ThisWorkbook.Sheets("ZFIN Comp").Range("A3:A" & lastRow)
                With ThisWorkbook.Sheets("ZFIN Comp")
                    Do While Not rs2.EOF
                        theRow = rng.Find(rs2.Fields("zfinIndex"), searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
                        uni = rs2.Fields("unitWeight").Value
                        val = rs.Fields("amount").Value
                        If uni > 0 Then
                            .Range("H" & theRow) = uni * val
                        Else
                            MsgBox "Unit weight of product " & c.Value & " couldn't be found in database. For this product quantity in ""Plan"" column is in pcs and marked in red.", vbInformation + vbOKOnly
                            .Range("H" & theRow) = val
                            .Range("H" & theRow).Interior.Color = vbRed
                        End If
                        rs2.MoveNext
                    Loop
                End With
                rs2.Close
            End If
        End If
    End If
End If

exit_here:
Set zrodlo = Nothing
Set rs = Nothing
Set cnn = Nothing
Set rs2 = Nothing
Set conn = Nothing
Exit Sub

err_trap:
If Err.Number = 91 Then
    theRow = lastRow + 1
    Resume Next
Else
    MsgBox "Error in importPlan. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If


End Sub

Sub importPlan()
Dim y As Integer
Dim fso As New FileSystemObject
Dim ndFile As Object
Dim path As String
Dim zrodlo As Object
Dim w As String
Dim theName As String
Dim rs As ADODB.Recordset
Dim cnn As ADODB.Connection
Dim theDate As Date
Dim theType As Integer
Dim aName As String
Dim pName As String
Dim aCol As String
Dim pCol As String
Dim conn As ADODB.Connection
Dim lastRow As Long
Dim rs2 As ADODB.Recordset
Dim c As Range
Dim rng As Range
Dim val As Double
Dim uni As Double

    On Error GoTo err_trap

If ThisWorkbook.CustomDocumentProperties("yearLoaded") = year(Date) Then
    path = "K:\Dział Planowania\PLANOWANIE_PPDS\Plan_PPDS_z_Mesa\"
Else
    path = "K:\Dział Planowania\PLANOWANIE_PPDS\Plan_PPDS_z_Mesa\" & yearLoaded & "\"
End If

Set zrodlo = fso.GetFolder(path)
If ThisWorkbook.CustomDocumentProperties("weekLoaded") < 10 Then
    w = "0" & ThisWorkbook.CustomDocumentProperties("weekLoaded")
Else
    w = ThisWorkbook.CustomDocumentProperties("weekLoaded")
End If

For Each ndFile In zrodlo.Files
    If Mid(ndFile.Name, 2, 2) = w Then
        If Len(theName) > 0 Then
            If DateDiff("d", theDate, ndFile.DateLastModified) > 0 Then
                theName = ndFile.Name
                theDate = ndFile.DateLastModified
            End If
        Else
            theName = ndFile.Name
            theDate = ndFile.DateLastModified
        End If
    End If
Next ndFile

path = path & theName

Set rs = importExcelData(path, , 1)
If Not rs.EOF Then
    rs.MoveFirst
    If rs.Fields(0).Value = "GRID_PROD_GRAPH" Then
        theType = 2
    ElseIf rs.Fields(0).Value = "GRID_PROD_CONTROL" Then
        theType = 1
    Else
        theType = 0
    End If
    If theType = 0 Then
        MsgBox "Error in importPlan. Unrecognized report type given as source data.", vbCritical + vbOKOnly
    Else
        If theType = 2 Then
            aName = "Il. plan. [j. art.]"
            pName = "Nr produktu"
        Else
            aName = "Il. plan. [j. art.]"
            pName = "Artykuł"
        End If
        Do Until rs.EOF
            If Len(aCol) = 0 Or Len(pCol) = 0 Then
                For i = 0 To rs.Fields.Count - 1
                    Select Case rs.Fields(i)
                    Case Is = aName
                        aCol = rs.Fields(i).Name
                    Case Is = pName
                        pCol = rs.Fields(i).Name
                    End Select
                Next i
            Else
                rs.Close
                Set rs = Nothing
                Exit Do
            End If
            rs.MoveNext
        Loop
        If Right(path, 1) = "s" Then
            ver = "8"
        Else
            ver = "12"
        End If
'
        allSheets = getExcelSheetName(path)
        Worksheet = allSheets(0)
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=YES;IMEX=0;"";"
        'by order number
'                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sub." & nCol & ", sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE " & grCol & " > 20) sub GROUP BY " & oCol & ", " & pCol & ", " & nCol & ";"
        'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
        strSQL = "SELECT " & pCol & ", SUM(" & aCol & ") as amount FROM [" & Worksheet & "] WHERE " & pCol & " is not null GROUP BY " & pCol & ";"
        'strSQL = "SELECT * FROM [" & Worksheet & "];"
        Set cnn = New ADODB.Connection
        cnn.Open conStr
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            rs.MoveFirst
            rs.filter = pCol & " <> 'Nr produktu'"
            Do Until rs.EOF
                prodStr = prodStr & rs.Fields(pCol).Value & ","
                rs.MoveNext
            Loop
            prodStr = Left(prodStr, Len(prodStr) - 1)
            Set conn = New ADODB.Connection
            conn.Open ConnectionString
            conn.CommandTimeout = 90
            strSQL = "SELECT zfinIndex, unitWeight FROM tbUom JOIN tbZfin on tbZfin.zfinId = tbUom.zfinId WHERE tbZfin.zfinIndex IN (" & prodStr & ");"
            Set rs2 = New ADODB.Recordset
            rs2.Open strSQL, conn, adOpenStatic, adLockOptimistic, adCmdText
            If Not rs2.EOF Then
                lastRow = ThisWorkbook.Sheets("ZFIN Comp").Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
                Set rng = ThisWorkbook.Sheets("ZFIN Comp").Range("A3:A" & lastRow)
                For Each c In rng
                    val = 0
                    uni = 0
                    rs2.MoveFirst
                    rs.MoveFirst
                    Do Until rs.EOF
                        If CLng(rs.Fields(pCol).Value) = c.Value Then
                            val = rs.Fields("amount").Value
                            Exit Do
                        End If
                        rs.MoveNext
                    Loop
                    If val > 0 Then
                        Do Until rs2.EOF
                            If CLng(rs2.Fields("zfinIndex").Value) = c.Value Then
                                uni = rs2.Fields("unitWeight").Value
                                Exit Do
                            End If
                            rs2.MoveNext
                        Loop
                    If uni > 0 Then
                        c.Offset(0, 7) = uni * val
                    Else
                        MsgBox "Unit weight of product " & c.Value & " couldn't be found in database. For this product quantity in ""Plan"" column is in pcs and marked in red.", vbInformation + vbOKOnly
                        c.Offset(0, 7) = val
                        c.Offset(0, 7).Interior.Color = vbRed
                    End If
                    End If
                Next c
                rs2.Close
            End If
        End If
    End If
End If

exit_here:
Set zrodlo = Nothing
Set rs = Nothing
Set cnn = Nothing
Set rs2 = Nothing
Set conn = Nothing
Exit Sub

err_trap:
MsgBox "Error in importPlan. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub


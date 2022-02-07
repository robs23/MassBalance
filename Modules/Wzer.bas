Attribute VB_Name = "Wzer"
Option Explicit
Public batches As New Collection
Public zfins As New Collection

Public Sub importQguar(control As IRibbonControl)
With ThisWorkbook.Sheets("QGUAR")
    .Cells.clear
    .Range("C1") = "PW"
    .Range("C1").Font.Bold = True
    .Range("C1").Font.Size = 18
    .Range("J1") = "WZ"
    .Range("J1").Font.Bold = True
    .Range("J1").Font.Size = 18
    .Range("Q1") = "BZ"
    .Range("Q1").Font.Bold = True
    .Range("Q1").Font.Size = 18
End With

importPw
importWz
importBz
formatQGUAR
End Sub

Public Sub importWz()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim artCol As Integer
Dim razCol As Integer
Dim boxCol As Integer
Dim aCol As Integer
Dim palCol As Integer
Dim kgCol As Integer
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim ZFIN As Long
Dim prop As String
'Dim zlec As clsZlecenie
Dim ver As String

prop = "wz qguar file"

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
            Set rs = importExcelData(path)
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If artCol = 0 Or razCol = 0 Or aCol = 0 Or palCol = 0 Or kgCol = 0 Or boxCol = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            Select Case rs.Fields(i)
                            Case Is = "Nr artykułu:"
                                artCol = i
                            Case Is = "Razem :"
                                razCol = i
                            Case Is = "Ilość"
                                aCol = i
                            Case Is = "palety"
                                palCol = i
                            Case Is = "kg netto"
                                kgCol = i
                            Case Is = "kartony"
                                boxCol = i
                            End Select
                        Next i
                    Else
                        rs.MoveFirst
                        Exit Do
                    End If
                    rs.MoveNext
                Loop

                With ThisWorkbook.Sheets("QGUAR")
                    u = 2
                    .Range("H" & u) = "ZFIN"
                    .Range("I" & u) = "PC"
                    .Range("J" & u) = "PAL"
                    .Range("K" & u) = "KG"
                    .Range("L" & u) = "BOX"
                    Do Until rs.EOF
                        
                        If rs.Fields(artCol).Value = "Nr artykułu:" And rs.Fields(artCol + 1).Value <> ZFIN Then
                            'we've got new zfin
                            ZFIN = rs.Fields(artCol + 1).Value
                            u = u + 1
                        ElseIf rs.Fields(razCol).Value = "Razem :" Then
                            'save results
                            .Range("H" & u) = ZFIN
                            .Range("I" & u) = CLng(rs.Fields(aCol).Value)
                            .Range("J" & u) = CDbl(rs.Fields(palCol).Value)
                            .Range("K" & u) = CDbl(rs.Fields(kgCol).Value)
                            .Range("L" & u) = CLng(rs.Fields(boxCol).Value)
                        End If
                        rs.MoveNext
                    Loop
                End With
            Else
                MsgBox "There's problem reading worksheet's name " & path & ". Check ""importWz"" sub"
            End If

            rs.Close
            Set rs = Nothing

        End If
    End If
End If

Set rs = Nothing
End Sub

Public Sub importPw()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim artCol As Integer
Dim razCol As Integer
Dim boxCol As Integer
Dim aCol As Integer
Dim palCol As Integer
Dim kgCol As Integer
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim ZFIN As Long
Dim prop As String
Dim sumPw As Double
'Dim zlec As clsZlecenie
Dim ver As String

prop = "pw qguar file"


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

            Set rs = importExcelData(path)
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If artCol = 0 Or razCol = 0 Or aCol = 0 Or palCol = 0 Or kgCol = 0 Or boxCol = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            Select Case rs.Fields(i)
                            Case Is = "Nr artykułu:"
                                artCol = i
                            Case Is = "Razem :"
                                razCol = i
                            Case Is = "Ilość"
                                aCol = i
                            Case Is = "palety"
                                palCol = i
                            Case Is = "kg netto"
                                kgCol = i
                            Case Is = "kartony"
                                boxCol = i
                            End Select
                        Next i
                    Else
                        rs.MoveFirst
                        Exit Do
                    End If
                    rs.MoveNext
                Loop

                With ThisWorkbook.Sheets("QGUAR")
                    u = 2
                    .Range("A" & u) = "ZFIN"
                    .Range("B" & u) = "PC"
                    .Range("C" & u) = "PAL"
                    .Range("D" & u) = "KG"
                    .Range("E" & u) = "BOX"
                    Do Until rs.EOF
                        
                        If rs.Fields(artCol).Value = "Nr artykułu:" And rs.Fields(artCol + 1).Value <> ZFIN Then
                            'we've got new zfin
                            ZFIN = rs.Fields(artCol + 1).Value
                            u = u + 1
                        ElseIf rs.Fields(razCol).Value = "Razem :" Then
                            'save results
                            .Range("A" & u) = ZFIN
                            .Range("B" & u) = CLng(rs.Fields(aCol).Value)
                            .Range("C" & u) = CDbl(rs.Fields(palCol).Value)
                            .Range("D" & u) = CDbl(rs.Fields(kgCol).Value)
                            sumPw = sumPw + CDbl(rs.Fields(kgCol).Value)
                            .Range("E" & u) = CLng(rs.Fields(boxCol).Value)
                        End If
                        rs.MoveNext
                    Loop
                End With
                If ThisWorkbook.Sheets("BM").Range("H45") = "" Then ThisWorkbook.Sheets("BM").Range("H45") = sumPw
            Else
                MsgBox "There's problem reading worksheet's name " & path & ". Check ""importPw"" sub"
            End If

            rs.Close
            Set rs = Nothing

        End If
    End If
End If

Set rs = Nothing
End Sub

Public Sub importBz()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim oCol As String
Dim pCol As String
Dim aCol As String
Dim palCol As String
Dim statCol As String
Dim expCol As String
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim allSheets As Variant
Dim prop As String
'Dim zlec As clsZlecenie
Dim ver As String

prop = "bz qguar file"
'On Error Resume Next

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
                    If Len(oCol) = 0 And Len(pCol) = 0 And Len(aCol) = 0 And Len(expCol) = 0 And Len(palCol) = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            If Not IsNull(rs.Fields(i)) Then
                                Select Case rs.Fields(i)
                                Case Is = "Partia"
                                    oCol = rs.Fields(i).Name
                                Case Is = "Nr artykułu"
                                    pCol = rs.Fields(i).Name
                                Case Is = "Ilość"
                                    aCol = rs.Fields(i).Name
                                Case Is = "Data ważn."
                                    expCol = rs.Fields(i).Name
                                Case Is = "Poz."
                                    palCol = rs.Fields(i).Name
                                End Select
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
                
                If Not IsNull(allSheets) Then
                    If UBound(allSheets) > 0 Then MsgBox "There's more than 1 worksheet in the source file. As it's not set which worksheet contains data, the first one was chosen (""" & allSheets(0) & """). In case there's no data imported or you suspect errors, please remove from the source file all sheets but the desired one and try again", vbInformation + vbOKOnly, "Possible errors"
                    Worksheet = allSheets(0)
                End If
                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=0;"";"
                'by order number
                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sum(sub." & aCol & ") as amount, sum(sub." & palCol & ") as pals, min(sub." & expCol & ") as expMin, max(sub." & expCol & ") as expMax FROM ( SELECT * FROM [" & Worksheet & "]  WHERE " & aCol & " >=2) sub GROUP BY " & oCol & ", " & pCol & ";"
                'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
                'strSQL = "SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100;"
                cnn.Open conStr
                Set rs = New ADODB.Recordset
                rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs.EOF Then
                    rs.MoveFirst
                    With ThisWorkbook.Sheets("QGUAR")
                        u = 2
                        .Range("O2") = "Batch"
                        .Range("P2") = "ZFIN"
                        .Range("Q2") = "Expiration Min"
                        .Range("R2") = "Expiration Max"
                        .Range("S2") = "Amount [pc]"
                        .Range("T2") = "Amount [pal]"
                        Do Until rs.EOF
                            u = u + 1
                            .Range("O" & u) = rs.Fields(oCol).Value
                            .Range("P" & u) = rs.Fields(pCol).Value
                            .Range("Q" & u) = rs.Fields("expMin").Value
                            .Range("R" & u) = rs.Fields("expMax").Value
                            .Range("S" & u) = rs.Fields("amount").Value
                            .Range("T" & u) = rs.Fields("pals").Value
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
                uploadStock
            End If
        End If
    End If
End If

Set rs = Nothing

End Sub

'Public Sub importWzpb(control As IRibbonControl)
'Dim rs As ADODB.Recordset
'Dim path As String
'Dim i As Integer
'Dim artCol As Integer
'Dim razCol As Integer
'Dim boxCol As Integer
'Dim aCol As Integer
'Dim palCol As Integer
'Dim kgCol As Integer
'Dim strSQL As String
'Dim cnn As New ADODB.Connection
'Dim conStr As String
'Dim u As Integer
'Dim Worksheet As String
'Dim zfin As Long
'Dim prop As String
''Dim zlec As clsZlecenie
'Dim ver As String
'
'prop = "wz qguar file"
'
'ThisWorkbook.Sheets("WZ").Cells.clear
'
'If propertyExists(prop) And ThisWorkbook.CustomDocumentProperties(prop) <> "" Then
'    If ThisWorkbook.CustomDocumentProperties("import path") <> "" Then
'        If FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls") Then
'            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls"
'        ElseIf FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx") Then
'            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx"
'        Else
'            MsgBox "Source file """ & ThisWorkbook.CustomDocumentProperties(prop) & """ could not be found in " & ThisWorkbook.CustomDocumentProperties("import path") & "\ . Check in settings if both file name and path are correct.", vbOKOnly + vbExclamation, "Error"
'        End If
'
'        If path <> "" Then
'            If Right(path, 1) = "s" Then
'                ver = "8"
'            Else
'                ver = "12"
'            End If
'            Worksheet = "'$'"
'            End If
'
'            If Not Worksheet = "" Then
'                Set rs = New ADODB.Recordset
'                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=0;"";"
'                strSQL = "SELECT * FROM [" & Worksheet & "];"
'                cnn.Open conStr
'                rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
'                If Not rs.EOF Then
'                    rs.MoveFirst
'                    Do Until rs.EOF
'                        If artCol = 0 Or razCol = 0 Or aCol = 0 Or palCol = 0 Or kgCol = 0 Or boxCol = 0 Then
'                            For i = 0 To rs.Fields.Count - 1
'                                Select Case rs.Fields(i)
'                                Case Is = "Nr artykułu:"
'                                    artCol = i
'                                Case Is = "Razem :"
'                                    razCol = i
'                                Case Is = "Ilość"
'                                    aCol = i
'                                Case Is = "palety"
'                                    palCol = i
'                                Case Is = "kg netto"
'                                    kgCol = i
'                                Case Is = "kartony"
'                                    boxCol = i
'                                End Select
'                            Next i
'                        Else
'                            rs.MoveFirst
'                            Exit Do
'                        End If
'                        rs.MoveNext
'                    Loop
'
'                    With ThisWorkbook.Sheets("WZ")
'                        u = 1
'                        .Range("A" & u) = "ZFIN"
'                        .Range("B" & u) = "PC"
'                        .Range("C" & u) = "PAL"
'                        .Range("D" & u) = "KG"
'                        .Range("E" & u) = "BOX"
'                        Do Until rs.EOF
'
'                            If rs.Fields(artCol).value = "Nr artykułu:" And rs.Fields(artCol + 1).value <> zfin Then
'                                'we've got new zfin
'                                zfin = rs.Fields(artCol + 1).value
'                                u = u + 1
'                            ElseIf rs.Fields(razCol).value = "Razem :" Then
'                                'save results
'                                .Range("A" & u) = zfin
'                                .Range("B" & u) = CLng(rs.Fields(aCol).value)
'                                .Range("C" & u) = CDbl(rs.Fields(palCol).value)
'                                .Range("D" & u) = CDbl(rs.Fields(kgCol).value)
'                                .Range("E" & u) = CInt(rs.Fields(boxCol).value)
'                            End If
'                            rs.MoveNext
'                        Loop
'                    End With
'            Else
'                MsgBox "There's problem reading worksheet's name " & path & ". Check ""importWz"" sub"
'            End If
'
'            rs.Close
'            Set rs = Nothing
'
'        End If
'    End If
'End If
'
'Set rs = Nothing
'End Sub


Function buildSQL(colNames As String, rng As Range) As String
Dim SQL As String
Dim sCell As Range
Dim i As Integer

If Not rng Is Empty And Not rng Is Nothing Then
    sCell = rng.item(1)
    
    For i = 1 To rng.Rows.Count
    
    Next i
End If

End Function

Public Sub formatQGUAR()
Dim rng As Range
Dim lastRow As Long

With ThisWorkbook.Sheets("QGUAR")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A2:E" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("A2:E2")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
    
    lastRow = .Range("H:H").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("H2:L" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("H2:L2")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
    
    lastRow = .Range("O:O").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("O2:T" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("O2:T2")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
End With

Set rng = Nothing

End Sub


Public Sub uploadStock()
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Dim Id As Long
Dim sqlStr As String
Dim theDate As Date 'the date of the moment when data was dropped from QGUAR
Dim putin As Variant
Dim addit As String
Dim rng As Range
Dim c As Range
Dim isError As Boolean
Dim lastRow As Long

On Error GoTo err_trap
initializeObjects
addMissingBatch

theDate = DateAdd("d", 7, DateAdd("h", 6, 7 * (w - 1) + DateSerial(y, 1, 4) - Weekday(DateSerial(y, 1, 4), 2) + 1))

Set conn = New ADODB.Connection

conn.Open ConnectionString
conn.CommandTimeout = 90

sqlStr = "SELECT invReconciliationId FROM tbInventoryReconciliation WHERE invDate = '" & theDate & "';"
'createProducts ("zfin")
Set rs = conn.Execute(sqlStr)
If Not rs.EOF Then
    'there's already openning stock for desired period yet
    rs.MoveFirst
    Id = rs.Fields("invReconciliationId")
    Set rs = conn.Execute("DELETE FROM tbStocks WHERE invReconciliationId = " & Id)
    Set rs = conn.Execute("DELETE FROM tbInventoryReconciliation WHERE invReconciliationId = " & Id)
End If
sqlStr = "INSERT INTO tbInventoryReconciliation(invDate, invCreatedOn, week, year) VALUES ('" & theDate & "','" & Now & "', " & IsoWeekNumber(theDate) & ", " & year(theDate) & ");SELECT SCOPE_IDENTITY() AS ID;"

Set rs = conn.Execute(sqlStr)

Id = rs.Fields(0).Value


With ThisWorkbook.Sheets("QGUAR")
    lastRow = .Range("O:O").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("O3:O" & lastRow)
    For Each c In rng
        isError = False
        sqlStr = "INSERT INTO tbStocks(batchId, StockSize, invReconciliationId) VALUES (" & batches(CStr(c)).bId & ", " & CDbl(c.Offset(0, 4)) & ", " & Id & ");"
        If isError = False Then Set rs = conn.Execute(sqlStr)
    Next c
End With


exit_here:
Set rs = Nothing
Set rng = Nothing
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
End If
Set conn = Nothing
Exit Sub

err_trap:
If Err.Number = 5 Then
    'we've hit new product, ommit it
    MsgBox "Product " & c.Offset(0, 1) & " has not been found in database. It will be omitted from stock upload", vbOKOnly + vbExclamation
    isError = True
    Resume Next
Else
    MsgBox "Error in uploadStock. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Public Sub addMissingBatch()
Dim str As String
Dim lastRow As Long
Dim rng As Range
Dim batchStr As String
Dim cSql As String
Dim iSql As String
Dim sSql As String
Dim uSql As String
Dim zfinStr As String
Dim c As Range

updateConnection

With ThisWorkbook.Sheets("QGUAR")
    lastRow = .Range("P:P").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    downloadZfins "'zfin'"
    
'    Set rng = .Range("P3:P" & lastRow)
'    For Each c In rng
'        zfinStr = zfinStr & "(" & c & ",'" & c.Offset(2, 0) & "','zfin','pr','" & Now & "',43),"
'    Next c
'
'    cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nchar(255),zfinType nchar(4),prodStatus nchar(2),creationDate datetime,createdBy int)"
'    adoConn.Execute cSql
'    iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) VALUES " & zfinStr
'    adoConn.Execute iSql
'    sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin)"
'    iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) " & sSql
'    adoConn.Execute iSql


    Set rng = .Range("O3:O" & lastRow)
    For Each c In rng
        batchStr = batchStr & "(" & c & ",'" & c.Offset(0, 2) & "','" & c.Offset(0, 3) & "'," & zfins(CStr(c.Offset(0, 1))).zfinId & "),"
    Next c
End With

batchStr = Left(batchStr, Len(batchStr) - 1)


cSql = "CREATE TABLE #batches(batchNumber bigint,expEarly datetime, expLate datetime, zfinId int)"
adoConn.Execute cSql
iSql = "INSERT INTO #batches(batchNumber,expEarly,expLate,zfinId) VALUES " & batchStr
adoConn.Execute iSql
sSql = "SELECT DISTINCT batchNumber,expEarly,expLate,zfinId FROM #batches WHERE batchNumber NOT IN (SELECT batchNumber FROM tbBatch)"
iSql = "INSERT INTO tbBatch (batchNumber,expEarly,expLate,zfinId) " & sSql
adoConn.Execute iSql
uSql = "UPDATE tbBatch SET tbBatch.expEarly = #batches.expEarly, tbBatch.expLate = #batches.expLate " _
        & "FROM tbBatch INNER JOIN #batches ON tbBatch.batchNumber = #batches.batchNumber " _
        & "WHERE tbBatch.expEarly Is Null"
adoConn.Execute uSql

downloadBatches

closeConnection
End Sub

Private Sub downloadBatches()
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newBatch As clsBatch

On Error GoTo err_trap

n = batches.Count
Do While batches.Count > 0
    batches.Remove n
    n = n - 1
Loop

sSql = "SELECT batchId, batchNumber FROM tbBatch;"
Set rs = New ADODB.Recordset
rs.Open sSql, adoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newBatch = New clsBatch
        With newBatch
            .bId = rs.Fields("batchId").Value
            .bNumber = rs.Fields("batchNumber").Value
            batches.Add newBatch, CStr(rs.Fields("batchNumber").Value)
        End With
        rs.MoveNext
    Loop
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadBatches. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub downloadZfins(typeStr As String)
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newZfin As clsZfin

On Error GoTo err_trap

n = zfins.Count
Do While zfins.Count > 0
    zfins.Remove n
    n = n - 1
Loop

sSql = "SELECT zfinId, zfinIndex FROM tbZfin WHERE zfinType IN (" & typeStr & ");"
Set rs = New ADODB.Recordset
rs.Open sSql, adoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newZfin = New clsZfin
        With newZfin
            .zfinId = rs.Fields("zfinId").Value
            .zfinIndex = rs.Fields("zfinIndex").Value
            zfins.Add newZfin, CStr(rs.Fields("zfinIndex").Value)
        End With
        rs.MoveNext
    Loop
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadZfins. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub


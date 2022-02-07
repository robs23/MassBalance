Attribute VB_Name = "MWG"
Public products As New Collection
Public adoConn As ADODB.Connection

Public Sub updateConnection()

If Not adoConn Is Nothing Then
    If adoConn.State = 0 Then
        adoConn.Open ConnectionString
        adoConn.CommandTimeout = 90
    End If
Else
    Set adoConn = New ADODB.Connection
    adoConn.Open ConnectionString
    adoConn.CommandTimeout = 90
End If
End Sub

Public Sub closeConnection()

If Not adoConn Is Nothing Then
    If adoConn.State = 1 Then
        adoConn.Close
    End If
    Set adoConn = Nothing
End If
End Sub

Public Sub getStockData()
initializeObjects
With ThisWorkbook.Sheets("MWG")
    .Cells.clear
    .Range("A1") = "ZFIN"
    .Range("B1") = "Description"
    .Range("C1") = "Opening balance"
    .Range("D1") = "PW"
    .Range("E1") = "WZ"
    .Range("F1") = "Other"
    .Range("G1") = "Closing balance"
    .Range("H1") = "Difference"
    .Range("I1") = "Comment"
    .Range("A1:I1").Font.Bold = True
    .Range("A1:I1").HorizontalAlignment = xlCenter
    .Range("A1:I1").Interior.ColorIndex = 15
End With
downloadStock w.Value, y.Value, True
downloadStock w.Value + 1, y.Value, False
transferOperations
finalize
End Sub

Public Sub saveStock()
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

putin = "xx"
isError = False
Do Until (IsDate(putin) Or putin = "") And isError = False
    putin = InputBox(addit & "Please put in date of the drop from QGUAR in valid date format", "Date of data", Now)
    isError = False
    If Not IsDate(putin) Then
        addit = "Your input was not in date format. "
    ElseIf CDate(putin) > Now Then
        addit = "The date you enter can't be in the future. "
        isError = True
    End If
Loop
If IsDate(putin) Then
    theDate = CDate(putin)
    Set conn = New ADODB.Connection
    conn.Open ConnectionString
    conn.CommandTimeout = 90
    '
    'Set rs = New ADODB.Recordset
    'Set rs = Conn.Execute("SELECT * FROM tbBM WHERE bmWeek = " & week, , adCmdText)
    sqlStr = "INSERT INTO tbInventoryReconciliation(invDate, invCreatedOn, week, year) VALUES ('" & theDate & "','" & Now & "', " & IsoWeekNumber(theDate) & ", " & year(theDate) & ");SELECT SCOPE_IDENTITY() AS ID;"
    
    'Set rs = New ADODB.Recordset
    'rs.Open sqlstr, conn, adOpenKeyset, adLockOptimistic
    Set rs = conn.Execute(sqlStr)
    'Set rs = rs.NextRecordset
    Id = rs.Fields(0).Value
    createProducts ("zfin")
    With ThisWorkbook.Sheets("QGUAR")
        lastRow = .Range("O:O").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
        Set rng = .Range("O3:O" & lastRow)
        For Each c In rng
            isError = False
            sqlStr = "INSERT INTO tbBatches(batchSize, batchNumber, expirationEarly, expirationLate, zfinId, invReconciliationId) VALUES (" & CLng(c.Offset(0, 4)) & ", " & CDbl(c.Value) & ", '" & c.Offset(0, 2) & "', '" & c.Offset(0, 3) & "', " & products(CStr(c.Offset(0, 1))).prodId & ", " & Id & ");"
            If isError = False Then Set rs = conn.Execute(sqlStr)
        Next c
    End With
Else
    MsgBox "Action has been aborted by user", vbOKOnly + vbExclamation, "Break"
End If

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
    MsgBox "Error in saveStock. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Public Sub createProducts(t As String)
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim newProd As clsProduct
Dim n As Integer

On Error GoTo err_trap

n = products.Count
Do While products.Count > 0
    products.Remove n
    n = n - 1
Loop

Set conn = New ADODB.Connection
conn.Open ConnectionString
conn.CommandTimeout = 90
'
'Set rs = New ADODB.Recordset
'Set rs = Conn.Execute("SELECT * FROM tbBM WHERE bmWeek = " & week, , adCmdText)
sqlStr = "SELECT zfinId, zfinIndex, zfinName FROM tbZfin WHERE zfinType = '" & t & "';"

'Set rs = New ADODB.Recordset
'rs.Open sqlstr, conn, adOpenKeyset, adLockOptimistic
Set rs = conn.Execute(sqlStr)
Do While rs.EOF = False
    Set newProd = New clsProduct
    newProd.prodId = rs.Fields("zfinId").Value
    newProd.prodIndex = rs.Fields("zfinIndex").Value
    newProd.prodName = rs.Fields("zfinName").Value
    products.Add newProd, CStr(rs.Fields("zfinIndex").Value)
    rs.MoveNext
Loop

exit_here:
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in createProducts. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub


Public Sub downloadStock(week As Integer, year As Integer, op As Boolean)
Dim col As String
Dim rs As ADODB.Recordset
Dim theRow As Long
Dim lastRow As Long
Dim rng As Range
Dim lacking As Boolean

On Error GoTo err_trap

updateConnection
If op Then
    col = "C"
Else
    col = "G"
End If

'sql = "SELECT z.zfinIndex, z.zfinName, SUM(b.batchSize) As Amount FROM tbBatches b LEFT JOIN tbZfin z ON z.zfinId = b.zfinId " _
'    & "WHERE invReconciliationId = (Select TOP(1) invReconciliationId FROM tbInventoryReconciliation WHERE week = " & week & " And year = " & year _
'    & " ORDER BY invDate ASC) GROUP BY z.zfinName, z.zfinIndex;"

SQL = "SELECT z.zfinIndex, z.zfinName, SUM(s.stockSize) As Amount FROM tbStocks s LEFT JOIN tbBatch b ON s.batchId=b.batchId LEFT JOIN tbZfin z ON z.zfinId = b.zfinId " _
    & "WHERE s.invReconciliationId = (Select TOP(1) invReconciliationId FROM tbInventoryReconciliation WHERE week = " & week & " And year = " & year & " ORDER BY invDate ASC) " _
    & "GROUP BY z.zfinName, z.zfinIndex;"
Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
    MsgBox "Stock data for week " & week & " of year " & year & " coudn't be found.", vbOKOnly + vbExclamation, "No data for chosen period"
Else
    If op Then
        With ThisWorkbook.Sheets("MWG")
            .Range("A2").CopyFromRecordset rs
        End With
    Else
        With ThisWorkbook.Sheets("MWG")
            Do While Not rs.EOF
                lacking = False
                lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
                Set rng = .Range("A2:A" & lastRow)
                theRow = rng.Find(rs.Fields("zfinIndex"), searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
                If lacking Then
                    .Range("A" & theRow) = rs.Fields("zfinIndex")
                    .Range("B" & theRow) = rs.Fields("zfinName")
                    .Range(col & theRow) = rs.Fields("amount")
                Else
                    .Range(col & theRow) = rs.Fields("amount")
                End If
                rs.MoveNext
            Loop
        End With
    End If
End If

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Set rng = Nothing
Exit Sub

err_trap:
If Err.Number = 91 Then
    theRow = lastRow + 1
    lacking = True
    Resume Next
Else
    MsgBox "Error in downloadStock. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Public Sub transferOperations()
Dim rng As Range
Dim rng2 As Range
Dim c As Range
Dim pw As Worksheet
Dim dest As Worksheet
Dim lacking As Boolean
Dim lastRow As Long
Dim lastRow2 As Long
Dim theRow As Long

On Error GoTo err_trap

Set pw = ThisWorkbook.Sheets("QGUAR")
Set dest = ThisWorkbook.Sheets("MWG")

lastRow = pw.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = pw.Range("A3:A" & lastRow)
    
For Each c In rng
    lacking = False
    lastRow2 = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng2 = dest.Range("A2:A" & lastRow2)
    theRow = rng2.Find(c.Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
    If lacking Then
        dest.Range("A" & theRow) = c.Value
        dest.Range("D" & theRow) = c.Offset(0, 1)
    Else
        dest.Range("D" & theRow) = c.Offset(0, 1)
    End If
Next c

lastRow = pw.Range("H:H").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = pw.Range("H3:H" & lastRow)
    
For Each c In rng
    lacking = False
    lastRow2 = dest.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng2 = dest.Range("A2:A" & lastRow2)
    theRow = rng2.Find(c.Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
    If lacking Then
        dest.Range("A" & theRow) = c.Value
        dest.Range("E" & theRow) = c.Offset(0, 1)
    Else
        dest.Range("E" & theRow) = c.Offset(0, 1)
    End If
Next c

exit_here:
Set rng = Nothing
Set rng2 = Nothing
Set dest = Nothing
Set pw = Nothing
Exit Sub

err_trap:
If Err.Number = 91 Then
    theRow = lastRow2 + 1
    lacking = True
    Resume Next
Else
    MsgBox "Error in transferOperations. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Sub finalize()
Dim rng As Range
Dim lastRow As Long
Dim rs As ADODB.Recordset

On Error GoTo err_trap

With ThisWorkbook.Sheets("MWG")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("H2:H" & lastRow)
    For Each c In rng
        c.Formula = "=G" & c.row & "- (C" & c.row & "+D" & c.row & "-E" & c.row & "+F" & c.row & ")"
        c.Offset(0, 2).Formula = "=ABS(G" & c.row & "- (C" & c.row & "+D" & c.row & "-E" & c.row & "+F" & c.row & "))"
        If c.Offset(0, -6) = "" Then
            SQL = SQL & c.Offset(0, -7) & ","
        End If
    Next c
    If Len(SQL) > 0 Then
        SQL = "SELECT zfinIndex, zfinName FROM tbZfin WHERE zfinIndex IN (" & Left(SQL, Len(SQL) - 1) & ")"
        updateConnection
        Set rs = New ADODB.Recordset
        rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
        Do While Not rs.EOF
            theRow = .Range("A2:A" & lastRow).Find(rs.Fields("zfinIndex"), searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
            .Range("B" & theRow) = rs.Fields("zfinName")
            rs.MoveNext
        Loop
    End If
    Set rng = .Range("A1:J" & lastRow)
    rng.Sort Key1:=.Range("J1"), order1:=xlDescending, header:=xlYes
    Set rng = .Range("A1:I" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
End With

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then
        rs.Close
    End If
    Set rng = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in finalize. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub


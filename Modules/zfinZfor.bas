Attribute VB_Name = "zfinZfor"
'Public adoConn As ADODB.Connection

'Public Sub updateConnection()
'
'If Not adoConn Is Nothing Then
'    If adoConn.State = 0 Then
'        adoConn.Open ConnectionString
'        adoConn.CommandTimeout = 90
'    End If
'Else
'    Set adoConn = New ADODB.Connection
'    adoConn.Open ConnectionString
'    adoConn.CommandTimeout = 90
'End If
'End Sub

'Public Sub closeConnection()
'
'If Not adoConn Is Nothing Then
'    If adoConn.State = 1 Then
'        adoConn.Close
'    End If
'    Set adoConn = Nothing
'End If
'End Sub

Public Sub formatZfinZfor()
With ThisWorkbook.Sheets("ZFIN-ZFOR")
    .Cells.clear
    .Range("A1:B1").Merge
    .Range("C1:D1").Merge
    .Range("E1:E2").Merge
    .Range("F1:F2").Merge
    .Range("G1:G2").Merge
    .Range("A1") = "ZFOR"
    .Range("A2") = "Index"
    .Range("B2") = "Description"
    .Range("C1") = "ZFIN"
    .Range("C2") = "Index"
    .Range("D2") = "Description"
    .Range("E1") = "SCADA [kg]"
    .Range("F1") = "PW [kg]"
    .Range("G1") = "Difference [kg]"
    .Range("A1:G2").Font.Bold = True
    .Range("A1:G2").HorizontalAlignment = xlCenter
End With
End Sub

Function sumScada() As Variant()
Dim sc As Worksheet
Dim sca() As Variant
Dim i As Integer
Dim ind As Long
Dim val As Double
Dim bool As Boolean

ReDim sca(2, 0) As Variant

Set sc = ThisWorkbook.Sheets("SCADA")

ind = sc.Range("K2")
sca(0, 0) = ind
val = WorksheetFunction.SumIf(sc.Range("K2:K1000"), ind, sc.Range("P2:P1000"))
sca(1, 0) = val
sca(2, 0) = sc.Range("L2")

For i = 3 To 1000
    ind = sc.Range("K" & i)
    If ind = 0 Then
        Exit For
    Else
        'check if already in array
            bool = False
            For n = LBound(sca, 2) To UBound(sca, 2)
                If sca(0, n) = ind Then
                    bool = True
                    Exit For
                End If
            Next n
            If bool = False Then
                ReDim Preserve sca(2, UBound(sca, 2) + 1) As Variant
                sca(0, UBound(sca, 2)) = ind
                val = WorksheetFunction.SumIf(sc.Range("K2:K1000"), ind, sc.Range("P2:P1000"))
                sca(1, UBound(sca, 2)) = val
                sca(2, UBound(sca, 2)) = sc.Range("L" & i)
            End If
    End If
Next i

sumScada = sca

Set sc = Nothing

End Function

Public Sub scadaVsPw(control As IRibbonControl)
Dim sc() As Variant
Dim i As Integer
Dim pw As Worksheet
Dim dest As Worksheet
Dim ind As Long
Dim rs As ADODB.Recordset
Dim sqlStr As String
Dim rng As Range
Dim lastRow As Long
Dim theRow As Long
Dim found As Boolean
Dim n As Integer
Dim g As Integer
Dim rec As Integer

On Error GoTo err_trap

formatZfinZfor

sc = sumScada

updateConnection

Set rs = New ADODB.Recordset
Set pw = ThisWorkbook.Sheets("QGUAR")
Set dest = ThisWorkbook.Sheets("ZFIN-ZFOR")


lastRow = pw.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
Set rng = pw.Range("A3:A" & lastRow)

n = 3
For i = LBound(sc, 2) To UBound(sc, 2)
    ind = sc(0, i)
    g = 0
    sqlStr = "SELECT nam.zfinIndex, nam.zfinName FROM tbZfin nam RIGHT JOIN tbZFinZfor zz ON nam.zfinId = zz.zfinId " _
        & "WHERE zforId = (SELECT zfinId FROM tbZfin WHERE zfinIndex = " & ind & ")"
    rs.Open sqlStr, adoConn, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveFirst
        dest.Range("A" & n) = ind
        dest.Range("E" & n) = sc(1, i)
        dest.Range("B" & n) = sc(2, i)
        rec = rs.RecordCount
        Do Until rs.EOF
            found = True
            theRow = rng.Find(rs.Fields(0).Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
            If found Then
                dest.Range("C" & n + g) = pw.Range("A" & theRow)
                dest.Range("F" & n + g) = pw.Range("D" & theRow)
                dest.Range("D" & n + g) = rs.Fields(1).Value
                g = g + 1
            End If
            rs.MoveNext
        Loop
        If g = 0 And rec > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
                dest.Range("C" & n + g) = rs.Fields(0).Value
                dest.Range("D" & n + g) = rs.Fields(1).Value
                g = g + 1
                rs.MoveNext
            Loop
        End If
        dest.Range("G" & n).Formula = "=E" & n & "-SUM(F" & n & ":F" & n + g - 1 & ")"
        If Abs(dest.Range("G" & n).Value) >= 400 Then dest.Range("G" & n).Interior.Color = vbRed
        If g > 1 Then
            dest.Range("A" & n & ":A" & n + g - 1).Merge
            dest.Range("B" & n & ":B" & n + g - 1).Merge
            dest.Range("E" & n & ":E" & n + g - 1).Merge
            dest.Range("G" & n & ":G" & n + g - 1).Merge
        End If
    End If
    If g = 0 Then
        n = n + 1
    Else
        n = n + g
    End If
    rs.Close
Next i
pwVsScada
finishMe

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Set rng = Nothing
Set pw = Nothing
Set dest = Nothing
closeConnection
Exit Sub

err_trap:
If Err.Number = 91 Then
    found = False
    Resume Next
Else
    MsgBox "Error in scadaVsPw. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Sub finishMe()
Dim lastRow As Long

With ThisWorkbook.Sheets("ZFIN-ZFOR")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A1:G" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("A1:G2")
    rng.Interior.ColorIndex = 15
    Set rng = Nothing
End With
End Sub

Sub pwVsScada()
Dim rng As Range
Dim rng2 As Range
Dim c As Range
Dim lastRow As Long
Dim lastRow2 As Long
Dim pw As Worksheet
Dim dest As Worksheet
Dim rs As ADODB.Recordset
Dim sqlStr As String
Dim ind As Long
Dim conStr As String
Dim theRow As Long

On Error GoTo err_trap

updateConnection

Set rs = New ADODB.Recordset

Set pw = ThisWorkbook.Sheets("QGUAR")
Set dest = ThisWorkbook.Sheets("ZFIN-ZFOR")

lastRow = pw.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
Set rng = pw.Range("A3:A" & lastRow)
lastRow2 = dest.Range("C:C").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
Set rng2 = dest.Range("C3:C" & lastRow2)

For Each c In rng
    If WorksheetFunction.CountIf(rng2, c.Value) = 0 Then
        ind = c.Value
        conStr = conStr & ind & ","
        sqlStr = "SELECT nam.zfinIndex, nam.zfinName FROM tbZfin nam RIGHT JOIN tbZFinZfor zz ON nam.zfinId = zz.zforId " _
        & "WHERE zz.zfinId = (SELECT zfinId FROM tbZfin WHERE zfinIndex = " & ind & ")"
        rs.Open sqlStr, adoConn, adOpenKeyset, adLockOptimistic
        If Not rs.EOF Then
            rs.MoveFirst
            lastRow2 = lastRow2 + 1
            dest.Range("A" & lastRow2) = rs.Fields(0).Value
            dest.Range("B" & lastRow2) = rs.Fields(1).Value
            dest.Range("C" & lastRow2) = c.Value
            dest.Range("E" & lastRow2) = c.Offset(0, 3)
            dest.Range("G" & lastRow2).Formula = "=E" & lastRow2 & "-F" & lastRow2
            If Abs(dest.Range("G" & lastRow2).Value) >= 400 Then dest.Range("G" & lastRow2).Interior.Color = vbRed
            Set rng2 = dest.Range("C3:C" & lastRow2)
        End If
        rs.Close
    End If
Next c

If Len(conStr) > 0 Then
    conStr = Left(conStr, Len(conStr) - 1)
    sqlStr = "SELECT zfinIndex, zfinName FROM tbZfin WHERE zfinIndex IN (" & conStr & ");"
    rs.Open sqlStr, adoConn, adOpenKeyset, adLockOptimistic
    Do While Not rs.EOF
        theRow = rng2.Find(rs.Fields(0).Value, searchorder:=xlByRows, SearchDirection:=xlPrevious, Lookat:=xlWhole).row
        dest.Range("D" & theRow) = rs.Fields(1).Value
        rs.MoveNext
    Loop
End If

exit_here:
Set rng = Nothing
Set rng2 = Nothing
Set pw = Nothing
Set dest = Nothing
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
closeConnection
Exit Sub

err_trap:
MsgBox "Error in pwVsScada. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here


End Sub


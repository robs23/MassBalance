Attribute VB_Name = "Scader"
Public conn As ADODB.Connection
'connection to SCADA
Public npdConn As ADODB.Connection
'connection to NPD

Public Sub importFromScada(StartDate As Date, endDate As Date, Optional roaster As Variant, Optional blends As Variant, Optional exclude As Variant)
    
'Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim rcrds As ADODB.Recordset
'Set cmd = New ADODB.Command
Dim blendString As String
Dim excludeString As String
Dim r3000() As Double
Dim r4000() As Double
Dim r3 As Double
Dim r4 As Double
Dim fin As Double
Dim i As Integer

connectScada
ThisWorkbook.Sheets("SCADA").Cells.clear
'    If conn.State = adStateOpen Then
'        MsgBox "Connection successful"
'    End If
    
    'conn.Close
'    Set conn = Nothing


'End With

   'przykladowe zapytanie
   'sqlstr = "select * FROM Zlecenia;"
'   SQLstr = "select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, w.READJUSTED_BY, L.Nazwisko, L.Imie" _
'& " from ZLECENIA_PALONA as z Join ZLECENIAWARTOSCI as w " _
'& " JOIN ZLECENIA as zl on (w.IDZLECENIE = zl.IDZLECENIE)" _
' & " ON (z.IDZLECENIE = w.IDZLECENIE)" _
'& " LEFT OUTER JOIN LOGINS as L on (w.READJUSTED_BY = ID_LOGINS)" _
'& "Where z.DTZAPIS > '01-06-2016' " _
' & " ORDER BY z.DTZAPIS;"
    
 If Not IsMissing(blends) Then
     If Not isArrayEmpty(blends) Then
         blendString = " AND ("
         For i = LBound(blends) To UBound(blends)
             If i = LBound(blends) Then
                 blendString = blendString & "zl.MaterialNumber = " & blends(i)
             Else
                 blendString = blendString & " OR zl.MaterialNumber = " & blends(i)
             End If
         Next i
         blendString = blendString & ")"
     End If
 End If
 If Not IsMissing(exclude) Then
     If Not isArrayEmpty(exclude) Then
         excludeString = " AND ("
         For i = LBound(exclude) To UBound(exclude)
             If i = LBound(exclude) Then
                 excludeString = excludeString & "zl.MaterialNumber <> " & exclude(i)
             Else
                 excludeString = excludeString & " AND zl.MaterialNumber <> " & exclude(i)
             End If
         Next i
         excludeString = excludeString & ")"
     End If
 End If

 sqlStr = "select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT" _
 & " from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)" _
 & " Where (z.DTZAPIS Between ('" & CStr(StartDate) & "') AND ('" & CStr(endDate) & "'))"
 If Not IsMissing(roaster) Then sqlStr = sqlStr & " AND z.NUMERPIECA = " & roaster
 If blendString <> "" Then sqlStr = sqlStr & blendString
 If excludeString <> "" Then sqlStr = sqlStr & excludeString
 sqlStr = sqlStr & " ORDER BY z.DTZAPIS;"

'
'wykonanie zapytania i przypisanie wyniku do zmiennej rekordow
Set rcrds = conn.Execute(sqlStr)

i = 1
With ThisWorkbook.Sheets("SCADA")
     'zapisywanie wyniku zapytania w arkuszu - iteracja zestawu rekordow
     .Cells(i, 1) = "Piec"
     .Cells(i, 2) = "Kawa zielona"
     .Cells(i, 3) = "Uprażono"
     .Cells(i, 4) = "Data"
     .Cells(i, 5) = "Zlecenie"
     .Cells(i, 6) = "ZFOR"
     .Cells(i, 7) = "Nazwa"
     .Cells(i, 8) = "Ubytek [%]"
     i = 2
     Do While Not rcrds.EOF
         .Cells(i, 1) = rcrds("NUMERPIECA")
         .Cells(i, 2) = rcrds("SUMA_ZIELONEJ")
         .Cells(i, 3) = rcrds("ILOSC_PALONA")
         .Cells(i, 4) = rcrds("DTZAPIS")
        If rcrds("NUMERPIECA") = 3000 Then
            r3 = r3 + rcrds("ILOSC_PALONA")
        ElseIf rcrds("NUMERPIECA") = 4000 Then
            r4 = r4 + rcrds("ILOSC_PALONA")
        End If
            
         .Cells(i, 4).NumberFormat = "dd-mm-yyyy hh:mm:ss"
         .Cells(i, 5) = CLng(rcrds("OrderNumber"))
         .Cells(i, 6) = CLng(rcrds("MaterialNumber"))
         .Cells(i, 7) = rcrds("NAZWARECEPT")
         If Not IsNull(rcrds("ILOSC_PALONA")) And Not IsNull(rcrds("SUMA_ZIELONEJ")) Then
              If rcrds("ILOSC_PALONA") <> 0 And rcrds("SUMA_ZIELONEJ") <> 0 Then
                .Cells(i, 8) = 1 - (rcrds("ILOSC_PALONA") / rcrds("SUMA_ZIELONEJ"))
                .Cells(i, 8).NumberFormat = "0.00%"
                If rcrds("MaterialNumber") = 34005471 Or rcrds("MaterialNumber") = 34001130 Then
                  fin = fin + rcrds("ILOSC_PALONA")
                  If rcrds("NUMERPIECA") = 3000 Then
                      If isArrayEmpty(r3000) Then
                          ReDim r3000(0) As Double
                          r3000(0) = .Cells(i, 8) * 100
                      Else
                          ReDim Preserve r3000(UBound(r3000) + 1) As Double
                          r3000(UBound(r3000)) = .Cells(i, 8) * 100
                      End If
                  ElseIf rcrds("NUMERPIECA") = 4000 Then
                      If isArrayEmpty(r4000) Then
                          ReDim r4000(0) As Double
                          r4000(0) = .Cells(i, 8) * 100
                      Else
                          ReDim Preserve r4000(UBound(r4000) + 1) As Double
                          r4000(UBound(r4000)) = .Cells(i, 8) * 100
                      End If
                  End If
                End If
              End If
         End If
         rcrds.MoveNext
         i = i + 1
     Loop
 End With
'zakonczenie polaczenia
rcrds.Close

summarize
blendsByRoaster
formatMe
uploadBlendLoss
If ThisWorkbook.Sheets("BM").Range("L4") = "" Then ThisWorkbook.Sheets("BM").Range("L4") = r3
If ThisWorkbook.Sheets("BM").Range("L5") = "" Then ThisWorkbook.Sheets("BM").Range("L5") = fin

finezjaGraph r3000, r4000
conn.Close
Set rcrds = Nothing
Set conn = Nothing
'If UserForm1.cboxGraph.value = True Then
'    createGraph
'End If
End Sub


Public Function isArrayEmpty(parArray As Variant, Optional dimension As Variant) As Boolean
'Returns true if:
'  - parArray is not an array
'  - parArray is a dynamic array that has not been initialised (ReDim)
'  - parArray is a dynamic array has been erased (Erase)

  If IsArray(parArray) = False Then isArrayEmpty = True
  On Error Resume Next
    If IsMissing(dimension) Then
        If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False
    Else
        If UBound(parArray, dimension) < LBound(parArray, dimension) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False
    End If
End Function

Public Sub createGraph()
Dim i As Integer
Dim lineChart As ChartObject
Dim chD() As Variant 'chartData
Dim found As Boolean
Dim n As Integer
Dim q As Integer
Dim blend3() As Long 'blends on RN3000
Dim blend4() As Long 'blends on RN4000
Dim b As Long 'single blend
Dim r As Long 'roaster
Dim rn4000max As Double
Dim rn4000min As Double
Dim rn3000max As Double
Dim rn3000min As Double
Dim x() As Long
Dim y() As Double
Dim row As Long
Dim ws As String

ws = "SCADA"

For Each lineChart In ThisWorkbook.Sheets(ws).ChartObjects
    lineChart.Delete
Next lineChart

For i = 2 To 10000
    If ThisWorkbook.Sheets(ws).Cells(i, 6) > 0 Then
        r = ThisWorkbook.Sheets(ws).Cells(i, 1)
        b = ThisWorkbook.Sheets(ws).Cells(i, 6)
        If CLng(b) <> 34005471 And CLng(b) <> 34001130 Then
            If r = 3000 Then
                If isArrayEmpty(blend3) Then
                    ReDim blend3(0) As Long
                    blend3(0) = b 'blend
                Else
                    found = False
                    For n = LBound(blend3) To UBound(blend3)
                        If blend3(n) = b Then
                            found = True
                            Exit For
                        End If
                    Next n
                    If found = False Then
                        ReDim Preserve blend3(UBound(blend3) + 1) As Long
                        blend3(UBound(blend3)) = b
                    End If
                End If
            ElseIf r = 4000 Then
                If isArrayEmpty(blend4) Then
                    ReDim blend4(0) As Long
                    blend4(0) = b 'blend
                Else
                    found = False
                    For n = LBound(blend4) To UBound(blend4)
                        If blend4(n) = b Then
                            found = True
                            Exit For
                        End If
                    Next n
                    If found = False Then
                        ReDim Preserve blend4(UBound(blend4) + 1) As Long
                        blend4(UBound(blend4)) = b
                    End If
                End If
            End If
        End If
    Else
        Exit For
    End If
Next i

If Not isArrayEmpty(blend3) Then
    rn3000max = 0
    rn3000min = 50
    Set rng = ThisWorkbook.Sheets(ws).Range("J60:R80")
    Set lineChart = ThisWorkbook.Sheets(ws).ChartObjects.Add(Left:=rng.Left, Width:=rng.Width, Top:=rng.Top, Height:=rng.Height)
    With lineChart
        .Chart.ChartWizard Gallery:=xlLine, HasLegend:=True, title:="RN3000"
        .Name = "RN3000"
        For i = LBound(blend3) To UBound(blend3)
            For row = 2 To 10000
                If ThisWorkbook.Sheets(ws).Cells(row, 1) = 3000 And ThisWorkbook.Sheets(ws).Cells(row, 6) = blend3(i) Then
                    If isArrayEmpty(y) Then
                        ReDim y(0) As Double
                        y(0) = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 > rn3000max Then rn3000max = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 < rn3000min Then rn3000min = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                    Else
                        ReDim Preserve y(UBound(y) + 1) As Double
                        y(UBound(y)) = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 > rn3000max Then rn3000max = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 < rn3000min Then rn3000min = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                    End If
                ElseIf ThisWorkbook.Sheets(ws).Cells(row, 1) = 0 Then
                    With .Chart
                     .SeriesCollection.NewSeries
                        With .SeriesCollection(i + 1)
                                .Name = blend3(i) & " " & getBlendName(blend3(i))
                                .values = y
                                .format.Line.Weight = 1
        '                    .Values = Worksheets("Charts").Range(Cells(2, xx), Cells(pSpan + 2, xx))
        '                    .XValues = Worksheets("Charts").Range("A2:A" & pSpan + 2)
                            .MarkerStyle = xlMarkerStyleNone
                            Erase y
        '                    .ApplyDataLabels
        ''                    .DataLabels.Select
                        End With
                        If rn3000min < 10 Then
                            rn3000min = 10
                        Else
                            rn3000min = rn3000min - 1
                        End If
                        If rn3000max > 20 Then
                            rn3000max = 20
                        Else
                            rn3000max = rn3000max + 1
                        End If
                        .Axes(xlValue).MinimumScale = Int(rn3000min)
                        .Axes(xlValue).MaximumScale = Int(rn3000max)
                    End With
                    Exit For
                End If
            Next row
        Next i
    End With
    Set lineChart = Nothing
End If
If Not isArrayEmpty(blend4) Then
    Set rng = ThisWorkbook.Sheets(ws).Range("J85:R105")
    Set lineChart = ThisWorkbook.Sheets(ws).ChartObjects.Add(Left:=rng.Left, Width:=rng.Width, Top:=rng.Top, Height:=rng.Height)
    With lineChart
        .Chart.ChartWizard Gallery:=xlLine, HasLegend:=True, title:="RN4000"
        .Name = "RN4000"
        For i = LBound(blend4) To UBound(blend4)
            For row = 2 To 10000
                If ThisWorkbook.Sheets(ws).Cells(row, 1) = 4000 And ThisWorkbook.Sheets(ws).Cells(row, 6) = blend4(i) Then
                    If isArrayEmpty(y) Then
                        ReDim y(0) As Double
                        y(0) = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 > rn4000max Then rn4000max = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 < rn4000min Then rn4000min = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                    Else
                        ReDim Preserve y(UBound(y) + 1) As Double
                        y(UBound(y)) = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 > rn4000max Then rn4000max = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                        If ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 <> 0 And ThisWorkbook.Sheets(ws).Cells(row, 8) * 100 < rn4000min Then rn4000min = ThisWorkbook.Sheets(ws).Cells(row, 8) * 100
                    End If
                ElseIf ThisWorkbook.Sheets(ws).Cells(row, 1) = 0 Then
                    With .Chart
                     .SeriesCollection.NewSeries
                        With .SeriesCollection(i + 1)
                                .Name = blend4(i) & " " & getBlendName(blend4(i))
                                .values = y
                                .format.Line.Weight = 1
        '                    .Values = Worksheets("Charts").Range(Cells(2, xx), Cells(pSpan + 2, xx))
        '                    .XValues = Worksheets("Charts").Range("A2:A" & pSpan + 2)
                            .MarkerStyle = xlMarkerStyleNone
                            Erase y
        '                    .ApplyDataLabels
        ''                    .DataLabels.Select
                        End With
                        If rn4000min < 10 Then
                            rn4000min = 10
                        Else
                            rn4000min = rn4000min - 1
                        End If
                        If rn4000max > 20 Then
                            rn4000max = 20
                        Else
                            rn4000max = rn4000max + 1
                        End If
                        .Axes(xlValue).MinimumScale = Int(rn4000min)
                        .Axes(xlValue).MaximumScale = Int(rn4000max)
                    End With
                    Exit For
                End If
            Next row
        Next i
    End With
    Set lineChart = Nothing
End If
'If ThisWorkbook.sheets(ws).ChartObjects.Count > 0 Then
'    ThisWorkbook.sheets(ws).ChartObjects.Delete
'End If
'
'n = 0
'
'For i = 2 To 10000
'    If ThisWorkbook.sheets(ws).Cells(i, 6) = blend Then
'    If isArrayEmpty(chD) Then
'        ReDim chD(0, 3) As Variant
'        chD(0, 0) = n
'    Else
'        ReDim chD(UBound(chD, 1) + 1, 3) As Variant
'    End If
'Next i
'
'

'


End Sub



Public Sub applyCustomPointLabels(seriesName As String, ch As ChartObject, Optional values As Variant)
Dim srs As Series, rng As Range, lbl As DataLabel
Dim iLbl As Long, nLbls As Long
Dim pnt As Point

Set srs = ch.Chart.SeriesCollection(seriesName)

If Not srs Is Nothing Then
    For Each pnt In srs.Points
        pnt.HasDataLabel = True
        Set lbl = pnt.DataLabel
        With lbl
            .Text = "Dupa"
            .Position = xlLabelPositionRight
        End With
        Set lbl = Nothing
    Next pnt
End If
Set srs = Nothing
End Sub

Public Sub PROD_PODSUMUJ() 'rng As Range, id As Integer, value As Integer)
Dim i As Integer
Dim idCell As Range
Dim valCell As Range
Dim val As Double
Dim index As Variant
Dim oW As Worksheet
Dim Target As Range
Dim n As Integer

Target = Application.ActiveCell
Set oW = rng.Worksheet
rng.Sort Key1:=oW.Cells(rng.Column + Id - 1), order1:=xlAscending
For i = rng.row To rng.Height + rng.row
    idCell = oW.Cells(i, rng.Column + Id - 1)
    If Not IsEmpty(idCell) Then
        valCell = oW.Cells(i, rng.Column + Value - 1)
        If IsNumeric(valCell) Then
            If idCell.Value = index Then
                val = val + idCell.Value
            Else
                Target.Worksheet.Cells(Target.row + n, Target.Column) = index
                Target.Worksheet.Cells(Target.row + n, Target.Column + 1) = val
                index = idCell.Value
                val = 0
                n = n + 1
            End If
        End If
        
    End If
Next i

End Sub

Public Function getBlendName(blend As Long) As String
Dim i As Integer
Dim ws As String

ws = "SCADA"

For i = 2 To 10000
    If ThisWorkbook.Sheets(ws).Cells(i, 6) = blend Then
        getBlendName = ThisWorkbook.Sheets(ws).Cells(i, 7)
        Exit For
    ElseIf ThisWorkbook.Sheets(ws).Cells(i, 6) = "" Then
        Exit For
    End If
Next i
End Function

Sub summarize()
Dim rs As ADODB.Recordset
Dim i As Integer
Dim SQL As String

connectScada

'sql = "SELECT DISTINCT sum(z.SUMA_ZIELONEJ) as sumaZielonej, sum(z.ILOSC_PALONA) as sumaPalonej, min(z.DTZAPIS) as minData, max(z.DTZAPIS) as maxData,zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT " _
'    & "FROM ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE) " _
'    & "WHERE (z.DTZAPIS Between ('" & ThisWorkbook.CustomDocumentProperties("roastingFrom") & "') AND ('" & ThisWorkbook.CustomDocumentProperties("roastingTo") & "')) " _
'    & "GROUP BY zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT " _
'    & "ORDER BY min(z.DTZAPIS);"

SQL = "SELECT rD.OrderNumber, rd.MaterialNumber, rd.NAZWARECEPT, Min(rd.DTZAPIS) as minData, Max(rd.DTZAPIS) as maxData, sum(rd.SUMA_ZIELONEJ) as sumaZielonej,sum(rd.ILOSC_PALONA) As sumaPalonej " _
    & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE) Where (z.DTZAPIS Between ('" & ThisWorkbook.CustomDocumentProperties("roastingFrom") & "') AND ('" & ThisWorkbook.CustomDocumentProperties("roastingTo") & "'))) as rD " _
    & "GROUP BY rd.OrderNumber,rd.MaterialNumber,rd.NAZWARECEPT ORDER BY minData;"
    
'Set rs = conn.Execute(sql)
Set rs = CreateObject("adodb.recordset")
rs.Open SQL, conn
If Not rs.EOF Then
    rs.MoveFirst
    i = 1
    With ThisWorkbook.Sheets("SCADA")
        .Range("J" & i) = "Numer zlecenia"
        .Range("K" & i) = "Index ZFORa"
        .Range("L" & i) = "Nazwa ZFORa"
        .Range("M" & i) = "Początek"
        .Range("N" & i) = "Koniec"
        .Range("O" & i) = "Kawa zielona"
        .Range("P" & i) = "Kawa uprażona"
        .Range("Q" & i) = "Strata"
    End With
    Do Until rs.EOF
        i = i + 1
        With ThisWorkbook.Sheets("SCADA")
            .Range("J" & i) = CLng(rs.Fields("OrderNumber")) '
            .Range("K" & i) = CLng(rs.Fields("MaterialNumber")) '
            .Range("L" & i) = rs.Fields("NAZWARECEPT") '
            .Range("M" & i) = rs.Fields("minData") '
            .Range("M" & i).NumberFormat = "dd-mm-yyyy hh:mm:ss"
            .Range("N" & i) = rs.Fields("maxData")
            .Range("N" & i).NumberFormat = "dd-mm-yyyy hh:mm:ss"
            .Range("O" & i) = rs.Fields("sumaZielonej")
            .Range("P" & i) = rs.Fields("sumaPalonej")
            If Not IsNull(rs.Fields("sumaPalonej")) And Not IsNull(rs.Fields("sumaZielonej")) Then
                .Range("Q" & i) = 1 - (rs.Fields("sumaPalonej") / rs.Fields("sumaZielonej"))
                .Range("Q" & i).NumberFormat = "0.00%"
            End If
        End With
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

End Sub

Public Sub blendsByRoaster()
Dim rs As ADODB.Recordset
Dim i As Integer
Dim SQL As String

connectScada

'sql = "SELECT DISTINCT sum(z.SUMA_ZIELONEJ) as sumaZielonej, sum(z.ILOSC_PALONA) as sumaPalonej, min(z.DTZAPIS) as minData, max(z.DTZAPIS) as maxData,zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT " _
'    & "FROM ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE) " _
'    & "WHERE (z.DTZAPIS Between ('" & ThisWorkbook.CustomDocumentProperties("roastingFrom") & "') AND ('" & ThisWorkbook.CustomDocumentProperties("roastingTo") & "')) " _
'    & "GROUP BY zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT " _
'    & "ORDER BY min(z.DTZAPIS);"

SQL = "SELECT rd.NUMERPIECA,rd.MaterialNumber,rd.NAZWARECEPT,Min(rd.DTZAPIS) as minData,Max(rd.DTZAPIS) as maxData,sum(rd.SUMA_ZIELONEJ) as sumaZielonej, " _
    & "sum(rd.ILOSC_PALONA) As sumaPalonej from (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.MaterialNumber, zl.NAZWARECEPT " _
    & "from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE) " _
    & "Where (z.DTZAPIS Between ('" & ThisWorkbook.CustomDocumentProperties("roastingFrom") & "') AND ('" & ThisWorkbook.CustomDocumentProperties("roastingTo") & "'))) as rD " _
    & "GROUP BY rd.NUMERPIECA, rd.MaterialNumber, rd.NAZWARECEPT ORDER BY rd.MaterialNumber;"
    
'Set rs = conn.Execute(sql)
Set rs = CreateObject("adodb.recordset")
rs.Open SQL, conn
If Not rs.EOF Then
    rs.MoveFirst
    i = 1
    With ThisWorkbook.Sheets("SCADA")
        .Range("S" & i) = "ZFOR"
        .Range("T" & i) = "Nazwa"
        .Range("U" & i) = "Piec"
        .Range("V" & i) = "Kawa zielona"
        .Range("W" & i) = "Kawa uprażona"
        .Range("X" & i) = "Strata"
    End With
    Do Until rs.EOF
        i = i + 1
        With ThisWorkbook.Sheets("SCADA")
            .Range("S" & i) = CLng(rs.Fields("MaterialNumber")) '
            .Range("T" & i) = rs.Fields("NAZWARECEPT") '
            .Range("U" & i) = rs.Fields("NUMERPIECA") '
            .Range("V" & i) = rs.Fields("sumaZielonej")
            .Range("W" & i) = rs.Fields("sumaPalonej")
            If Not IsNull(rs.Fields("sumaPalonej")) And Not IsNull(rs.Fields("sumaZielonej")) Then
                .Range("X" & i) = 1 - (rs.Fields("sumaPalonej") / rs.Fields("sumaZielonej"))
                .Range("X" & i).NumberFormat = "0.00%"
            End If
        End With
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

End Sub

Public Sub connectScada()
'Dim cmd As ADODB.Command
'Set cmd = New ADODB.Command

If conn Is Nothing Then
    Set conn = New ADODB.Connection
    conn.Provider = "SQLOLEDB"
    conn.ConnectionString = ScadaConnectionString
    conn.Open
    conn.CommandTimeout = 90
Else
    If conn.State = adStateClosed Then
        Set conn = New ADODB.Connection
        conn.Provider = "SQLOLEDB"
        conn.ConnectionString = ScadaConnectionString
        conn.Open
        conn.CommandTimeout = 90
    End If
End If


End Sub

Public Sub connectNpd()
'Dim cmd As ADODB.Command
'Set cmd = New ADODB.Command

If npdConn Is Nothing Then
    Set npdConn = New ADODB.Connection
    npdConn.Provider = "SQLOLEDB"
    npdConn.ConnectionString = ConnectionString
    npdConn.Open
    npdConn.CommandTimeout = 90
Else
    If npdConn.State = adStateClosed Then
        Set npdConn = New ADODB.Connection
        npdConn.Provider = "SQLOLEDB"
        npdConn.ConnectionString = ConnectionString
        npdConn.Open
        npdConn.CommandTimeout = 90
    End If
End If


End Sub

Sub bringBeans(products As String)
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String
Dim i As Integer

On Error GoTo exit_here


Set conn = New ADODB.Connection
conn.Open ConnectionString
conn.CommandTimeout = 90

iStr = "INSERT INTO tbZfor [zforIndex, zforName] VALUES ["
sSql = "SELECT zforIndex, [beans?] FROM tbZforProperties JOIN tbZfor on tbZfor.zforId = tbZforProperties.zforId WHERE tbZfor.zforIndex IN (" & products & ");"
Set rs = conn.Execute(sSql)

If Not rs.EOF Then
    rs.MoveFirst
    If rs.Fields("beans?") = 0 Then
        isZforBeans = False
    Else
        isZforBeans = True
    End If
Else
    isZforBeans = Null
End If
rs.Close

exit_here:
conn.Close
Set rs = Nothing
Set conn = Nothing
End Sub

Sub finezjaGraph(r3 As Variant, r4 As Variant)
Dim ws As String
Dim rng As Range
Dim lineChart As ChartObject
ws = "SCADA"

If Not isArrayEmpty(r3) And Not isArrayEmpty(r4) Then
    Set rng = ThisWorkbook.Sheets(ws).Range("J110:R131")
    Set lineChart = ThisWorkbook.Sheets(ws).ChartObjects.Add(Left:=rng.Left, Width:=rng.Width, Top:=rng.Top, Height:=rng.Height)
    With lineChart
        .Chart.ChartWizard Gallery:=xlLine, HasLegend:=True, title:="Tydzień " & ThisWorkbook.CustomDocumentProperties("weekLoaded")
        .Name = "finezja"
        With .Chart
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                    .Name = "RN3000"
                    .values = r3
                    .format.Line.Weight = 1
'                    .Values = Worksheets("Charts").Range(Cells(2, xx), Cells(pSpan + 2, xx))
'                    .XValues = Worksheets("Charts").Range("A2:A" & pSpan + 2)
                .MarkerStyle = xlMarkerStyleNone

'                    .ApplyDataLabels
''                    .DataLabels.Select
            End With
            .SeriesCollection.NewSeries
            With .SeriesCollection(2)
                    .Name = "RN4000"
                    .values = r4
                    .format.Line.Weight = 1
'                    .Values = Worksheets("Charts").Range(Cells(2, xx), Cells(pSpan + 2, xx))
'                    .XValues = Worksheets("Charts").Range("A2:A" & pSpan + 2)
                .MarkerStyle = xlMarkerStyleNone

'                    .ApplyDataLabels
''                    .DataLabels.Select
            End With
            .Axes(xlValue).MinimumScale = 11.5
            .Axes(xlValue).MaximumScale = 16.5
        End With
    End With
    Set rng = Nothing
    Set lineChart = Nothing
End If


End Sub

Public Sub formatMe()
Dim rng As Range
Dim lastRow As Long

With ThisWorkbook.Sheets("SCADA")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A1:H" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("A1:H1")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
    
    lastRow = .Range("J:J").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("J1:Q" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("J1:Q1")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
    
    lastRow = .Range("S:S").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("S1:X" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("S1:X1")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
End With

Set rng = Nothing

End Sub

Sub uploadBlendLoss()
Dim rs As ADODB.Recordset
Dim lastRow As Long
Dim rng As Range
Dim m As Integer
Dim c As Range
Dim bmId As Integer
Dim isError As Boolean
Dim sqlStr As String
Dim Roasted As String
Dim green As String
Dim Id As Integer

On Error GoTo err_trap

updateConnection

SQL = "SELECT bmId FROM tbBM WHERE bmWeek = " & ThisWorkbook.CustomDocumentProperties("weekLoaded") & " AND bmYear = " & ThisWorkbook.CustomDocumentProperties("yearLoaded") & ";"
    
Set rs = CreateObject("adodb.recordset")
rs.Open SQL, adoConn
If rs.EOF Then
    MsgBox "There's no data for chosen week/year yet. Can't upload information about roasting losses until you create the period first", vbOKOnly + vbExclamation, "Period w" & ThisWorkbook.CustomDocumentProperties("weekLoaded") & "|" & ThisWorkbook.CustomDocumentProperties("yearLoaded") & " doesn't exist!"
Else
    rs.MoveFirst
    bmId = rs.Fields("bmId")
    rs.Close
    Set rs = Nothing
    createProducts ("zfor")
    With ThisWorkbook.Sheets("SCADA")
        lastRow = .Range("S:S").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

        Set rs = adoConn.Execute("DELETE FROM tbRoastingLoss WHERE bmId = " & bmId)

        Set rng = .Range("S2:S" & lastRow)
        For Each c In rng
            Roasted = 0
            green = 0
            If c.Offset(0, 2) = 3000 Then
                m = 13
            ElseIf c.Offset(0, 2) = 4000 Then
                m = 14
            End If
            If IsNull(c.Offset(0, 3)) Or c.Offset(0, 3) = 0 Then
                green = "0"
            Else
                green = Replace(c.Offset(0, 3), ",", ".")
            End If
            If IsNull(c.Offset(0, 4)) Or c.Offset(0, 4) = 0 Then
                Roasted = "0"
            Else
                Roasted = Replace(c.Offset(0, 4), ",", ".")
            End If
            isError = False
            sqlStr = "INSERT INTO tbRoastingLoss(zforId, machineId, greenCoffee, roastedCoffee, bmId) VALUES (" & products(CStr(c.Value)).prodId & ", " & m & ", " & green & ", " & Roasted & ", " & bmId & ");"
            If isError Then
                sqlStr = "INSERT INTO tbRoastingLoss(zforId, machineId, greenCoffee, roastedCoffee, bmId) VALUES (" & Id & ", " & m & ", " & green & ", " & Roasted & ", " & bmId & ");"
            End If
            Set rs = adoConn.Execute(sqlStr)
        Next c
    End With
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Set rng = Nothing
closeConnection
Exit Sub

err_trap:
If Err.Number = 5 Then
    'we've hit new product, create it
    sqlStr = "INSERT INTO tbZfin(zfinIndex, zfinName, zfinType, creationDate, createdBy) VALUES (" & CLng(c.Value) & ", '" & c.Offset(0, 1).Value & "', 'zfor', '" & Now & "', 43);SELECT SCOPE_IDENTITY() AS ID;"
    Set rs = adoConn.Execute(sqlStr)
    Id = rs.Fields(0).Value
    isError = True
    Resume Next
Else
    MsgBox "Error in uploadBlendLoss. Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub



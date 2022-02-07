VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} roastingHistory 
   Caption         =   "Update roasting history"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "roastingHistory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "roastingHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rFrom As Date
Private rTo As Date

Private Sub btnUpdate_Click()
Dim w0 As Integer
Dim y0 As Integer
Dim w1 As Integer
Dim y1 As Integer
Dim x As Integer
Dim SQL As String
Dim pStr As String 'periodString
Dim zStr As String 'zfin string
Dim title As String

ThisWorkbook.Sheets("Roasting history").Range("A3:J30000").Cells.clear
'For Each chrt In ThisWorkbook.Sheets("Roasting history").ChartObjects
'    chrt.Delete
'Next chrt

title = "Roasting loss for "

If verify Then
    Select Case cmbSummary
        Case Is = "Daily"
            pStr = "rD.DTZAPIS"
        Case Is = "Weekly"
            pStr = "CONVERT(nchar(4),YEAR(rD.DTZAPIS)) +'/' + CASE WHEN DATEPART(ISO_WEEK,rD.DTZAPIS) > 9 THEN CONVERT(nchar(2),DATEPART(ISO_WEEK,rD.DTZAPIS)) ELSE '0' + CONVERT(nchar(1),DATEPART(ISO_WEEK,rD.DTZAPIS)) END"
        Case Is = "Monthly"
            pStr = "CONVERT(nchar(4),YEAR(rD.DTZAPIS)) +'/' + CASE WHEN MONTH(rD.DTZAPIS) > 9 THEN CONVERT(nchar(2),MONTH(rD.DTZAPIS)) ELSE '0' + CONVERT(nchar(1),MONTH(rD.DTZAPIS)) END"
        Case Is = "Quarterly"
            pStr = "CONVERT(nchar(4),YEAR(rD.DTZAPIS)) +'/' + '0' + CONVERT(nchar(1),DATEPART(qq,rD.DTZAPIS))"
        Case Is = "Yearly"
            pStr = "CONVERT(nchar(4),YEAR(rD.DTZAPIS))"
    End Select
    
    If cmbBlend.Value = "All beans" Then
        zStr = getBeansOrGround("b")
        title = title & "beans"
    ElseIf cmbBlend.Value = "All ground" Then
        zStr = getBeansOrGround("g")
        title = title & "ground"
    ElseIf cmbBlend.Value = "All" Then
        zStr = getBeansOrGround("a")
        title = title & "all blends"
    Else
        zStr = cmbBlend.Value
        title = title & cmbBlend.Value
    End If
    
    If Me.cmbOptions.ListIndex = 0 Then
        w0 = 1
    '    y0 = year(Date)
        w1 = CInt(IsoWeekNumber(Date))
        y1 = year(Date)
        x = w1 - w0
        rFrom = DateSerial(year(Date), 1, 1)
        rTo = Date
        
        SQL = "SELECT " & pStr & " as Period, " _
            & "ROUND(sum(rd.SUMA_ZIELONEJ)/1000,1) as [TotalIn], " _
            & "ROUND(sum(rd.ILOSC_PALONA)/1000,1) As [TotalOut], " _
            & "ROUND(100*(1-(sum(rd.ILOSC_PALONA)/sum(rd.SUMA_ZIELONEJ))),2) as TotalLoss, " _
            & "ROUND(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000,1) as [r3In], " _
            & "ROUND(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000,1) as [r3Out], " _
            & "ROUND((1-(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000)/(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000))*100,2) as r3Loss, " _
            & "ROUND(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000,1) as [r4In], " _
            & "ROUND(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000,1) as [r4Out], " _
            & "ROUND((1-(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000)/(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000))*100,2) as r4Loss " _
            & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE) as rD " _
            & "WHERE rd.MaterialNumber IN (" & zStr & ") And year(rd.DTZAPIS) = " & year(Date) _
            & " GROUP BY " & pStr _
            & " ORDER BY Period;"
    
    ElseIf Me.cmbOptions.ListIndex = 1 Then
        If IsNumeric(Me.txtX.Value) Then
            If Me.txtX.Value < 1 Or Me.txtX.Value > 200 Then
                MsgBox "Podana wartość musi mieścić się w zakresie 1 - 200", vbOKOnly + vbInformation, "Niewłaściwa wartość"
            Else
                w1 = CInt(IsoWeekNumber(Date))
                y1 = year(Date)
    '            w0 = CInt(IsoWeekNumber(DateAdd("ww", -1 * Me.txtX.Value, Date)))
    '            y0 = year(DateAdd("ww", -1 * Me.txtX.Value, Date))
                x = Me.txtX.Value
                If cmbSummary.Value = "Weekly" Then
                    rFrom = DateAdd("ww", -1 * x, Date)
                ElseIf cmbSummary.Value = "Monthly" Then
                    rFrom = DateAdd("m", -1 * x, Date)
                ElseIf cmbSummary.Value = "Quarterly" Then
                    rFrom = DateAdd("q", -1 * x, Date)
                Else
                    rFrom = DateAdd("yyyy", -1 * x, Date)
                End If
                
                rTo = Date
                SQL = "SELECT " & pStr & " as Period, " _
                    & "ROUND(sum(rd.SUMA_ZIELONEJ)/1000,1) as [TotalIn], " _
                    & "ROUND(sum(rd.ILOSC_PALONA)/1000,1) As [TotalOut], " _
                    & "ROUND(100*(1-(sum(rd.ILOSC_PALONA)/sum(rd.SUMA_ZIELONEJ))),2) as TotalLoss, " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000,1) as [r3In], " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000,1) as [r3Out], " _
                    & "ROUND((1-(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000)/(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000))*100,2) as r3Loss, " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000,1) as [r4In], " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000,1) as [r4Out], " _
                    & "ROUND((1-(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000)/(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000))*100,2) as r4Loss " _
                    & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE) as rD " _
                    & "WHERE rd.MaterialNumber IN (" & zStr & ") And rd.DTZAPIS >= '" & rFrom & "'" _
                    & " GROUP BY " & pStr _
                    & " ORDER BY Period;"
            End If
        Else
            MsgBox "Podana wartość jest nienumeryczna. Podaj wartość z zakresu 1 - 200", vbOKOnly + vbInformation, "Niewłaściwa wartość"
        End If
    ElseIf Me.cmbOptions.ListIndex = 2 Then
        If IsDate(Me.dFrom.Value) And IsDate(Me.dTo.Value) Then
            If Me.dFrom.Value > Me.dTo.Value Then
                MsgBox "Początkowa data musi być mniejsza od daty końcowej", vbOKOnly + vbInformation, "Niewłaściwa wartość"
            Else
    '            w0 = IsoWeekNumber(Me.dFrom.Value)
    '            y0 = year(Me.dFrom.Value)
                x = DateDiff("ww", Me.dFrom.Value, Me.dTo.Value)
                w1 = IsoWeekNumber(Me.dTo.Value)
                y1 = year(Me.dTo.Value)
                SQL = "SELECT " & pStr & " as Period, " _
                    & "ROUND(sum(rd.SUMA_ZIELONEJ)/1000,1) as [TotalIn], " _
                    & "ROUND(sum(rd.ILOSC_PALONA)/1000,1) As [TotalOut], " _
                    & "ROUND(100*(1-(sum(rd.ILOSC_PALONA)/sum(rd.SUMA_ZIELONEJ))),2) as TotalLoss, " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000,1) as [r3In], " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000,1) as [r3Out], " _
                    & "ROUND((1-(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000)/(sum(CASE WHEN rD.NUMERPIECA=3000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000))*100,2) as r3Loss, " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000,1) as [r4In], " _
                    & "ROUND(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000,1) as [r4Out], " _
                    & "ROUND((1-(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.ILOSC_PALONA ELSE NULL END)/1000)/(sum(CASE WHEN rD.NUMERPIECA=4000 THEN rD.SUMA_ZIELONEJ ELSE NULL END)/1000))*100,2) as r4Loss " _
                    & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE) as rD " _
                    & "WHERE rd.MaterialNumber IN (" & zStr & ") And rd.DTZAPIS >= '" & Me.dFrom.Value & "' AND rd.DTZAPIS <='" & Me.dTo.Value & "'" _
                    & " GROUP BY " & pStr _
                    & " ORDER BY Period;"
            End If
        Else
            MsgBox "Oba pola powinny być wypełnione wartością w formacie daty", vbOKOnly + vbInformation, "Niewłaściwa wartość"
        End If
    End If
    
    If w1 > 0 And y1 > 0 And x > 0 Then
        bringHistory SQL, title & " " & LCase(cmbSummary.Value)
        ThisWorkbook.Sheets("Roasting history").Select
        Me.Hide
    End If
End If

End Sub

Private Function getBeansOrGround(theType As String) As String
Dim rs As ADODB.Recordset
Dim SQL As String
Dim out As String

On Error GoTo err_trap

connectNpd

If theType = "b" Then
    SQL = "SELECT z.zfinIndex FROM tbZfin z LEFT JOIN tbZfinProperties zp ON z.zfinId=zp.zfinId WHERE z.zfinType='zfor' AND zp.[beans?]<>0"
ElseIf theType = "g" Then
    SQL = "SELECT z.zfinIndex FROM tbZfin z LEFT JOIN tbZfinProperties zp ON z.zfinId=zp.zfinId WHERE z.zfinType='zfor' AND zp.[beans?]=0"
Else
    SQL = "SELECT z.zfinIndex FROM tbZfin z WHERE z.zfinType='zfor'"
End If

Set rs = New ADODB.Recordset
rs.Open SQL, npdConn, adOpenDynamic, adLockOptimistic

out = ""

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        out = out & rs.Fields("zfinIndex").Value & ","
        rs.MoveNext
    Loop
    out = Left(out, Len(out) - 1)
End If

exit_here:
getBeansOrGround = out
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Function

err_trap:
MsgBox "Error in getBeans. Description: " & Err.Description
Resume exit_here

End Function

Private Function verify() As Boolean
Dim bool As Boolean

bool = False

If IsNull(Me.cmbBlend) Then
    MsgBox "Choose ZFOR from drop-down list!", vbOKOnly + vbCritical, "No choice"
Else
    If IsNull(Me.cmbSummary) Then
        MsgBox "Choose summary type from drop-down list!", vbOKOnly + vbCritical, "No choice"
    Else
        bool = True
    End If
End If

verify = bool

End Function


Private Sub cmbOptions_Change()
If Me.cmbOptions.ListIndex = 0 Then
    Me.lX.Visible = False
    Me.txtX.Visible = False
    Me.dFrom.Visible = False
    Me.dTo.Visible = False
    Me.lFrom.Visible = False
    Me.lTo.Visible = False
ElseIf Me.cmbOptions.ListIndex = 1 Then
    Me.lX.Visible = True
    Me.txtX.Visible = True
    Me.dFrom.Visible = False
    Me.dTo.Visible = False
    Me.lFrom.Visible = False
    Me.lTo.Visible = False
ElseIf Me.cmbOptions.ListIndex = 2 Then
    Me.lX.Visible = False
    Me.txtX.Visible = False
    Me.dFrom.Visible = True
    Me.dTo.Visible = True
    Me.lFrom.Visible = True
    Me.lTo.Visible = True
End If
End Sub

Private Sub UserForm_Initialize()
Me.cmbOptions.clear
Me.cmbOptions.AddItem "This year"
Me.cmbOptions.AddItem "Last X periods"
Me.cmbOptions.AddItem "Date range"
Me.cmbOptions.ListIndex = 0

Me.cmbSummary.clear
Me.cmbSummary.AddItem "Weekly"
Me.cmbSummary.AddItem "Monthly"
Me.cmbSummary.AddItem "Quarterly"
Me.cmbSummary.AddItem "Yearly"
Me.cmbSummary.ListIndex = 0

fillBlends
End Sub

Private Sub fillBlends()
Dim rs As ADODB.Recordset
Dim SQL As String
Dim i As Integer

On Error GoTo err_trap

With cmbBlend
    .ColumnCount = 2
    .BoundColumn = 1
    .ColumnWidths = "2cm;2cm"
    .clear
End With

SQL = "SELECT zfinIndex, zfinName FROM tbZfin WHERE zfinType='zfor' ORDER BY zfinIndex"

updateConnection

'first add all beans / all ground option

cmbBlend.AddItem "All"
cmbBlend.AddItem "All beans"
cmbBlend.AddItem "All ground"

i = 3
Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        cmbBlend.AddItem rs.Fields("zfinIndex")
        cmbBlend.List(i, 1) = rs.Fields("zfinName")
        i = i + 1
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
MsgBox "Error in ""fillBlends"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub bringHistory(SQL As String, title As String)
Dim rs As ADODB.Recordset
Dim i As Integer
Dim sht As Worksheet
Dim isError As Boolean

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets("Roasting history")

connectScada

Set rs = New ADODB.Recordset
rs.Open SQL, conn, adOpenDynamic, adLockOptimistic

isError = False

If Not rs.EOF Then
    rs.MoveFirst
    i = 3
    Do Until rs.EOF
        sht.Cells(i, 1) = rs.Fields("Period")
        sht.Cells(i, 8) = rs.Fields("TotalIn")
        sht.Cells(i, 9) = rs.Fields("TotalOut")
        sht.Cells(i, 10) = rs.Fields("TotalLoss")
        sht.Cells(i, 2) = rs.Fields("r3In")
        sht.Cells(i, 3) = rs.Fields("r3Out")
        sht.Cells(i, 4) = rs.Fields("r3Loss")
        sht.Cells(i, 5) = rs.Fields("r4In")
        sht.Cells(i, 6) = rs.Fields("r4Out")
        sht.Cells(i, 7) = rs.Fields("r4Loss")
        i = i + 1
        rs.MoveNext
    Loop
    updateRoastingChart i - 1, title
Else
    isError = True
    MsgBox "No results for given period/blend", vbInformation + vbOKOnly, "No results"
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
'If Not isError Then createGraph
Exit Sub

err_trap:
isError = True
MsgBox "Error in ""BringHistory"" of RoastingHistory. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub updateRoastingChart(tot As Integer, title As String)
Dim lineChart As ChartObject
Dim srs As Series
'ThisWorkbook.Worksheets("Results").ChartObjects("Wykres 4").Name = "grpGreenCoffee"
For Each lineChart In ThisWorkbook.Worksheets("Roasting history").ChartObjects
    'Debug.Print "Wszystkie serie danych " & lineChart.Name
    lineChart.Chart.ChartTitle.Text = title
    With Worksheets("Roasting history")
        For Each srs In lineChart.Chart.SeriesCollection
            Select Case srs.Name
            Case Is = "Roasted on RN3000"
                srs.xValues = .Range("A3:A" & tot)
                srs.values = .Range("C3:C" & tot)
            Case Is = "Roasted on RN4000"
                srs.xValues = .Range("A3:A" & tot)
                srs.values = .Range("F3:F" & tot)
            Case Is = "RN3000 loss"
                srs.xValues = .Range("A3:A" & tot)
                srs.values = .Range("D3:D" & tot)
            Case Is = "RN4000 loss"
                srs.xValues = .Range("A3:A" & tot)
                srs.values = .Range("G3:G" & tot)
            Case Is = "Total loss"
                srs.xValues = .Range("A3:A" & tot)
                srs.values = .Range("J3:J" & tot)
            End Select
        Next srs
    End With
Next lineChart

End Sub

Private Sub createGraph()
Dim graph As clsGraph
Dim chrt As ChartObject
Dim i As Integer
Dim xVal As String
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("Roasting history")

Set graph = New clsGraph

graph.initialize "Roasting history " & Me.cmbSummary, "roastingGraph", xlLine, "Period", "Loss [%]", "Roasting history", sht.Range("L4:W26")

For i = 3 To 10000
    xVal = sht.Cells(i, 1)
    If Len(xVal) = 0 Then
        Exit For
    Else
        graph.append xVal, sht.Cells(i, 10)
        graph.append2nd xVal, sht.Cells(i, 4)
        graph.append3rd xVal, sht.Cells(i, 7)
    End If
Next i

graph.createChart
End Sub

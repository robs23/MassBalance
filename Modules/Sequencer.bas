Attribute VB_Name = "Sequencer"
Public blends As New Collection
Public missingBatches As String
Private scadValues As Variant
Public inProgress As Integer
Public ordersFromBeyond As String
Public gSource As String
Public pSource As String
Public period As String
Public blendKeepers As New Collection
Public SessionId As Integer

Public Sub updateGreatSummary(control As IRibbonControl)
getDates.Show
'ScadaSequancer
End Sub

Public Sub formatMe(Optional nSht As Variant)
Dim sht As Worksheet

If IsMissing(nSht) Then
    Set sht = ThisWorkbook.Sheets("Operations sequence")
Else
    Set sht = nSht
End If

With sht
    '.Cells.clear
    .Range("A1:A2").Merge
    .Range("B1:B2").Merge
    .Range("C1:D1").Merge
    .Range("E1:F1").Merge
    .Range("G1:H1").Merge
    .Range("I1:J1").Merge
    .Range("K1:L1").Merge
    .Range("M1:M2").Merge
    .Range("N1:N2").Merge
    .Range("O1:P1").Merge
    .Range("Q1:R1").Merge
    .Range("S1:T1").Merge
    .Range("U1:V1").Merge
    .Range("W1:X1").Merge
    .Range("Y1:Z1").Merge
    .Range("AA1:AC1").Merge
    .Range("AD1:AF1").Merge
    .Range("AG1:AH1").Merge
    .Range("AI1:AJ1").Merge
    .Range("AK1:AM1").Merge
    .Range("AN1:AN2").Merge
    .Range("AO1:AO2").Merge
    .Range("AP1:AP2").Merge
    .Range("AQ1:AQ2").Merge
    .Range("AR1:AR2").Merge
    .Range("AS1:AS2").Merge
    .Range("A1").Value = "ZFOR"
    .Range("B1").Value = "Description"
    .Range("C1").Value = "Green coffee"
    .Range("C2").Value = "ID"
    .Range("D2").Value = "Amount [kg]"
    .Range("E1").Value = "Roasting"
    .Range("E2").Value = "ID"
    .Range("F2").Value = "Amount [kg]"
    .Range("G1").Value = "Loss"
    .Range("G2").Value = "kg"
    .Range("H2").Value = "%"
    .Range("I1").Value = "Grinding"
    .Range("I2").Value = "ID"
    .Range("J2").Value = "Amount [kg]"
    .Range("K1").Value = "Loss"
    .Range("K2").Value = "kg"
    .Range("L2").Value = "%"
    .Range("M1").Value = "ZFIN"
    .Range("N1").Value = "Description"
    .Range("O1").Value = "Packing"
    .Range("O2").Value = "ID"
    .Range("P2").Value = "Amount [kg]"
    .Range("Q1").Value = "Loss"
    .Range("Q2").Value = "kg"
    .Range("R2").Value = "%"
    .Range("S1").Value = "Warehouse"
    .Range("S2").Value = "ID"
    .Range("T2").Value = "Amount [kg]"
    .Range("U1").Value = "Loss"
    .Range("U2").Value = "kg"
    .Range("V2").Value = "%"
    .Range("W1").Value = "TOTAL LOSS"
    .Range("W2").Value = "kg"
    .Range("X2").Value = "%"
    .Range("y1").Value = "G + P LOSS"
    .Range("Y2").Value = "kg"
    .Range("Z2").Value = "%"
    .Range("AA1").Value = "BOM's scrap [%]"
    .Range("AA2").Value = "R+G"
    .Range("AB2").Value = "P"
    .Range("AC2").Value = "Total"
    .Range("AD1").Value = "BOM vs real scrap [%]"
    .Range("AD2").Value = "R+G"
    .Range("AE2").Value = "P"
    .Range("AF2").Value = "Total"
    .Range("AG1").Value = "RN3000's loss"
    .Range("AG2").Value = "kg"
    .Range("AH2").Value = "%"
    .Range("AI1").Value = "RN4000's loss"
    .Range("AI2").Value = "kg"
    .Range("AJ2").Value = "%"
    .Range("AK1").Value = "Coffee value [k€]"
    .Range("AK2").Value = "Roasted"
    .Range("AL2").Value = "Packed"
    .Range("AM2").Value = "Loss"
    .Range("AN1").Value = "Grinding rework [kg]"
    .Range("AN1").WrapText = True
    .Range("AO1").Value = "Packing rework [kg]"
    .Range("AO1").WrapText = True
    .Range("AP1").Value = "Roasting vs average"
    .Range("AP1").WrapText = True
    .Range("AQ1").Value = "Grinding vs average"
    .Range("AQ1").WrapText = True
    .Range("AR1").Value = "Packing vs average"
    .Range("AR1").WrapText = True
    .Range("AS1").Value = "G+P vs average"
    .Range("AS1").WrapText = True
    .Range("AT1").Value = "Total vs average"
    .Range("AT1").WrapText = True
    .Range("A1:AT2").HorizontalAlignment = xlCenter
    .Range("A1:AT2").Font.Bold = True
    .Range("A1:AT2").Interior.ColorIndex = 15
End With
End Sub

Public Sub ScadaSequancer()
Dim i As Integer
Dim lastRow As Long
Dim sht As Worksheet
Dim cSheet As Worksheet
Dim rng As Range

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets("SCADA")
Set cSheet = ThisWorkbook.Sheets("Operations sequence")

lastRow = sht.Range("J:J").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
Set rng = sht.Range("J2:J" & lastRow)
rng.Copy cSheet.Range("A3:A" & lastRow + 1)
rng.Copy cSheet.Range("C3:C" & lastRow + 1)
Set rng = sht.Range("O2:O" & lastRow)
rng.Copy cSheet.Range("B3:B" & lastRow + 1)
Set rng = sht.Range("P2:P" & lastRow)
rng.Copy cSheet.Range("D3:D" & lastRow + 1)

exit_here:
Set sht = Nothing
Set cSheet = Nothing
Set rng = Nothing
Exit Sub

err_trap:
MsgBox "Error in ScadaSequancer. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Sub scadaSummary(Optional fromD As Variant, Optional toD As Variant, Optional bOption As Integer, Optional ex As Variant, Optional lim As Variant, Optional formatArray As Variant, Optional progress As Variant, Optional grindingSource As Variant, Optional packingSource As Variant, Optional toExpand As Variant, Optional notToExpand)
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim i As Integer 'zfor operation counter
Dim i1 As Integer
Dim ii As Integer
Dim i2 As Integer
Dim step As Integer
Dim step1 As Integer
Dim s As Integer 'column "A","B","G","H" counter
Dim SQL As String
Dim sTime As Date
Dim eTime As Date
Dim zforInd As Long
Dim zforName As String
Dim roast As Double
Dim green As Double
Dim ords() As Double
Dim nBlend As clsBlend
Dim nzfin As clsZfin
Dim totalZfins As Integer
Dim sht As Worksheet
Dim rng As Range
Dim c As Range
Dim n As Integer
Dim g As Integer
Dim totalZforOrders As Integer
Dim totalZfinOrders As Integer
Dim totalGreen As Double
Dim TotalRoast As Double
Dim totalPacked As Double
Dim totalStocked As Double
Dim totalGround As Double
Dim groundCounted As Double
Dim beansCounted As Double
Dim packedCounted As Double
Dim rgCounted As Double
Dim totalRg As Double
Dim stockCounted As Double
Dim initialGreen As Double
Dim greenCounted As Double
Dim greenUncounted As Double
'Dim fldNames(1) As String
Dim roastCounted As Double
Dim zforStr As String
Dim ind As Integer
Dim showOnlyCountedSums As Boolean
Dim repStr As String
Dim bool As Boolean
Dim inconsistentOrder As String
Dim invalidZfor As String
Dim m As Integer
Dim countMe As Boolean
Dim greenReceipt As Variant
Dim msgStr As String
Dim stroke As Boolean
Dim cutOffStr As String
Dim bm As clsBM
Dim greenDetails As Variant
Dim ranger As Double
Dim v() As String
Dim blendKeeper As New clsBlendKeeper
Dim processType As String
Dim valueType As String
Dim orderType As String
Dim vsBom As Boolean
Dim vsAvg As Boolean
Dim roaster As Variant
Dim ggreen As Double
Dim rroasted As Double
Dim countRework As Boolean
Dim countReworkWarehouse As Boolean
Dim expansionList As String
Dim orderList As String
Dim z As Integer

countRework = getDates.cboxRework
countReworkWarehouse = getDates.cboxReworkWarehouse

ordersFromBeyond = ""

If IsMissing(grindingSource) Then
    gSource = "MES"
Else
    gSource = grindingSource
End If

If IsMissing(packingSource) Then
    pSource = "SAP"
Else
    pSource = packingSource
End If

If IsMissing(progress) Then
    inProgress = 10
Else
    inProgress = progress
End If

Set bm = New clsBM

ThisWorkbook.Sheets("Operations sequence").Cells.clear
formatMe

showOnlyCountedSums = True

n = blends.Count
Do While blends.Count > 0
    blends.Remove n
    n = n - 1
Loop

missingBatches = ""

'fldNames(0) = "MaterialNumber"
'fldNames(1) = "theOrder"

connectScada
Set sht = ThisWorkbook.Sheets("Operations sequence")

If Not IsMissing(fromD) Then
    If IsDate(fromD) Then
        sTime = fromD
    Else
        sTime = ThisWorkbook.CustomDocumentProperties("roastingFrom")
    End If
Else
    sTime = ThisWorkbook.CustomDocumentProperties("roastingFrom")
End If
If Not IsMissing(toD) Then
    If IsDate(toD) Then
        eTime = toD
    Else
        eTime = ThisWorkbook.CustomDocumentProperties("roastingTo")
    End If
Else
    eTime = ThisWorkbook.CustomDocumentProperties("roastingTo")
End If

'sql = "SELECT DISTINCT sum(z.SUMA_ZIELONEJ) as sumaZielonej, sum(z.ILOSC_PALONA) as sumaPalonej, min(z.DTZAPIS) as minData, max(z.DTZAPIS) as maxData,zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT " _
'    & "FROM ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE) " _
'    & "WHERE (z.DTZAPIS Between ('" & ThisWorkbook.CustomDocumentProperties("roastingFrom") & "') AND ('" & ThisWorkbook.CustomDocumentProperties("roastingTo") & "')) " _
'    & "GROUP BY zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT " _
'    & "ORDER BY min(z.DTZAPIS);"

SQL = "SELECT DISTINCT rD.OrderNumber as theOrder FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE) " _
    & "Where (z.DTZAPIS Between ('" & sTime & "') AND ('" & eTime & "'))) as rD"

Set rs1 = CreateObject("adodb.recordset")
rs1.Open SQL, conn
If Not rs1.EOF Then
    rs1.MoveFirst
    Do Until rs1.EOF
        zforStr = zforStr & "'" & rs1.Fields("theOrder") & "',"
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing

If Len(zforStr) = 0 Then
    MsgBox "No data has been found for chosen period.", vbOKOnly + vbInformation, "No data"
Else
    For Each blendKeeper In blendKeepers
        If blendKeeper.Id > highestId Then highestId = blendKeeper.Id
    Next blendKeeper
    If Len(period) > 0 Then
        v = Split(period, "|", , vbTextCompare)
        blendKeeper.week = CInt(v(0))
        blendKeeper.year = CInt(v(1))
    End If
    blendKeeper.Id = highestId + 1
    zforStr = Left(zforStr, Len(zforStr) - 1)
    
    If getDates.cBoxBeyond.Value = True Then
        SQL = "select DISTINCT zl.OrderNumber as theOrder from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON z.IDZLECENIE = w.IDZLECENIE JOIN ZLECENIA zl on w.IDZLECENIE = zl.IDZLECENIE " _
            & "WHERE zl.OrderNumber IN (" & zforStr & ") AND (z.DTZAPIS < '" & sTime & "' OR z.DTZAPIS > '" & eTime & "')"
    
        Set rs = CreateObject("adodb.recordset")
        rs.Open SQL, conn
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                cutOffStr = cutOffStr & rs.Fields("theOrder") & ","
                rs.MoveNext
            Loop
            rs.Close
            If Len(cutOffStr) > 0 Then cutOffStr = Left(cutOffStr, Len(cutOffStr) - 1)
        End If
    End If

'    sql = "SELECT rD.OrderNumber as theOrder, rd.MaterialNumber as zfor, rd.NAZWARECEPT as name, sum(rd.SUMA_ZIELONEJ) as sumaZielonej,sum(rd.ILOSC_PALONA) As sumaPalonej " _
'        & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE) Where (z.DTZAPIS Between ('" & sTime & "') AND ('" & eTime & "'))) as rD " _
'        & "GROUP BY rd.OrderNumber,rd.MaterialNumber,rd.NAZWARECEPT ORDER BY zfor;"
    
    SQL = "SELECT rD.OrderNumber as theOrder, rd.MaterialNumber as zfor, rd.NAZWARECEPT as name, sum(rd.SUMA_ZIELONEJ) as sumaZielonej,sum(rd.ILOSC_PALONA) As sumaPalonej, rD.NUMERPIECA as roaster " _
    & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)) as rD " _
    & "WHERE rd.OrderNumber IN (" & zforStr & ") " _
    & "GROUP BY rd.OrderNumber,rd.MaterialNumber,rd.NAZWARECEPT, rD.NUMERPIECA " _
    & "ORDER BY zfor"

    
    'Set rs = conn.Execute(sql)
    Set rs = CreateObject("adodb.recordset")
    rs.Open SQL, conn
    If Not rs.EOF Then
        rs.MoveFirst
        scadValues = rs.GetRows(Fields:="theOrder")
        If getDates.cboxAllowExpansion.Value = True Then
            expandScope 'check's if  we need to expand scope
            If Len(ordersFromBeyond) > 0 Then
                Erase scadValues
                rs.Close
                Set rs = Nothing
                SQL = "SELECT rD.OrderNumber as theOrder, rd.MaterialNumber as zfor, rd.NAZWARECEPT as name, sum(rd.SUMA_ZIELONEJ) as sumaZielonej,sum(rd.ILOSC_PALONA) As sumaPalonej,rD.NUMERPIECA as roaster " _
                    & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)) as rD " _
                    & "WHERE rd.OrderNumber IN (" & zforStr & "," & ordersFromBeyond & ") " _
                    & "GROUP BY rd.OrderNumber,rd.MaterialNumber,rd.NAZWARECEPT,rD.NUMERPIECA " _
                    & "ORDER BY zfor"
                Set rs = CreateObject("adodb.recordset")
                rs.Open SQL, conn
                If Not rs.EOF Then
                    rs.MoveFirst
                    scadValues = rs.GetRows(Fields:="theOrder")
                End If
            End If
        ElseIf Not IsMissing(toExpand) Or Not IsMissing(notToExpand) Then
            For z = LBound(scadValues, 2) To UBound(scadValues, 2)
                orderList = orderList & scadValues(0, z) & ","
            Next z
            If Len(orderList) > 0 Then
                orderList = Left(orderList, Len(orderList) - 1)
                If Not IsMissing(toExpand) Then
                    'expand roasting range for given blends
                    expansionList = blendKeeper.trimOrdersToBlends(orderList, toExpand)
                    expandScope expansionList, True 'check's if  we need to expand scope
                Else
                    'expand roasting range for all blends except given blends
                    expansionList = blendKeeper.trimOrdersToBlends(orderList, notToExpand)
                    expandScope expansionList, False 'check's if  we need to expand scope
                End If
                If Len(ordersFromBeyond) > 0 Then
                    Erase scadValues
                    rs.Close
                    Set rs = Nothing
                    SQL = "SELECT rD.OrderNumber as theOrder, rd.MaterialNumber as zfor, rd.NAZWARECEPT as name, sum(rd.SUMA_ZIELONEJ) as sumaZielonej,sum(rd.ILOSC_PALONA) As sumaPalonej,rD.NUMERPIECA as roaster " _
                        & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)) as rD " _
                        & "WHERE rd.OrderNumber IN (" & zforStr & "," & ordersFromBeyond & ") " _
                        & "GROUP BY rd.OrderNumber,rd.MaterialNumber,rd.NAZWARECEPT,rD.NUMERPIECA " _
                        & "ORDER BY zfor"
                    Set rs = CreateObject("adodb.recordset")
                    rs.Open SQL, conn
                    If Not rs.EOF Then
                        rs.MoveFirst
                        scadValues = rs.GetRows(Fields:="theOrder")
                    End If
                End If
            End If
        End If
        If getDates.cboxSession.Value = True Then
            Erase scadValues
                rs.Close
                Set rs = Nothing
                SQL = "SELECT rD.OrderNumber as theOrder, rd.MaterialNumber as zfor, rd.NAZWARECEPT as name, sum(rd.SUMA_ZIELONEJ) as sumaZielonej,sum(rd.ILOSC_PALONA) As sumaPalonej,rD.NUMERPIECA as roaster " _
                    & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)) as rD " _
                    & "WHERE rd.OrderNumber IN (" & Trim2Session(zforStr) & ") " _
                    & "GROUP BY rd.OrderNumber,rd.MaterialNumber,rd.NAZWARECEPT,rD.NUMERPIECA " _
                    & "ORDER BY zfor"
                Set rs = CreateObject("adodb.recordset")
                rs.Open SQL, conn
                If Not rs.EOF Then
                    rs.MoveFirst
                    scadValues = rs.GetRows(Fields:="theOrder")
                End If
        End If
        rs.MoveFirst
        Do Until rs.EOF
            Set nBlend = newBlend(rs.Fields("zfor"), rs.Fields("name"), blendKeeper)
            With nBlend
                'initialGreen = initialGreen + CDbl(rs.Fields("sumaZielonej").value)
                If InStr(1, cutOffStr, rs.Fields("theOrder").Value, vbTextCompare) = 0 And IsNull(rs.Fields("sumaZielonej").Value) = False And IsNull(rs.Fields("sumaPalonej").Value) = False Then
                    If isOrderConsistent(rs.Fields("theOrder").Value) Then
                        .addGreen CDbl(rs.Fields("theOrder").Value), CDbl(rs.Fields("sumaZielonej").Value), rs.Fields("roaster").Value
                        .addRoasted CDbl(rs.Fields("theOrder").Value), CDbl(rs.Fields("sumaPalonej").Value), rs.Fields("roaster").Value
                        initialGreen = initialGreen + CDbl(rs.Fields("sumaZielonej").Value)
                    Else
                        greenUncounted = greenUncounted + CDbl(rs.Fields("sumaZielonej").Value)
                        If Len(inconsistentOrder) = 0 Then
                            inconsistentOrder = "No packing order(s) has been found for roasting order(s) " & rs.Fields("theOrder").Value & " (zfor " & rs.Fields("zfor") & ")"
                        Else
                            inconsistentOrder = inconsistentOrder & "," & rs.Fields("theOrder").Value & " (zfor " & rs.Fields("zfor") & ")"
                        End If
                    End If
                End If
            End With
            rs.MoveNext
        Loop
        If Len(inconsistentOrder) > 0 Then
            inconsistentOrder = inconsistentOrder & ". Hence they are not taken into account in Mass balance result" & vbNewLine & vbNewLine
        End If
        i = 4
        i2 = 4
        For Each nBlend In blendKeeper.blends
            If nBlend.numberOfOrders > 0 Then
                bool = True
                nBlend.setPacked
                If Not IsMissing(ex) Then
                    If Not isArrayEmpty(ex) Then
                        For ind = LBound(ex) To UBound(ex)
                            If CLng(ex(ind)) = nBlend.index Then
                                bool = False
                                Exit For
                            End If
                        Next ind
                    End If
                ElseIf Not IsMissing(lim) Then
                    bool = False
                    If Not isArrayEmpty(lim) Then
                        For ind = LBound(lim) To UBound(lim)
                            If CLng(lim(ind)) = nBlend.index Then
                                bool = True
                                Exit For
                            End If
                        Next ind
                    End If
                End If
                If bOption <> 0 Then
                    If bOption = 1 And nBlend.IsBeans Then
                        bool = False
                    ElseIf bOption = 2 And nBlend.IsBeans = False Then
                        bool = False
                    End If
                End If
                If bool Then
                    nBlend.inScope = True
                    If i > 4 Then
                        i = ii + n
                    End If
                    ii = i
                    i2 = i
                    totalGreen = totalGreen + nBlend.getGreen
                    TotalRoast = TotalRoast + nBlend.getRoasted
                    If nBlend.IsBeans = False Then totalGround = totalGround + nBlend.getGround
                    m = nBlend.numberOfOrders
                    n = nBlend.numberOfZfinOrders
                    totalZfinOrders = totalZfinOrders + n
                    totalZforOrders = totalZforOrders + m
                    bool = nBlend.isConsistent
                    If bool Then
                        bm.groundStr = bm.groundStr & nBlend.index & vbNewLine
                        greenCounted = greenCounted + nBlend.getGreen
                        roastCounted = roastCounted + nBlend.getRoasted
                        stockCounted = stockCounted + nBlend.getStocked
                        packedCounted = packedCounted + nBlend.getPacked
                        bm.addGreen nBlend.getGreen, nBlend.IsBeans
                        bm.addRoast nBlend.getRoasted, nBlend.IsBeans
                        bm.addWarehoused nBlend.getStocked, nBlend.IsBeans
                        bm.addPacked nBlend.getPacked, nBlend.IsBeans
                        If nBlend.IsBeans = False Then
                            If nBlend.isGroundConsistent Then
                                groundCounted = groundCounted + nBlend.getGround
                                rgCounted = rgCounted + nBlend.getRoasted
                            End If
                            bm.addGround nBlend.getGround, False
                        Else
                            beansCounted = beansCounted + nBlend.getRoasted
                            bm.addGround nBlend.getRoasted, True
                        End If

                    End If
                    totalZfins = totalZfins + nBlend.numberOfZfins

                    For Each nzfin In nBlend.getZfins
                        If IsArray(nzfin.getOrders) Then
                            If nzfin.getPacked > 0 Then
                                sht.Range("V" & i) = -1 * Round(Round(nzfin.getPacked - nzfin.getStocked, 1) / nzfin.getPacked, 4)
                            End If
                            totalPacked = totalPacked + nzfin.getPacked
                            totalStocked = totalStocked + nzfin.getStocked

                        End If
                    Next nzfin
                End If
            End If
        Next nBlend
        If Len(getDates.cmbSortType) > 0 Then
            If Len(getDates.cmbSortOrder) = 0 Then
                orderType = "ASC"
            Else
                orderType = getDates.cmbSortOrder
            End If
            vsBom = False
            roaster = Null
            vsAvg = False
            Select Case getDates.cmbSortType
                Case Is = "Roasting loss in %"
                    processType = "r"
                    valueType = "%"
                Case Is = "Grinding loss in %"
                    processType = "g"
                    valueType = "%"
                Case Is = "Packing loss in %"
                    processType = "p"
                    valueType = "%"
                Case Is = "Total loss in %"
                    processType = "t"
                    valueType = "%"
                Case Is = "Roasting loss in kg"
                    processType = "r"
                    valueType = "kg"
                Case Is = "Grinding loss in kg"
                    processType = "g"
                    valueType = "kg"
                Case Is = "Packing loss in kg"
                    processType = "p"
                    valueType = "kg"
                Case Is = "Total loss in kg"
                    processType = "t"
                    valueType = "kg"
                Case Is = "Grinding + packing loss in kg"
                    processType = "g+p"
                    valueType = "kg"
                Case Is = "Grinding + packing loss in %"
                    processType = "g+p"
                    valueType = "%"
                Case Is = "Real vs BOM for roasting + grinding in %"
                    processType = "r+g"
                    valueType = "%"
                    vsBom = True
                Case Is = "Real vs BOM for packing in %"
                    processType = "p"
                    valueType = "%"
                    vsBom = True
                Case Is = "Real vs BOM for total in %"
                    processType = "t"
                    valueType = "%"
                    vsBom = True
                Case Is = "Roasting loss on RN3000 in %"
                    valueType = "%"
                    processType = "r"
                    roaster = 3000
                Case Is = "Roasting loss on RN3000 in kg"
                    valueType = "kg"
                    processType = "r"
                    roaster = 3000
                Case Is = "Roasting loss on RN4000 in %"
                    valueType = "%"
                    processType = "r"
                    roaster = 4000
                Case Is = "Roasting loss on RN4000 in kg"
                    valueType = "kg"
                    processType = "r"
                    roaster = 4000
                Case Is = "Roasted coffee value"
                    valueType = "$"
                    processType = "r"
                Case Is = "Packed coffee value"
                    valueType = "$"
                    processType = "p"
                Case Is = "Lost value on grinding + packing"
                    valueType = "$"
                    processType = "g+p"
                Case Is = "Roasting loss vs average"
                    valueType = "%"
                    processType = "r"
                    vsAvg = True
                Case Is = "Grinding loss vs average"
                    valueType = "%"
                    processType = "g"
                    vsAvg = True
                Case Is = "Packing loss vs average"
                    valueType = "%"
                    processType = "p"
                    vsAvg = True
                Case Is = "Total loss vs average"
                    valueType = "%"
                    processType = "t"
                    vsAvg = True
                Case Is = "Grinding+Packing loss vs average"
                    valueType = "%"
                    processType = "g+p"
                    vsAvg = True
            End Select
        Else
            'default settings
            processType = "t"
            valueType = "%"
            orderType = "ASC"
        End If
        blendKeeper.getScraps
        blendKeeper.downloadCost
        blendKeeper.downloadAvg
        blendKeeper.order processType, valueType, orderType, vsBom, roaster, vsAvg
        If countRework Then
            blendKeeper.calculateRework
            If countReworkWarehouse Then blendKeeper.calculateReworkAtPacking
            bm.addRework blendKeeper.GetRework
            bm.addReworkAtPacking blendKeeper.GetReworkAtPacking
            bm.addReworkAtPacking blendKeeper.GetReworkAtPacking("beans"), "beans"
            bm.addReworkAtPacking blendKeeper.GetReworkAtPacking("ground"), "ground"
        End If
        blendKeeper.display
        
        sht.Range("C3") = totalZforOrders & " ordrers"
        sht.Range("O3") = totalZfinOrders & " orders"
        sht.Range("S3") = totalZfinOrders & " batches"
        sht.Range("A3") = blends.Count & " blends"
        sht.Range("M3") = totalZfins & " products"
        If showOnlyCountedSums Then
            sht.Range("D3") = Round(greenCounted, 1) & " kg"
            sht.Range("F3") = Round(roastCounted, 1) & " kg"
            sht.Range("G3") = -1 * Round(greenCounted - roastCounted, 1) & " kg"
            sht.Range("H3") = -1 * Round(Round(greenCounted - roastCounted, 1) / greenCounted, 4)
            sht.Range("J3") = Round(groundCounted + beansCounted, 1) & " kg"
            sht.Range("K3") = -1 * Round(rgCounted - groundCounted + blendKeeper.GetRework, 1) & " kg"
            If bOption <> 2 And rgCounted > 0 Then
                sht.Range("L3") = -1 * Round(Round(rgCounted - groundCounted + blendKeeper.GetRework, 1) / rgCounted, 4)
                sht.Range("R3") = -1 * Round(Round(groundCounted + beansCounted - stockCounted, 1) / (groundCounted + beansCounted), 4)
            End If
            sht.Range("P3") = Round(packedCounted, 1) & " kg"
            sht.Range("Q3") = -1 * Round(groundCounted + beansCounted - stockCounted, 1) & " kg"
            sht.Range("T3") = Round(stockCounted, 1) & " kg"
            sht.Range("U3") = -1 * Round(packedCounted - stockCounted, 1) & " kg"
            sht.Range("V3") = -1 * Round(Round(packedCounted - stockCounted, 1) / packedCounted, 4)
            greenReceipt = gcwLosses(sTime, eTime)
            If IsNumeric(greenReceipt) Then
                greenReceipt = Abs(greenReceipt)
                bm.greenCoffee = bm.rBeansIn + bm.rGroundIn + greenReceipt
            Else
                greenReceipt = 0
            End If
            greenDetails = gcwDetails(sTime, eTime)
            If IsArray(greenDetails) Then
                bm.mksDiff = greenDetails(0)
                bm.mksPurge = greenDetails(1)
                bm.mksReceipt = greenDetails(2)
            End If
            If initialGreen > totalGreen And greenReceipt <> 0 Then
                ranger = ((100 - Round(100 - ((totalGreen / initialGreen) * 100), 2)) / 100)
                greenReceipt = ranger * greenReceipt
                bm.mksDiff = bm.mksDiff * ranger
                bm.mksPurge = bm.mksPurge * ranger
                bm.mksReceipt = bm.mksReceipt * ranger
                bm.greenCoffee = bm.rBeansIn + bm.rGroundIn + Abs(bm.mksDiff + bm.mksPurge + bm.mksReceipt)
            End If
            sht.Range("W3") = -1 * Round(stockCounted - (greenCounted + greenReceipt + bm.rework + bm.reworkAtPacking), 1) & " kg"
            sht.Range("X3") = -1 * Round(Round(stockCounted - (greenCounted + greenReceipt + bm.rework + bm.reworkAtPacking), 1) / (greenCounted + greenReceipt + bm.rework + bm.reworkAtPacking), 4)
            sht.Range("Y3") = -1 * Round(TotalRoast - stockCounted, 1)
            sht.Range("Z3") = -1 * Round(Round(TotalRoast - stockCounted, 1) / TotalRoast, 4)
            sht.Range("AA3") = -1 * (blendKeeper.roastGroundScrap / 100)
            sht.Range("AB3") = -1 * (blendKeeper.packScrap / 100)
            sht.Range("AC3") = -1 * (blendKeeper.totalScrap / 100)
            sht.Range("AG3") = -1 * (blendKeeper.getRoastingScrap(3000, "kg"))
            sht.Range("AH3") = -1 * (blendKeeper.getRoastingScrap(3000, "%"))
            sht.Range("AI3") = -1 * (blendKeeper.getRoastingScrap(4000, "kg"))
            sht.Range("AJ3") = -1 * (blendKeeper.getRoastingScrap(4000, "%"))
            sht.Range("AK3") = Round((blendKeeper.getRoastedValue + blendKeeper.reworkValue) / 1000, 2)
            sht.Range("AL3") = Round(blendKeeper.getPackedValue / 1000, 2)
            sht.Range("AM3") = Round((blendKeeper.getPackedValue - (blendKeeper.getRoastedValue + blendKeeper.reworkValue)) / 1000, 2)
            sht.Range("AN3") = Round(blendKeeper.GetRework, 2)
            sht.Range("AO3") = Round(blendKeeper.GetReworkAtPacking, 2)
        Else
            sht.Range("D3") = Round(totalGreen, 1) & " kg"
            sht.Range("F3") = Round(TotalRoast, 1) & " kg"
            sht.Range("G3") = -1 * Round(totalGreen - TotalRoast, 1) & " kg"
            sht.Range("H3") = -1 * Round(Round(totalGreen - TotalRoast, 1) / totalGreen, 4)
            sht.Range("J3") = Round(totalGround, 1) & " kg"
            sht.Range("K3") = -1 * Round(TotalRoast - totalGround, 1) & " kg"
            sht.Range("L3") = -1 * Round(Round(TotalRoast - totalGround, 1) / TotalRoast, 4)
            sht.Range("P3") = Round(totalPacked, 1) & " kg"
            sht.Range("Q3") = -1 * Round(totalGround - totalStocked, 1) & " kg"
            sht.Range("R3") = -1 * Round(Round(totalGround - totalStocked, 1) / totalGround, 4)
            sht.Range("T3") = Round(totalStocked, 1) & " kg"
            sht.Range("U3") = -1 * Round(totalPacked - totalStocked, 1) & " kg"
            sht.Range("V3") = -1 * Round(Round(totalPacked - totalStocked, 1) / totalPacked, 4)
            If IsNumeric(greenReceipt) Then
                greenReceipt = Abs(greenReceipt)
            Else
                greenReceipt = 0
            End If
            greenDetails = gcwDetails(sTime, eTime)
            If IsArray(greenDetails) Then
                bm.mksDiff = greenDetails(0)
                bm.mksPurge = greenDetails(1)
                bm.mksLoss = greenDetails(2)
            End If
            If initialGreen > totalGreen And greenReceipt <> 0 Then greenReceipt = ((100 - Round(100 - ((totalGreen / initialGreen) * 100), 2)) / 100) * greenReceipt
            sht.Range("W3") = -1 * Round(stockCounted - (greenCounted + greenReceipt + bm.rework + bm.reworkAtPacking), 1) & " kg"
            sht.Range("X3") = -1 * Round(Round(stockCounted - (greenCounted + greenReceipt + bm.rework + bm.reworkAtPacking), 1) / (greenCounted + greenReceipt + bm.rework + bm.reworkAtPacking), 4)
            sht.Range("AA3") = -1 * (blendKeeper.roastGroundScrap / 100)
            sht.Range("AB3") = -1 * (blendKeeper.packScrap / 100)
            sht.Range("AC3") = -1 * (blendKeeper.totalScrap / 100)
            sht.Range("AG3") = -1 * (blendKeeper.getRoastingScrap(3000, "kg"))
            sht.Range("AH3") = -1 * (blendKeeper.getRoastingScrap(3000, "%"))
            sht.Range("AI3") = -1 * (blendKeeper.getRoastingScrap(4000, "kg"))
            sht.Range("AJ3") = -1 * (blendKeeper.getRoastingScrap(4000, "%"))
            sht.Range("AK3") = Round((blendKeeper.getRoastedValue + blendKeeper.reworkValue) / 1000, 2)
            sht.Range("AL3") = Round(blendKeeper.getPackedValue / 1000, 2)
            sht.Range("AM3") = Round((blendKeeper.getPackedValue - (blendKeeper.getRoastedValue + blendKeeper.reworkValue)) / 1000, 2)
            sht.Range("AN3") = Round(blendKeeper.GetRework, 2)
            sht.Range("AO3") = Round(blendKeeper.GetReworkAtPacking, 2)
        End If
        
        bm.rStart = sTime
        bm.rEnd = eTime
        bm.deployResults
    End If
    rs.Close
    Set rs = Nothing
    
    If Not IsMissing(formatArray) Then putFormating formatArray
    validateBM
    If Len(missingBatches) > 0 Then
        missingBatches = Left(missingBatches, Len(missingBatches) - 1)
        If Len(inconsistentOrder) > 0 Then
            'repStr = repStr & "Please reimport ""Powiązania operacji"" from MES." & vbNewLine & vbNewLine & "Batch numbers for packing orders " & missingBatches & " could not be found. Please reimport COOIS data for these orders, otherwise presented results will not be correct."
            repStr = inconsistentOrder & "Batch numbers for packing orders " & missingBatches & " could not be found. Please reimport COOIS data for these orders, otherwise presented results will not be correct."
        Else
            repStr = "Batch numbers for packing orders " & missingBatches & " could not be found. Please reimport COOIS data for these orders, otherwise presented results will not be correct."
        End If
    Else
        If Len(inconsistentOrder) > 0 Then
            repStr = inconsistentOrder
        End If
    End If
    If Len(repStr) > 0 Then
        MsgBox repStr, vbOKOnly + vbCritical, "Missing data"
    End If
    If Len(ordersFromBeyond) > 0 Then
        MsgBox "Following orders have been included into scope even though they were roasted earlier/later: " & ordersFromBeyond, vbOKOnly + vbExclamation, "Scope expansion"
    End If
    If stockCounted <> totalStock Then
        msgStr = "Mass balance accuracy = " & Round((((stockCounted / totalStocked) + (greenCounted / (totalGreen + greenUncounted))) / 2) * 100, 2) & "% (Operation sequance found for " & Round(stockCounted, 1) & " kg of total " & Round(totalStocked, 1) & " kg."
        If bOption <> 2 And totalGround > 0 Then msgStr = msgStr & vbNewLine & "Mass balance for ground coffee acuracy = " & Round((groundCounted / totalGround) * 100, 1) & " %."
        If initialGreen > totalGreen Then msgStr = msgStr & vbNewLine & Round(100 - ((totalGreen / initialGreen) * 100), 2) & " % of volume has been excluded by user's settings"
        MsgBox msgStr, vbOKOnly + vbInformation, "Inaccuracy disclaimer"
        'MsgBox "Mass balance accuracy = " & Round((greenCounted / initialGreen) * 100, 2) & "% (Operation sequance found for " & Round(greenCounted, 1) & " kg of total " & Round(initialGreen, 1) & " kg.", vbOKOnly + vbInformation, "Inaccuracy disclaimer"
    End If
End If

End Sub

Sub putFormating(fa As Variant)
Dim unit As Integer '0-kg, 1-%
Dim absolut As Boolean
Dim rh As Variant
Dim rl As Variant
Dim gh As Variant
Dim gl As Variant
Dim ph As Variant
Dim pl As Variant
Dim eh As Variant
Dim el As Variant
Dim lastRow As Long
Dim sht As Worksheet
Dim c As Range
Dim rng As Range

Set sht = ThisWorkbook.Sheets("Operations sequence")

If IsArray(fa) Then
    If Not IsNull(fa(0)) Then
        unit = fa(0)
        If Not IsNull(fa(1)) Then absolut = fa(1)
        rh = fa(2)
        rl = fa(3)
        gh = fa(4)
        gl = fa(5)
        ph = fa(6)
        pl = fa(7)
        eh = fa(8)
        el = fa(9)
        lastRow = sht.Range("E:E").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
        If IsNull(rh) = False Or IsNull(rl) = False Then
            If unit = 0 Then
                Set rng = sht.Range("G4:G" & lastRow)
            Else
                Set rng = sht.Range("H4:H" & lastRow)
            End If
            markValues rng, rl, rh, absolut, unit
        End If
        If IsNull(gh) = False Or IsNull(gl) = False Then
            If unit = 0 Then
                Set rng = sht.Range("K4:K" & lastRow)
            Else
                Set rng = sht.Range("L4:L" & lastRow)
            End If
            markValues rng, gl, gh, absolut, unit
        End If
        If IsNull(ph) = False Or IsNull(pl) = False Then
            If unit = 0 Then
                Set rng = sht.Range("Q4:Q" & lastRow)
            Else
                Set rng = sht.Range("R4:R" & lastRow)
            End If
            markValues rng, pl, ph, absolut, unit
        End If
        If IsNull(eh) = False Or IsNull(el) = False Then
           If unit = 0 Then
               Set rng = sht.Range("W4:W" & lastRow)
           Else
               Set rng = sht.Range("X4:X" & lastRow)
           End If
           markValues rng, el, eh, absolut, unit
        End If
    End If
End If
End Sub

Sub markValues(rng As Range, min As Variant, max As Variant, absolut As Boolean, unit As Integer)
Dim c As Range
Dim val As Double
Dim comp As Double

For Each c In rng
    If Not IsNull(max) Then
        If absolut Then val = Abs(c.Value) Else val = c.Value
        If unit = 1 Then comp = max / 100 Else comp = max
        If val > comp Then c.Interior.Color = vbGreen
    End If
    If Not IsNull(min) Then
        If absolut Then val = Abs(c.Value) Else val = c.Value
        If unit = 1 Then comp = min / 100 Else comp = min
        If val < comp Then c.Interior.Color = vbRed
    End If
Next c
End Sub
Private Function newBlend(bNumber As Long, Name As String, Optional bk As clsBlendKeeper) As clsBlend
Dim found As Boolean
Dim b As clsBlend

found = False

If bk.blends.Count > 0 Then
    For Each b In bk.blends
        If b.index = bNumber Then
            found = True
            Set newBlend = b
            Exit For
        End If
    Next b
End If

If found = False Then
    Set newBlend = New clsBlend
    newBlend.index = bNumber
    newBlend.Name = Name
    If Not bk Is Nothing Then
        bk.append newBlend
    Else
        blends.Add newBlend, CStr(bNumber)
    End If
End If

End Function

Public Function isOrderConsistent(ord As Long) As Boolean
Dim bool As Boolean
Dim rs As ADODB.Recordset

On Error GoTo err_trap

updateConnection

bool = True

SQL = "SELECT o.sapId as sapId, COUNT(od.zfinOrder) as Counter " _
                & "FROM tbOrders o LEFT JOIN tbOrderDep od ON od.zforOrder=o.orderId " _
                & "WHERE o.sapId IN (" & ord & ") GROUP BY o.sapId"
            
Set rs = CreateObject("adodb.recordset")
rs.Open SQL, adoConn
If rs.EOF Then
    bool = False
Else
    rs.MoveFirst
    Do Until rs.EOF
        If rs.Fields("counter") = 0 Then
            bool = False
        End If
        rs.MoveNext
    Loop
End If
rs.Close

If getDates.cboxAllowExpansion.Value = False Then
    SQL = "SELECT DISTINCT o2.sapId as zfor " _
        & "FROM tbOrders o LEFT JOIN tbOrderDep od ON od.zforOrder=o.orderId LEFT JOIN tbOrders o1 ON o1.orderId = od.zfinOrder LEFT JOIN tbOrderDep od1 ON od1.zfinOrder = o1.orderId LEFT JOIN tbOrders o2 ON o2.orderId = od1.zforOrder " _
        & "WHERE o.sapId IN (" & ord & ") AND o2.isCancelled <> 1 AND (od.isRemoved IS NULL OR od.isRemoved = 0) AND (od1.isRemoved IS NULL OR od1.isRemoved = 0)"
    rs.Open SQL, adoConn
    If Not rs.EOF Then
        Do Until rs.EOF
            If Not orderInRange(rs.Fields("zfor").Value) Then
                bool = False
                Exit Do
            End If
            rs.MoveNext
        Loop
    End If
End If

isOrderConsistent = bool

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Function

err_trap:
MsgBox "Error in isOrderConsistent . Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Function isBlendValid(rs As ADODB.Recordset, bNumber As Long) As Boolean
Dim os As String
Dim i As Integer
Dim rs1 As ADODB.Recordset

If rs.State = 0 Then rs.Open
If Not rs.EOF Then
    Do Until rs.EOF
        If rs.Fields("zfor").Value = bNumber Then
            os = os & rs.Fields("theOrder").Value & ","
        End If
        rs.MoveNext
    Loop
    If Len(os) > 0 Then os = Left(os, Len(os) - 1)
End If
rs.Close
Set rs = Nothing

End Function

Public Function Trim2Session(currStr As String) As String
'take currStr and check what session number dominates. Then get in all other order numbers from this session and get out all numbers not from the session
Dim rs As ADODB.Recordset
Dim nStr As String
Dim ses As Integer

updateConnection

SQL = "SELECT DISTINCT op.SessionNumber, COUNT(o.sapId) as OrdersPerSession " _
    & "FROM tbOrders o JOIN tbOperations op ON op.orderId=o.orderId " _
    & "WHERE o.sapId IN (" & currStr & ") " _
    & "GROUP BY SessionNumber ORDER BY OrdersPerSession DESC"
            
Set rs = CreateObject("adodb.recordset")
rs.Open SQL, adoConn
If rs.EOF Then
    nStr = currStr
Else
    rs.MoveFirst
    ses = rs.Fields("SessionNumber").Value
End If
rs.Close
Set rs = Nothing

If ses > 0 Then
    SQL = "SELECT DISTINCT o.sapId " _
        & "FROM tbOrders o JOIN tbOperations op ON op.orderId=o.orderId " _
        & "WHERE op.SessionNumber=" & ses & " AND o.type = 'r'"
    
    Set rs = CreateObject("adodb.recordset")
    rs.Open SQL, adoConn
    If rs.EOF Then
        nStr = currStr
    Else
        rs.MoveFirst
        Do Until rs.EOF
            If Len(nStr) = 0 Then
                nStr = "'" & rs.Fields("sapId").Value & "'"
            Else
                nStr = nStr & ",'" & rs.Fields("sapId").Value & "'"
            End If
            rs.MoveNext
        Loop
    End If
End If

Trim2Session = nStr

End Function

Public Sub expandScope(Optional oList As Variant, Optional toBeLimited As Variant)
'If any of roasting orders in current scope points to ZFIN's order that points back to roasting order beyond current scope
'then current scope gets expanded

Dim i As Integer
Dim ord As Long
Dim rs As ADODB.Recordset
Dim os() As Long
Dim added As Variant

updateConnection

If IsMissing(oList) Then
    ReDim os(UBound(scadValues, 2)) As Long
    
    For i = LBound(scadValues, 2) To UBound(scadValues, 2)
        os(i) = scadValues(0, i)
    Next i
Else
    
    For i = LBound(scadValues, 2) To UBound(scadValues, 2)
        If toBeLimited Then
            If InStr(1, oList, scadValues(0, i), vbTextCompare) > 0 Then
                'order is in given order list
                If isArrayEmpty(os) Then
                    ReDim os(0) As Long
                Else
                    ReDim Preserve os(UBound(os) + 1) As Long
                End If
                os(UBound(os)) = scadValues(0, i)
            End If
        Else
            If InStr(1, oList, scadValues(0, i), vbTextCompare) = 0 Then
                'order is NOT in given order list
                If isArrayEmpty(os) Then
                    ReDim os(0) As Long
                Else
                    ReDim Preserve os(UBound(os) + 1) As Long
                End If
                os(UBound(os)) = scadValues(0, i)
            End If
        End If
    Next i
End If
    
added = expand(os)
If Not IsNull(added) Then
    'we haven't added all needed blends to base (scadValues) yet - loop once more
    Do Until IsNull(added)
        added = expand(added)
    Loop
End If
If Len(ordersFromBeyond) > 0 Then ordersFromBeyond = Left(ordersFromBeyond, Len(ordersFromBeyond) - 1)
End Sub

Public Function expand(ord As Variant) As Variant
'checks if any of roasting orders in ord (array of longs) points to ZFIN's order that points back to roasting order beyond current scope (scadValues)
'then current scope (scadValues) gets expanded
Dim i As Integer
Dim rs As ADODB.Recordset
Dim added() As Long

updateConnection

'ord = scadValues(0, i)

Set rs = CreateObject("adodb.recordset")

For i = LBound(ord) To UBound(ord)
    SQL = "SELECT DISTINCT o2.sapId as zfor " _
        & "FROM tbOrders o LEFT JOIN tbOrderDep od ON od.zforOrder=o.orderId LEFT JOIN tbOrders o1 ON o1.orderId = od.zfinOrder LEFT JOIN tbOrderDep od1 ON od1.zfinOrder = o1.orderId LEFT JOIN tbOrders o2 ON o2.orderId = od1.zforOrder " _
        & "WHERE o.sapId IN (" & ord(i) & ") AND o2.isCancelled <> 1 AND od.isRemoved IS NULL"
    rs.Open SQL, adoConn
    If Not rs.EOF Then
        Do Until rs.EOF
            If Not orderInRange(rs.Fields("zfor").Value) Then
                ReDim Preserve scadValues(0, UBound(scadValues, 2) + 1) As Variant
                scadValues(0, UBound(scadValues, 2)) = rs.Fields("zfor").Value
                If isArrayEmpty(added) Then
                    ReDim added(0) As Long
                    added(0) = rs.Fields("zfor").Value
                Else
                    ReDim Preserve added(UBound(added) + 1) As Long
                    added(UBound(added)) = rs.Fields("zfor").Value
                End If
                If InStr(1, ordersFromBeyond, CStr(rs.Fields("zfor").Value), vbTextCompare) = 0 Then
                    ordersFromBeyond = ordersFromBeyond & "'" & rs.Fields("zfor").Value & "',"
                End If
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    
Next i

If isArrayEmpty(added) Then
    expand = Null
Else
    expand = added
End If

End Function

Public Function gcwLosses(sDate As Date, eDate As Date) As Variant
'gcw = green coffee warehouse
Dim sWeek As Integer
Dim eWeek As Integer
Dim sYear As Integer
Dim eYear As Integer
Dim totalLoss As Double
Dim SQL As String
Dim rs As ADODB.Recordset
Dim dif As Double
Dim subDif As Double
Dim tot As Double
Dim param As Double
Dim bmId As Integer
Dim bmstr As String

On Error GoTo err_trap

updateConnection

'sWeek = CInt(IsoWeekNumber(sDate))
'eWeek = CInt(IsoWeekNumber(eDate))
'sYear = year(sDate)
'eYear = year(eDate)
'
'If sWeek = eWeek And sYear = eYear Then
'    'we've got 1 week only
'    sql = "SELECT * FROM tbBM WHERE bmWeek = " & sWeek & " AND bmYear = " & sYear
'End If

SQL = "SELECT bmId, bmMonth FROM tbBM WHERE roastingFrom >= '" & sDate & "' AND roastingTo <= '" & eDate & "'"

Set rs = CreateObject("adodb.recordset")
rs.Open SQL, adoConn
If Not rs.EOF Then
    Do Until rs.EOF
        If ThisWorkbook.CustomDocumentProperties("PeriodType") = "custom" Then
            'ignore monthlys
            If IsNull(rs.Fields("bmMonth")) Then
                bmstr = bmstr & rs.Fields("bmId") & ","
            End If
        ElseIf ThisWorkbook.CustomDocumentProperties("PeriodType") = "monthly" Then
            'get monthly if available and finish, or sum up appropriate weeklys
            If IsNull(rs.Fields("bmMonth")) Then
                bmstr = bmstr & rs.Fields("bmId") & ","
            Else
                bmstr = rs.Fields("bmId") & ","
                Exit Do
            End If
        Else
            'assume weekly
            bmstr = bmstr & rs.Fields("bmId") & ","
        End If
        rs.MoveNext
    Loop
    rs.Close
    If Len(bmstr) > 0 Then
        bmstr = Left(bmstr, Len(bmstr) - 1)
        SQL = "SELECT SUM(zLoss+cleaningLoss+receiptLoss) as losses " _
            & "FROM tbBMDetails WHERE bmId IN (" & bmstr & ")"
        rs.Open SQL, adoConn
        If Not rs.EOF Then
            tot = rs.Fields("losses")
        End If
    End If
End If
rs.Close

'sql = "SELECT TOP 1 bmId, roastingFrom, roastingTo FROM tbBM WHERE roastingFrom < '" & sDate & "' AND roastingTo > '" & sDate & "' ORDER BY roastingFrom DESC"
'
''Set rs = CreateObject("adodb.recordset")
'rs.Open sql, adoConn
'If Not rs.EOF Then
'    dif = Abs(DateDiff("h", rs.Fields("roastingTo"), rs.Fields("roastingFrom")))
'    subDif = Abs(DateDiff("h", rs.Fields("roastingFrom"), sDate))
'    If dif = 0 Then
'        param = 0
'    Else
'        param = subDif / dif
'        bmId = rs.Fields("bmId")
'    End If
'End If
'rs.Close
'If param > 0 Then
'    sql = "SELECT SUM(zLoss+cleaningLoss+receiptLoss) as losses " _
'        & "FROM tbBMDetails WHERE bmId = " & bmId
'    rs.Open sql, adoConn
'    If Not rs.EOF Then
'        rs.MoveFirst
'        tot = tot + (rs.Fields("losses") * param)
'    End If
'    rs.Close
'End If
'
'param = 0
'
'sql = "SELECT TOP 1 bmId, roastingFrom, roastingTo FROM tbBM WHERE roastingTo > '" & eDate & "' AND roastingFrom < '" & eDate & "' ORDER BY roastingTo DESC"
'
''Set rs = CreateObject("adodb.recordset")
'rs.Open sql, adoConn
'If Not rs.EOF Then
'    dif = Abs(DateDiff("h", rs.Fields("roastingTo"), rs.Fields("roastingFrom")))
'    subDif = Abs(DateDiff("h", eDate, rs.Fields("roastingFrom")))
'    If dif = 0 Then
'        param = 0
'    Else
'        param = subDif / dif
'        bmId = rs.Fields("bmId")
'    End If
'End If
'rs.Close
'If param > 0 Then
'    sql = "SELECT SUM(zLoss+cleaningLoss+receiptLoss) as losses " _
'        & "FROM tbBMDetails WHERE bmId = " & bmId
'    rs.Open sql, adoConn
'    If Not rs.EOF Then
'        rs.MoveFirst
'        tot = tot + (rs.Fields("losses") * param)
'    End If
'    rs.Close
'End If

gcwLosses = tot

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Function

err_trap:
MsgBox "Error in gcwLosses . Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Function gcwDetails(sDate As Date, eDate As Date) As Variant
'gcw = green coffee warehouse
Dim sWeek As Integer
Dim eWeek As Integer
Dim sYear As Integer
Dim eYear As Integer
Dim totalLoss As Double
Dim SQL As String
Dim rs As ADODB.Recordset
Dim arr(2) As Double
Dim dif As Double
Dim subDif As Double
Dim tot As Double
Dim param As Double
Dim bmId As Integer
Dim bmstr As String

On Error GoTo err_trap

updateConnection

arr(0) = 0
arr(1) = 0
arr(2) = 0

SQL = "SELECT bmId, bmMonth FROM tbBM WHERE roastingFrom >= '" & sDate & "' AND roastingTo <= '" & eDate & "'"

Set rs = CreateObject("adodb.recordset")
rs.Open SQL, adoConn
If Not rs.EOF Then
    Do Until rs.EOF
        If ThisWorkbook.CustomDocumentProperties("PeriodType") = "custom" Then
            'ignore monthlys
            If IsNull(rs.Fields("bmMonth")) Then
                bmstr = bmstr & rs.Fields("bmId") & ","
            End If
        ElseIf ThisWorkbook.CustomDocumentProperties("PeriodType") = "monthly" Then
            'get monthly if available and finish, or sum up appropriate weeklys
            If IsNull(rs.Fields("bmMonth")) Then
                bmstr = bmstr & rs.Fields("bmId") & ","
            Else
                bmstr = rs.Fields("bmId") & ","
                Exit Do
            End If
        Else
            'assume weekly
            bmstr = bmstr & rs.Fields("bmId") & ","
        End If
        rs.MoveNext
    Loop
    rs.Close
    If Len(bmstr) > 0 Then
        bmstr = Left(bmstr, Len(bmstr) - 1)
        SQL = "SELECT SUM(zLoss) as zetki, SUM(cleaningLoss) as czyszczenie, SUM(receiptLoss) as odbior " _
            & "FROM tbBMDetails WHERE bmId IN (" & bmstr & ")"
        rs.Open SQL, adoConn
        If Not rs.EOF Then
            arr(0) = rs.Fields("zetki")
            arr(1) = rs.Fields("czyszczenie")
            arr(2) = rs.Fields("odbior")
        End If
    End If
End If
rs.Close

If arr(0) = 0 And arr(1) = 0 And arr(2) = 0 Then
    gcwDetails = Null
Else
    gcwDetails = arr
End If

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Function

err_trap:
MsgBox "Error in gcwDetails . Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Private Function orderInRange(ord As Long) As Boolean
'checks if order "ord" is in currently chosen scada range.
Dim i As Long
Dim bool As Boolean

bool = False

If IsArray(scadValues) Then
    For i = LBound(scadValues, 2) To UBound(scadValues, 2)
        If ord = scadValues(0, i) Then
            bool = True
            Exit For
        End If
    Next i
End If

orderInRange = bool

End Function

Public Sub hideSheet(val As Boolean, Optional shtName As Variant)

Dim sht As Worksheet

For Each sht In ThisWorkbook.Sheets
    If Not IsMissing(shtName) Then
        If sht.Name = shtName Then
            If val Then
                If sht.Visible <> xlSheetVeryHidden Then sht.Visible = xlSheetVeryHidden
            Else
                If sht.Visible = xlSheetVeryHidden Then sht.Visible = xlSheetVisible
            End If
        End If
    Else
        If val Then
            If sht.Visible <> xlSheetVeryHidden Then sht.Visible = xlSheetVeryHidden
        Else
            If sht.Visible = xlSheetVeryHidden Then sht.Visible = xlSheetVisible
        End If
    End If
Next sht
End Sub

Private Sub prepareNew()
Dim v(2) As String
Dim i As Integer

v(0) = "BM_OLD"
v(1) = "Results"
v(2) = "MWG"


For i = LBound(v) To UBound(v)
    hideSheet True, v(i)
Next i

End Sub

Sub validateBM()
Dim c As Range
Dim rng As Range
Dim lastRow As Long
Dim sht As Worksheet
Dim i As Integer
Dim ord As Long
Dim duplicates() As Long
Dim firstAddress As String
Dim found As Boolean

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets("Operations sequence")

lastRow = sht.Range("E:E").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row

Set rng = sht.Range("O4:O" & lastRow)

For i = 4 To lastRow
    ord = sht.Range("O" & i)
    If ord <> 0 Then
        If i < lastRow Then
            Set c = sht.Range("O" & i + 1 & ":O" & lastRow).Find(ord, searchorder:=xlByRows, SearchDirection:=xlNext, Lookat:=xlWhole)
            If Not c Is Nothing Then
                If isArrayEmpty(duplicates) Then
                    ReDim duplicates(0) As Long
                    duplicates(0) = ord
                Else
                    ReDim Preserve duplicates(UBound(duplicates) + 1) As Long
                    duplicates(UBound(duplicates)) = ord
                End If
            End If
        End If
    End If
Next i

If Not isArrayEmpty(duplicates) Then
    For Each c In rng
        found = False
        ord = c
        If c <> 0 Then
            For i = LBound(duplicates) To UBound(duplicates)
                If duplicates(i) = ord Then
                    found = True
                    Exit For
                End If
            Next i
            If found Then c.Interior.Color = vbYellow
        End If
    Next c
    MsgBox "One or more ZFIN order(s) point to more than 1 ZFOR order. The orders have been marked in yellow and will wait for you to resolve the conflict. In order to resolve it, right-click on inappropriate ZFIN order and choose ""Remove connection""." & vbNewLine & "If you don't do that, the total result will most likely be inflated.", vbOKOnly + vbExclamation
End If

exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""validateBM"". Error number: " & Err.Number & ", " & Err.Description, vbOKOnly + vbCritical, "Error"
Resume exit_here

End Sub

Public Sub removeConnection()
Dim c As Range
Dim c1 As Range
Dim SQL As String
Dim str As String
Dim zfinOrd As Long
Dim zforOrd As Long
Dim sht As Worksheet
Dim i As Integer

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets("Operations sequence")

If propertyExists("userId") And propertyExists("isUserLogged") Then
    If ThisWorkbook.CustomDocumentProperties("userId") > 0 And ThisWorkbook.CustomDocumentProperties("isUserLogged") = True Then
        If authorize(48) Then
            Set c = ActiveCell
            If c.Column <> 15 Then
                MsgBox "This operation can only be carried on in column ""O"".", vbInformation + vbOKOnly, "Wrong column"
            Else
                 If Not IsNumeric(c) Then
                     MsgBox "This operation must be carried on over ZFIN order number! Selected cell doesn't seem to contain valid order number.", vbOKOnly + vbInformation, "No order number selected"
                 Else
                     zfinOrd = c
                     If zfinOrd = 0 Then
                         MsgBox "This operation must be carried on over ZFIN order number! Selected cell doesn't seem to contain valid order number.", vbOKOnly + vbInformation, "No order number selected"
                     Else
                         i = 0
                         Set c1 = c.Offset(0, -12)
                            str = ZforOrderString4Zfin(c)
                         If Not IsNumeric(c1) Then
                             MsgBox "Something went wrong. I can't find proper ZFOR order number for selected ZFIN number", vbOKOnly + vbInformation, "Something went wrong"
                         Else
                            updateConnection
                             zforOrd = c1
                             SQL = "UPDATE od SET isRemoved=1, RemovedOn=CURRENT_TIMESTAMP " _
                             & "FROM tbOrders o LEFT JOIN tbOrderDep od ON od.zfinOrder = o.orderId LEFT JOIN tbOrders o1 ON o1.orderId = od.zforOrder " _
                             & "WHERE o.sapId = " & zfinOrd & " And o1.sapId IN (" & str & ")"
                             adoConn.Execute SQL
                             MsgBox "This connection has been removed. Please update the report", vbInformation + vbOKOnly, "Success"
                         End If
        
                     End If
                 End If
            End If
        End If
    Else
        MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
        logger.Show
    End If
Else
    MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
    logger.Show
End If

exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""removeConnection"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Function ZforOrderString4Zfin(c As Range) As String
Dim rngStart As Range
Dim rngEnd As Range
Dim rng As Range
Dim i As Integer
Dim str As String
Dim nStr As String

Set rng = c.Offset(0, -14)
str = ""

If rng.MergeCells Then
    Set rng = rng.MergeArea
    Set rngStart = rng.Cells(1, 1)
    Set rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)
    For i = rngStart.row To rngEnd.row
        nStr = c.Worksheet.Range("C" & i)
        If nStr <> "" Then
            If InStr(1, str, nStr, vbTextCompare) = 0 Then
                str = str & nStr & ","
            End If
        End If
    Next i
    If Len(str) > 0 Then str = Left(str, Len(str) - 1)
    ZforOrderString4Zfin = str
Else
    ZforOrderString4Zfin = c.Offset(0, -12)
End If

End Function

Public Sub cancelOrder()
Dim c As Range
Dim c1 As Range
Dim SQL As String
Dim zfinOrd As Long
Dim zforOrd As Long
Dim sht As Worksheet
Dim i As Integer

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets("Operations sequence")

If propertyExists("userId") And propertyExists("isUserLogged") Then
    If ThisWorkbook.CustomDocumentProperties("userId") > 0 And ThisWorkbook.CustomDocumentProperties("isUserLogged") = True Then
        If authorize(48) Then
            Set c = ActiveCell
            If c.Column <> 15 Then
                MsgBox "This operation can only be carried on in column ""O"".", vbInformation + vbOKOnly, "Wrong column"
            Else
                 If Not IsNumeric(c) Then
                     MsgBox "This operation must be carried on over ZFIN order number! Selected cell doesn't seem to contain valid order number.", vbOKOnly + vbInformation, "No order number selected"
                 Else
                     zfinOrd = c
                     If zfinOrd = 0 Then
                         MsgBox "This operation must be carried on over ZFIN order number! Selected cell doesn't seem to contain valid order number.", vbOKOnly + vbInformation, "No order number selected"
                     Else
                         i = 0
                         Set c1 = c.Offset(0, -12)
                         Do Until c1 <> ""
                             i = i - 1
                             Set c1 = c.Offset(i, -12)
                         Loop
                         If Not IsNumeric(c1) Then
                             MsgBox "Something went wrong. I can't find proper ZFOR order number for selected ZFIN number", vbOKOnly + vbInformation, "Something went wrong"
                         Else
                            updateConnection
                             zforOrd = c1
                             SQL = "UPDATE tbOrders SET isCancelled=1 " _
                             & "FROM tbOrders " _
                             & "WHERE sapId = " & zfinOrd
                             adoConn.Execute SQL
                             SQL = "UPDATE od Set od.isRemoved = 1 " _
                            & "FROM tbOrderDep od INNER JOIN tbOrders o ON o.orderId=od.zfinOrder " _
                            & "WHERE o.sapId = " & zfinOrd
                            adoConn.Execute SQL
                             MsgBox "This order has been marked as cancelled. Please update the report", vbInformation + vbOKOnly, "Success"
                         End If
        
                     End If
                 End If
            End If
        End If
    Else
        MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
        logger.Show
    End If
Else
    MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
    logger.Show
End If

exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""cancelOrder"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Sub correctOrder()
Dim c As Range
Dim c1 As Range
Dim SQL As String
Dim zfinOrd As Long
Dim zforOrd As Long
Dim sht As Worksheet
Dim i As Integer

On Error GoTo err_trap

Set sht = ThisWorkbook.Sheets("Operations sequence")

If propertyExists("userId") And propertyExists("isUserLogged") Then
    If ThisWorkbook.CustomDocumentProperties("userId") > 0 And ThisWorkbook.CustomDocumentProperties("isUserLogged") = True Then
        If authorize(48) Then
            Set c = ActiveCell
            If c.Column <> 15 And c.Column <> 3 And c.Column <> 5 And c.Column <> 9 Then
                MsgBox "This operation can only be carried out on columns containing order number (either ZFOR or ZFIN) or batch number. Please select column ""C"", ""E"", ""I"", ""O"" or ""S""", vbInformation + vbOKOnly, "Wrong column"
            Else
                 If Not IsNumeric(c) Then
                     MsgBox "This operation can only be carried out on columns containing order number (either ZFOR or ZFIN) or batch number. Selected cell doesn't seem to contain valid order number.", vbOKOnly + vbInformation, "No number selected"
                 Else
                     zfinOrd = c
                     If zfinOrd = 0 Then
                         MsgBox "This operation can only be carried out on columns containing order number (either ZFOR or ZFIN) or batch number. Selected cell doesn't seem to contain valid order number.", vbOKOnly + vbInformation, "No number selected"
                     Else
                         correct2order.Show
                     End If
                 End If
            End If
        End If
    Else
        MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
        logger.Show
    End If
Else
    MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
    logger.Show
End If

exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""cancelOrder"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Function inArray(item As Variant, arr As Variant) As Variant
Dim bool As Variant
Dim i As Integer
Dim ind As Integer

bool = False

If Not isArrayEmpty(arr) Then
    For i = LBound(arr) To UBound(arr)
        If arr(i) = item Then
            bool = i
            Exit For
        End If
    Next i
End If

inArray = bool

End Function

Public Function arrayToString(arr As Variant) As String
Dim i As Integer

If isArrayEmpty(arr) Then
    arrayToString = "Array is empty"
Else
    For i = LBound(arr) To UBound(arr)
        arrayToString = arrayToString & arr(i) & ", "
    Next i
    arrayToString = Left(arrayToString, Len(arrayToString) - 2)
End If

End Function

Public Sub roundToProducts()
Dim rng As Range
Dim c As Range
Dim i As Integer
Dim sht As Worksheet
Dim index As Long
Dim green As Double
Dim perc As Double
Dim lastRow As Integer

ThisWorkbook.Sheets("X").Cells.clear

Set sht = ThisWorkbook.Sheets("Operations sequence")
lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
If sht.Cells(lastRow, 1).MergeArea.Cells.Count > 1 Then
    lastRow = lastRow + sht.Cells(lastRow, 1).MergeArea.Cells.Count - 1
End If

For i = 4 To lastRow
    index = sht.Cells(i, 1)
    perc = sht.Range("AF" & i)
    If sht.Cells(i, 1).MergeArea.Cells.Count > 1 Then
        green = WorksheetFunction.sum(Range(sht.Cells(i, 4), sht.Cells(i + sht.Cells(i, 1).MergeArea.Cells.Count - 1, 4)))
        i = i + sht.Cells(i, 1).MergeArea.Cells.Count - 1
    Else
        green = sht.Cells(i, 4)
    End If
    saveToRound index, green, perc
Next i

End Sub

Public Sub saveToRound(ind As Long, green As Double, perc As Double)
Dim sht As Worksheet
Dim i As Integer
Dim index As Long

Set sht = ThisWorkbook.Sheets("X")

If sht.Cells(1, 1) = "" Then
    sht.Cells(1, 1) = "ZFOR"
    sht.Cells(1, 2) = "Green [KG]"
    sht.Cells(1, 3) = "Total loss vs BOM [%]"
End If

For i = 2 To 10000
    index = sht.Cells(i, 1)
    If index = 0 Then
        sht.Cells(i, 1) = ind
        sht.Cells(i, 2) = green
        sht.Cells(i, 3) = perc
        Exit For
    End If
Next

End Sub

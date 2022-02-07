VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlendKeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Integer
Private pBlends As New Collection
Private pWeek As Integer
Private pMonth As Integer
Private pYear As Integer
Private pTotalGreen As Double
Private pTotalRoast As Double
Private pSessionId As Integer
Private pRework As Double
Private pReworkAtPackingBeans As Double
Private pReworkAtPackingGround As Double

Public Property Get reworkValue() As Double
Dim b As clsBlend
Dim val As Double

For Each b In pBlends
    If b.isConsistent And b.inScope Then
        val = val + b.reworkValue
    End If
Next b

reworkValue = val

End Property

Public Property Let SessionId(val As Integer)
pSessionId = val
End Property

Public Property Get SessionId() As Integer
SessionId = pSessionId
End Property

Public Property Get roastGroundScrap() As Double
Dim b As clsBlend
Dim sum As Double
Dim counter As Integer

For Each b In pBlends
    sum = sum + b.roastGroundScrap
    counter = counter + 1
Next b

roastGroundScrap = sum / counter

End Property

Public Property Get packScrap() As Double
Dim b As clsBlend
Dim sum As Double
Dim counter As Integer

For Each b In pBlends
    sum = sum + b.packScrap
    counter = counter + 1
Next b

packScrap = sum / counter

End Property

Public Property Get totalScrap() As Double
Dim b As clsBlend
Dim sum As Double
Dim counter As Integer

For Each b In pBlends
    sum = sum + b.totalScrap
    counter = counter + 1
Next b

totalScrap = sum / counter

End Property

Public Property Let totalGreen(val As Double)
pTotalGreen = val
End Property

Public Property Get totalGreen() As Double
Dim b As clsBlend

If pTotalGreen > 0 Then
    totalGreen = pTotalGreen
Else
    For Each b In pBlends
        totalGreen = totalGreen + b.getGreen
    Next b
End If

End Property

Public Property Let TotalRoast(val As Double)
pTotalRoast = val
End Property

Public Property Get TotalRoast() As Double
Dim b As clsBlend

If pTotalRoast > 0 Then
    TotalRoast = pTotalRoast
Else
    For Each b In pBlends
        TotalRoast = TotalRoast + b.getRoasted
    Next b
End If

End Property

Public Property Get CountedGreen() As Double
Dim b As clsBlend

For Each b In pBlends
    If b.isConsistent Then CountedGreen = CountedGreen + b.getGreen
Next b

End Property

Public Property Get CountedRoast() As Double
Dim b As clsBlend

For Each b In pBlends
    If b.isConsistent Then CountedRoast = CountedRoast + b.getRoasted
Next b

End Property

Public Property Let week(val As Integer)
pWeek = val
End Property

Public Property Get week() As Integer
week = pWeek
End Property

Public Property Let year(val As Integer)
pYear = val
End Property

Public Property Get year() As Integer
year = pYear
End Property

Public Property Let month(val As Integer)
pMonth = val
End Property

Public Property Get month() As Integer
month = pMonth
End Property

Public Sub append(b As clsBlend)
pBlends.Add b, CStr(b.index)
End Sub

Public Property Let Id(val As Integer)
pId = val
End Property

Public Property Get Id() As Integer
Id = pId
End Property

Public Sub order(processType As String, valueType As String, orderType As String, vsBom As Boolean, Optional roaster As Variant, Optional vsAvg As Variant)
'eg order("r","%","DESC") to get collection of all blends ordered by loss on roasting in % DESC
Dim blends As New Collection
Dim blendsSorted As New Collection
Dim b As clsBlend
Dim highest As Double
Dim highestBlend As Long
Dim i As Integer
Dim x As Integer
Dim this As Double

On Error GoTo err_trap

Set blends = pBlends

If IsMissing(vsAvg) Then vsAvg = False

Do Until blends.Count = 0
    i = i + 1
    highest = 0
    highestBlend = 0
    For Each b In blends
        x = x + 1
        If Not vsBom Then
            If vsAvg Then
                this = Round(b.getLoss(processType, "%", , True), 5)
            Else
                If Not IsNull(roaster) Then
                    this = Round(b.getLoss(processType, valueType, roaster), 5)
                Else
                    this = Round(b.getLoss(processType, valueType), 5)
                End If
            End If
        Else
            this = Round(b.getRealVsBom(processType, valueType), 5)
        End If
        If highestBlend = 0 Then
            'the very first blend
            highest = this
            highestBlend = b.index
        End If
        If orderType = "DESC" Then
            If this > highest Then
                highest = this
                highestBlend = b.index
            End If
        Else
            If this < highest Then
                highest = this
                highestBlend = b.index
            End If
        End If
    Next b
    blendsSorted.Add blends(CStr(highestBlend)), CStr(highestBlend)
    blends.Remove CStr(highestBlend)
Loop

Set pBlends = blendsSorted

exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""Order"" of clsBlendKeeper. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Function blends(Optional ind As Variant) As Variant
If Not IsMissing(ind) Then
    Set blends = pBlends(CStr(ind))
Else
    Set blends = pBlends
End If
End Function

Public Sub display()
Dim stroke As Boolean
Dim bool As Boolean
Dim totalZfinOrders As Integer
Dim totalZforOrders As Integer
Dim i As Integer
Dim step As Integer
Dim step1 As Integer
Dim sht As Worksheet
Dim rng As Range
Dim nBlend As clsBlend
Dim m As Integer

i = 4

Set sht = ThisWorkbook.Sheets("Operations sequence")

For Each nBlend In pBlends
    If nBlend.inScope Then
        If i > 4 Then
            i = ii + n
        End If
        ii = i
        i2 = i
        stroke = False
        If nBlend.numberOfOrders > 0 Then
            m = nBlend.numberOfOrders
            n = nBlend.numberOfZfinOrders
            bool = nBlend.isConsistent
            If bool = False Then stroke = True
            If n > 1 Then
                Set rng = sht.Range("A" & i & ":A" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("B" & i & ":B" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("G" & i & ":G" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("H" & i & ":H" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("K" & i & ":K" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("L" & i & ":L" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("Q" & i & ":Q" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("R" & i & ":R" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("W" & i & ":W" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("X" & i & ":X" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("Y" & i & ":Y" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("Z" & i & ":Z" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AA" & i & ":AA" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AB" & i & ":AB" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AC" & i & ":AC" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AD" & i & ":AD" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AE" & i & ":AE" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AF" & i & ":AF" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AG" & i & ":AG" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AH" & i & ":AH" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AI" & i & ":AI" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AJ" & i & ":AJ" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AK" & i & ":AK" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AL" & i & ":AL" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AM" & i & ":AM" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AN" & i & ":AN" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AO" & i & ":AO" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AP" & i & ":AP" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AQ" & i & ":AQ" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AR" & i & ":AR" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AS" & i & ":AS" & i + n - 1)
                rng.Merge
                Set rng = sht.Range("AT" & i & ":AT" & i + n - 1)
                rng.Merge
                If stroke Then sht.Range("A" & i & ":AT" & i + n - 1).Font.Strikethrough = True
            Else
                If stroke Then sht.Range("A" & i & ":AT" & i).Font.Strikethrough = True
            End If
            If m >= 1 And n > 0 Then
                step = n / m
                If step * m > n And step = 1 Then
                    MsgBox "Something wrong in BlendKeeper's display. Number of zfor orders=" & m & " and number of zfin orders=" & n & ". What step should it be then? Zfor: " & nBlend.index, vbOKOnly + vbCritical
                ElseIf step * m > n Then
                    step = step - 1
                End If
                step1 = step
                i1 = m
                Do While i1 > 0
                    If i1 = 1 Then step = n - ((m - 1) * step)
                    Set rng = sht.Range("C" & i2 & ":C" & i2 + step - 1)
                    rng.Merge
                    Set rng = sht.Range("D" & i2 & ":D" & i2 + step - 1)
                    rng.Merge
                    Set rng = sht.Range("E" & i2 & ":E" & i2 + step - 1)
                    rng.Merge
                    Set rng = sht.Range("F" & i2 & ":F" & i2 + step - 1)
                    rng.Merge
                    Set rng = sht.Range("I" & i2 & ":I" & i2 + step - 1)
                    rng.Merge
                    Set rng = sht.Range("J" & i2 & ":J" & i2 + step - 1)
                    rng.Merge
'                    Set rng = sht.Range("AN" & i2 & ":AN" & i2 + step - 1)
'                    rng.Merge
                    i2 = i2 + step
                    i1 = i1 - 1
                Loop
                sht.Range("A" & i) = nBlend.index
                sht.Range("B" & i) = nBlend.Name
                sht.Range("G" & i) = -1 * Round(nBlend.getGreen - nBlend.getRoasted, 1)
                sht.Range("H" & i) = -1 * Round(Round(nBlend.getGreen - nBlend.getRoasted, 1) / nBlend.getGreen, 4)
                sht.Range("K" & i) = -1 * Round(nBlend.getRoasted + nBlend.rework - nBlend.getGround, 1)
                sht.Range("L" & i) = -1 * Round(Round(nBlend.getRoasted + nBlend.rework - nBlend.getGround, 1) / (nBlend.getRoasted + nBlend.rework), 4)
                sht.Range("AN" & i) = Round(nBlend.rework, 1)
                sht.Range("AO" & i) = Round(nBlend.reworkAtPacking, 1)
                sht.Range("Q" & i) = -1 * Round(nBlend.getGround + nBlend.reworkAtPacking - nBlend.getStocked, 1)
                If nBlend.getGround = 0 Then
                    sht.Range("R" & i) = 0
                Else
                    sht.Range("R" & i) = -1 * Round(Round(nBlend.getGround + nBlend.reworkAtPacking - nBlend.getStocked, 1) / (nBlend.getGround + nBlend.reworkAtPacking), 4)
                End If
                If IsArray(nBlend.getOrders) Then
                    For g = LBound(nBlend.getOrders, 2) To UBound(nBlend.getOrders, 2)
                        sht.Range("C" & i) = nBlend.getOrders()(0, g)
                        sht.Range("D" & i) = Round(nBlend.getOrders()(1, g), 1)
                        sht.Range("E" & i) = nBlend.getOrders()(0, g)
                        sht.Range("F" & i) = Round(nBlend.getOrders()(2, g), 1)
                        If nBlend.IsBeans Then
                            sht.Range("J" & i) = Round(nBlend.getOrders()(2, g), 1)
                        Else
                            sht.Range("J" & i) = Round(nBlend.getOrders()(3, g), 1)
                        End If
                        sht.Range("I" & i) = nBlend.getOrders()(0, g)
                        'If g = UBound(nBlend.getOrders, 2) - 1 Then step1 = step
                        i = i + step1
                    Next g
                End If
                i = ii
                For Each nzfin In nBlend.getZfins
                    If IsArray(nzfin.getOrders) Then
                        Set rng = sht.Range("M" & i & ":M" & i + nzfin.numberOfOrders - 1)
                        rng.Merge
                        Set rng = sht.Range("N" & i & ":N" & i + nzfin.numberOfOrders - 1)
                        rng.Merge
                        Set rng = sht.Range("U" & i & ":U" & i + nzfin.numberOfOrders - 1)
                        rng.Merge
                        Set rng = sht.Range("V" & i & ":V" & i + nzfin.numberOfOrders - 1)
                        rng.Merge
                        sht.Range("M" & i) = nzfin.zfinIndex
                        sht.Range("N" & i) = nzfin.zfinName
                        sht.Range("U" & i) = -1 * (nzfin.getPacked - nzfin.getStocked)
                        If nzfin.getPacked > 0 Then
                            sht.Range("V" & i) = -1 * Round(Round(nzfin.getPacked - nzfin.getStocked, 1) / nzfin.getPacked, 4)
                        End If
    
                        For g = LBound(nzfin.getOrders, 2) To UBound(nzfin.getOrders, 2)
                            sht.Range("O" & i) = nzfin.getOrders()(0, g)
                            sht.Range("S" & i) = nzfin.getOrders()(1, g)
                            sht.Range("P" & i) = nzfin.getOrders()(2, g)
                            sht.Range("T" & i) = nzfin.getOrders()(3, g)
                            i = i + 1
                        Next g
                    End If
                Next nzfin
                If m > 1 Or n > 1 Then
                    If n > m Then
                        sht.Range("W" & i - n) = -1 * Round(nBlend.getGreen + nBlend.rework + nBlend.reworkAtPacking - nBlend.getStocked, 1)
                        sht.Range("X" & i - n) = -1 * Round(Round(nBlend.getGreen + nBlend.rework + nBlend.reworkAtPacking - nBlend.getStocked, 1) / (nBlend.getGreen + nBlend.rework + nBlend.reworkAtPacking), 4)
                        sht.Range("Y" & i - n) = -1 * Round(nBlend.getRoasted - nBlend.getStocked, 1)
                        sht.Range("Z" & i - n) = -1 * Round(Round(nBlend.getRoasted - nBlend.getStocked, 1) / nBlend.getRoasted, 4)
                        sht.Range("AA" & i - n) = -1 * nBlend.roastGroundScrap / 100
                        sht.Range("AB" & i - n) = -1 * nBlend.packScrap / 100
                        sht.Range("AC" & i - n) = -1 * nBlend.totalScrap / 100
                        sht.Range("AD" & i - n) = nBlend.getRealVsBom("r+g", "%")
                        sht.Range("AE" & i - n) = nBlend.getRealVsBom("p", "%")
                        sht.Range("AF" & i - n) = nBlend.getRealVsBom("t", "%")
                        sht.Range("AG" & i - n) = nBlend.getLoss("r", "kg", 3000)
                        sht.Range("AH" & i - n) = nBlend.getLoss("r", "%", 3000)
                        sht.Range("AI" & i - n) = nBlend.getLoss("r", "kg", 4000)
                        sht.Range("AJ" & i - n) = nBlend.getLoss("r", "%", 4000)
                        sht.Range("AK" & i - n) = Round((nBlend.RoastedValue + nBlend.reworkValue) / 1000, 2)
                        sht.Range("AL" & i - n) = Round(nBlend.PackedValue / 1000, 2)
                        sht.Range("AM" & i - n) = Round((nBlend.getLoss("g+p", "$") - nBlend.reworkValue) / 1000, 2)
                        sht.Range("AP" & i - n) = Round((nBlend.getLoss("r", "%", , True)), 2) / 100
                        sht.Range("AQ" & i - n) = Round((nBlend.getLoss("g", "%", , True)), 2) / 100
                        sht.Range("AR" & i - n) = Round((nBlend.getLoss("p", "%", , True)), 2) / 100
                        sht.Range("AS" & i - n) = Round((nBlend.getLoss("g+p", "%", , True)), 2) / 100
                        sht.Range("AT" & i - n) = Round((nBlend.getLoss("t", "%", , True)), 2) / 100
                    Else
                        sht.Range("W" & i - m) = -1 * Round(nBlend.getGreen - nBlend.getStocked, 1)
                        sht.Range("X" & i - m) = -1 * Round(Round(nBlend.getGreen - nBlend.getStocked, 1) / nBlend.getGreen, 4)
                        sht.Range("Y" & i - m) = -1 * Round(nBlend.getRoasted - nBlend.getStocked, 1)
                        sht.Range("Z" & i - m) = -1 * Round(Round(nBlend.getRoasted - nBlend.getStocked, 1) / nBlend.getRoasted, 4)
                        sht.Range("AA" & i - m) = -1 * nBlend.roastGroundScrap / 100
                        sht.Range("AB" & i - m) = -1 * nBlend.packScrap / 100
                        sht.Range("AC" & i - m) = -1 * nBlend.totalScrap / 100
                        sht.Range("AD" & i - n) = nBlend.getRealVsBom("r+g", "%")
                        sht.Range("AE" & i - n) = nBlend.getRealVsBom("p", "%")
                        sht.Range("AF" & i - n) = nBlend.getRealVsBom("t", "%")
                        sht.Range("AG" & i - n) = nBlend.getLoss("r", "kg", 3000)
                        sht.Range("AH" & i - n) = nBlend.getLoss("r", "%", 3000)
                        sht.Range("AI" & i - n) = nBlend.getLoss("r", "kg", 4000)
                        sht.Range("AJ" & i - n) = nBlend.getLoss("r", "%", 4000)
                        sht.Range("AK" & i - n) = Round((nBlend.RoastedValue) / 1000, 2)
                        sht.Range("AL" & i - n) = Round(nBlend.PackedValue / 1000, 2)
                        sht.Range("AM" & i - n) = Round(nBlend.getLoss("g+p", "$") / 1000, 2)
                        sht.Range("AP" & i - n) = Round((nBlend.getLoss("r", "%", , True)), 2) / 100
                        sht.Range("AQ" & i - n) = Round((nBlend.getLoss("g", "%", , True)), 2) / 100
                        sht.Range("AR" & i - n) = Round((nBlend.getLoss("p", "%", , True)), 2) / 100
                        sht.Range("AS" & i - n) = Round((nBlend.getLoss("g+p", "%", , True)), 2) / 100
                        sht.Range("AT" & i - n) = Round((nBlend.getLoss("t", "%", , True)), 2) / 100
                    End If
                Else
                    sht.Range("W" & i - 1) = -1 * Round(nBlend.getGreen - nBlend.getStocked, 1)
                    sht.Range("X" & i - 1) = -1 * Round(Round(nBlend.getGreen - nBlend.getStocked, 1) / nBlend.getGreen, 4)
                    sht.Range("Y" & i - 1) = -1 * Round(nBlend.getRoasted - nBlend.getStocked, 1)
                    sht.Range("Z" & i - 1) = -1 * Round(Round(nBlend.getRoasted - nBlend.getStocked, 1) / nBlend.getRoasted, 4)
                    sht.Range("AA" & i - 1) = -1 * nBlend.roastGroundScrap / 100
                    sht.Range("AB" & i - 1) = -1 * nBlend.packScrap / 100
                    sht.Range("AC" & i - 1) = -1 * nBlend.totalScrap / 100
                    sht.Range("AD" & i - 1) = nBlend.getRealVsBom("r+g", "%")
                    sht.Range("AE" & i - 1) = nBlend.getRealVsBom("p", "%")
                    sht.Range("AF" & i - 1) = nBlend.getRealVsBom("t", "%")
                    sht.Range("AG" & i - 1) = nBlend.getLoss("r", "kg", 3000)
                    sht.Range("AH" & i - 1) = nBlend.getLoss("r", "%", 3000)
                    sht.Range("AI" & i - 1) = nBlend.getLoss("r", "kg", 4000)
                    sht.Range("AJ" & i - 1) = nBlend.getLoss("r", "%", 4000)
                    sht.Range("AK" & i - 1) = Round((nBlend.RoastedValue) / 1000, 2)
                    sht.Range("AL" & i - 1) = Round(nBlend.PackedValue / 1000, 2)
                    sht.Range("AM" & i - 1) = Round(nBlend.getLoss("g+p", "$") / 1000, 2)
                    sht.Range("AP" & i - 1) = Round((nBlend.getLoss("r", "%", , True)), 2) / 100
                    sht.Range("AQ" & i - 1) = Round((nBlend.getLoss("g", "%", , True)), 2) / 100
                    sht.Range("AR" & i - 1) = Round((nBlend.getLoss("p", "%", , True)), 2) / 100
                    sht.Range("AS" & i - 1) = Round((nBlend.getLoss("g+p", "%", , True)), 2) / 100
                    sht.Range("AT" & i - 1) = Round((nBlend.getLoss("t", "%", , True)), 2) / 100
                    'sht.Range("AD" & i - N) = (-1 * ROUND(ROUND(nBlend.getGreen - nBlend.getGround, 1) / nBlend.getGreen, 4)) - (-1 * nBlend.roastGroundScrap / 100)
                End If
                If step = step1 Then
                    i = i + (step - step1)
                Else
                    i = i + (step - step1) - 1
                End If
            Else
                MsgBox "Funny situtation", vbInformation + vbOKOnly, "Funny"
            End If
        Else
            n = 0
        End If
    End If
Next nBlend
sht.Range("H" & 3 & ":H" & i).NumberFormat = "0.00%"
sht.Range("L" & 3 & ":L" & i).NumberFormat = "0.00%"
sht.Range("R" & 3 & ":R" & i).NumberFormat = "0.00%"
sht.Range("V" & 3 & ":V" & i).NumberFormat = "0.00%"
sht.Range("X" & 3 & ":X" & i).NumberFormat = "0.00%"
sht.Range("Z" & 3 & ":Z" & i).NumberFormat = "0.00%"
sht.Range("AA" & 3 & ":AF" & i).NumberFormat = "0.00%"
sht.Range("AH" & 3 & ":AH" & i).NumberFormat = "0.00%"
sht.Range("AJ" & 3 & ":AJ" & i).NumberFormat = "0.00%"
sht.Range("AM" & 3 & ":AM" & i).NumberFormat = "#,##0.00"
sht.Range("D3:D" & i).NumberFormat = "#,##0.00"
sht.Range("F3:F" & i).NumberFormat = "#,##0.00"
sht.Range("G3:G" & i).NumberFormat = "#,##0.00"
sht.Range("J3:J" & i).NumberFormat = "#,##0.00"
sht.Range("K3:K" & i).NumberFormat = "#,##0.00"
sht.Range("P3:P" & i).NumberFormat = "#,##0.00"
sht.Range("T3:T" & i).NumberFormat = "#,##0.00"
sht.Range("U3:U" & i).NumberFormat = "#,##0.00"
sht.Range("W3:W" & i).NumberFormat = "#,##0.00"
sht.Range("Y3:Y" & i).NumberFormat = "#,##0.00"
sht.Range("AN3:AN" & i).NumberFormat = "#,##0.00"
sht.Range("AO3:AO" & i).NumberFormat = "#,##0.00"
sht.Range("AP" & 3 & ":AT" & i).NumberFormat = "0.00%"
sht.Range("A" & 3 & ":F" & i).HorizontalAlignment = xlGeneral
sht.Range("A" & 3 & ":AT" & i).VerticalAlignment = xlCenter
End Sub

Public Function getRoastedValue() As Double
Dim b As clsBlend
Dim totalRoasted As Double

On Error GoTo err_trap

For Each b In pBlends
    If b.isConsistent Then
        totalRoasted = totalRoasted + b.RoastedValue
    End If
Next b

exit_here:
getRoastedValue = totalRoasted
Exit Function

err_trap:
MsgBox "Error in ""getRoastedValue"" of clsBlendKeeper. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Function getPackedValue() As Double
Dim b As clsBlend
Dim totalPacked As Double

On Error GoTo err_trap

For Each b In pBlends
    If b.isConsistent Then
        totalPacked = totalPacked + b.PackedValue
    End If
Next b

exit_here:
getPackedValue = totalPacked
Exit Function

err_trap:
MsgBox "Error in ""getPackedValue"" of clsBlendKeeper. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Function getRoastingScrap(roaster As Integer, format As String) As Double
Dim b As clsBlend
Dim totalRoasted As Double
Dim totalGreen As Double

On Error GoTo err_trap

For Each b In pBlends
    If b.consistent Then
        totalRoasted = totalRoasted + b.getRoasted(, roaster)
        totalGreen = totalGreen + b.getGreen(, roaster)
    End If
Next b

If totalGreen > 0 Then
    If format = "%" Then
        getRoastingScrap = 1 - (totalRoasted / totalGreen)
    Else
        getRoastingScrap = totalGreen - totalRoasted
    End If
Else
    getRoastingScrap = 0
End If

exit_here:
Exit Function

err_trap:
MsgBox "Error in ""getRoastingScrap"" of clsBlendKeeper. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Sub getScraps()
Dim rs As ADODB.Recordset
Dim SQL As String

updateConnection

SQL = "SELECT SUM(bomy.amount) as TotalKg,ROUND(((SUM(bomy.amount)-1000)/SUM(bomy.amount))*100,2) as RandGScrap, mainZfin.zfinIndex, mainZfin.zfinName, freshBom.dateAdded " _
    & "FROM tbBom bomy RIGHT JOIN " _
    & "(SELECT oBom.zfinId,  MAX(oBom.bomRecId) as bomRecId, MAX(oBom.dateAdded) as dateAdded FROM " _
    & "(SELECT iBom.bomRecId, zfinId, br.dateAdded FROM tbBomReconciliation br JOIN ( " _
    & "SELECT bomRecId, bom.zfinId " _
    & "FROM tbBom bom " _
    & "GROUP BY bomRecId, bom.zfinId) iBom ON iBom.bomRecId=br.bomRecId) oBom " _
    & "WHERE oBom.dateAdded <='" & Now & "' " _
    & "GROUP BY oBom.zfinId) freshBom ON freshBom.zfinId=bomy.zfinId AND freshBom.bomRecId=bomy.bomRecId " _
    & "LEFT JOIN tbUom u ON u.zfinId=bomy.zfinId LEFT JOIN tbZfin zfin ON zfin.zfinId=bomy.materialId " _
    & "LEFT JOIN tbZfin mainZfin ON mainZfin.zfinId=bomy.zfinId LEFT JOIN tbZfin zcom ON zcom.zfinId=bomy.materialId " _
    & "WHERE bomy.zfinId IN " _
    & "(SELECT zfor.zfinId FROM tbZfin zfor WHERE zfor.zfinType='zfor' AND zfor.prodStatus<>'OB' AND zfor.prodStatus<>'HI') " _
    & "GROUP BY mainZfin.zfinIndex, mainZfin.zfinName, freshBom.dateAdded"

Set rs = New ADODB.Recordset

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable

If Not rs.EOF Then
    Do Until rs.EOF
        If inCollection(CStr(rs.Fields("zfinIndex")), pBlends) Then
            'update zfor with R&G scrap
            pBlends(CStr(rs.Fields("zfinIndex"))).roastGroundScrap = rs.Fields("RandGScrap").Value
        End If
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

SQL = "SELECT conScrap.zfinId, conScrap.scrap, freshScrap.dateAdded, z.zfinIndex " _
    & "FROM tbComponentScrap conScrap LEFT JOIN tbZfin z ON z.zfinId=conScrap.zfinId RIGHT JOIN " _
    & "(SELECT oScrap.zfinId,  MAX(oScrap.scrapReconciliationId) as scrapRecId, MAX(oScrap.dateAdded) as dateAdded FROM " _
    & "(SELECT iScrap.scrapReconciliationId, zfinId, sr.dateAdded FROM tbScrapReconciliation sr JOIN " _
    & "(SELECT scrapReconciliationId, cs.zfinId FROM tbComponentScrap cs GROUP BY scrapReconciliationId, cs.zfinId) iScrap ON iScrap.scrapReconciliationId=sr.scrapReconciliationId) oScrap " _
    & "WHERE oScrap.dateAdded <='" & Now & "' " _
    & "GROUP BY oScrap.zfinId) freshScrap ON freshScrap.zfinId=conScrap.zfinId AND freshScrap.scrapRecId=conScrap.scrapReconciliationId " _
    & "WHERE z.zfinType='zfor'"

Set rs = New ADODB.Recordset

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable

If Not rs.EOF Then
    Do Until rs.EOF
        If inCollection(CStr(rs.Fields("zfinIndex")), pBlends) Then
            'update zfor with R&G scrap
            pBlends(CStr(rs.Fields("zfinIndex"))).packScrap = rs.Fields("scrap").Value
        End If
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

End Sub

Public Sub restoreFromSheet()
Dim sht As Worksheet
Dim y As Integer
Dim i As Integer
Dim c As Range
Dim c1 As Range
Dim b As clsBlend

Set sht = ThisWorkbook.Sheets("Operations sequence")

For y = 4 To 30000
    Set c = sht.Range("A" & y)
    If c.MergeCells = False And c.Value = Empty Then
        Exit For 'we've hit last row, let's finish
    Else
        If c.MergeCells = False Then
            'single-row blend, should be easy to handle
            Set b = New clsBlend
            b.index = CLng(c.Value)
            b.Name = c.Offset(0, 1)
            If c.Font.Strikethrough Then
                b.consistent = False
            Else
                b.consistent = True
            End If
            b.green = CDbl(c.Offset(0, 3))
            b.Roasted = CDbl(c.Offset(0, 5))
            b.Ground = CDbl(c.Offset(0, 9))
            b.Packed = CDbl(c.Offset(0, 15))
            b.Stocked = CDbl(c.Offset(0, 19))
            b.bomVsReal = c.Offset(0, 31)
            b.gpValueLoss = c.Offset(0, 38)
            b.rework = c.Offset(0, 39)
            b.reworkAtPacking = c.Offset(0, 40)
            pBlends.Add b, CStr(b.index)
        Else
            Set b = New clsBlend
            b.green = WorksheetFunction.sum(Range(sht.Range("D" & y), sht.Range("D" & y + c.MergeArea.Cells.Count - 1)))
            b.Roasted = WorksheetFunction.sum(Range(sht.Range("F" & y), sht.Range("F" & y + c.MergeArea.Cells.Count - 1)))
            b.Ground = WorksheetFunction.sum(Range(sht.Range("J" & y), sht.Range("J" & y + c.MergeArea.Cells.Count - 1)))
            b.Packed = WorksheetFunction.sum(Range(sht.Range("P" & y), sht.Range("P" & y + c.MergeArea.Cells.Count - 1)))
            b.Stocked = WorksheetFunction.sum(Range(sht.Range("T" & y), sht.Range("T" & y + c.MergeArea.Cells.Count - 1)))
            For i = y To y + c.MergeArea.Cells.Count - 1
                If sht.Range("A" & i).Value <> Empty Then
                    b.index = CLng(sht.Range("A" & i).Value)
                    b.Name = sht.Range("B" & i).Value
                    b.bomVsReal = sht.Range("AF" & i).Value
                    b.gpValueLoss = sht.Range("AM" & i).Value
                    b.rework = sht.Range("AN" & i).Value
                    b.reworkAtPacking = sht.Range("AO" & i).Value
                    If sht.Range("A" & i).Font.Strikethrough Then
                        b.consistent = False
                    Else
                        b.consistent = True
                    End If
                    Exit For
                End If
            Next i
            pBlends.Add b, CStr(b.index)
            y = y + c.MergeArea.Cells.Count - 1
        End If
    End If
Next y

End Sub

Public Sub ToOverview()
Dim nSht As Worksheet
Dim nWbook As Workbook
Dim b As clsBlend
Dim i As Integer

On Error GoTo err_trap

Set nWbook = Workbooks.Add
Set nSht = nWbook.Sheets.Add

initializeObjects


Me.restoreFromSheet
Sequencer.formatMe nSht

i = 3

For Each b In Me.blends
    With nSht
        .Range("A" & i).Value = b.index
        .Range("B" & i).Value = b.Name
        .Range("D" & i).Value = Round(b.getGreen, 2)
        .Range("F" & i).Value = Round(b.getRoasted, 2)
        .Range("G" & i).Value = Round(b.getRoasted - b.getGreen, 2)
        .Range("H" & i).Value = Round(((b.getRoasted - b.getGreen) / b.getGreen) * 100, 2)
        .Range("J" & i).Value = Round(b.getGround, 2)
        .Range("K" & i).Value = Round(b.getGround - b.getRoasted, 2)
        .Range("L" & i).Value = Round(((b.getGround - b.getRoasted) / b.getRoasted) * 100, 2)
        .Range("P" & i).Value = Round(b.getPacked, 2)
        .Range("Q" & i).Value = Round(b.getPacked - b.getGround, 2)
        .Range("R" & i).Value = Round(((b.getPacked - b.getGround) / b.getGround) * 100, 2)
        .Range("T" & i).Value = Round(b.getStocked, 2)
        .Range("U" & i).Value = Round(b.getStocked - b.getPacked, 2)
        .Range("V" & i).Value = Round(((b.getStocked - b.getPacked) / b.getPacked) * 100, 2)
        .Range("W" & i).Value = Round(b.getStocked - b.getGreen, 2)
        .Range("X" & i).Value = Round(((b.getStocked - b.getGreen) / b.getGreen) * 100, 2)
        .Range("Y" & i).Value = Round((b.getGround - b.getRoasted) + (b.getPacked - b.getGround), 2)
        .Range("Z" & i).Value = Round((((b.getGround - b.getRoasted) + (b.getPacked - b.getGround)) / b.getRoasted) * 100, 2)
        i = i + 1
    End With
Next b

exit_here:
Exit Sub

err_trap:
MsgBox "Error in ToOverview. Description: " & Err.Description, vbCritical + vbOKOnly, "Error occurred"
Resume exit_here

End Sub

Public Sub downloadCost()
Dim rs As ADODB.Recordset
Dim SQL As String

updateConnection

SQL = "SELECT z.zfinId, z.zfinIndex, z.zfinName, c.cost, c.CostLotSize, c.CostLotSizeUnit, cr.dateAdded " _
    & "FROM tbCosting c LEFT JOIN tbCostingReconciliation cr ON cr.CostingReconciliationId=c.reconciliationId LEFT JOIN tbZfin z ON z.zfinId=c.zfinId " _
    & "WHERE cr.CostingReconciliationId = (SELECT TOP(1) CostingReconciliationId FROM tbCostingReconciliation ORDER BY dateAdded DESC) AND (z.zfinType='zfor' OR z.zfinType='ZFOR')"

Set rs = New ADODB.Recordset

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If inCollection(rs.Fields("zfinIndex").Value, pBlends) Then
            pBlends(CStr(rs.Fields("zfinIndex").Value)).cost = rs.Fields("cost").Value
            pBlends(CStr(rs.Fields("zfinIndex").Value)).CostLotSize = rs.Fields("CostLotSize").Value
        End If
        rs.MoveNext
    Loop
End If

closeConnection

End Sub

Public Sub downloadAvg(Optional avgRange As Variant)
Dim rs As ADODB.Recordset
Dim SQL As String
Dim prevDate As Date
Dim prevWeek As Integer

If IsMissing(avgRange) Then
    avgRange = 180
End If

updateConnection

prevDate = DateAdd("d", -1 * avgRange, Date)
prevWeek = CInt(IsoWeekNumber(prevDate))


SQL = "SELECT z.zfinIndex, z.zfinName, " _
    & "AVG((bmo.roastingOut - bmo.roastingIn)/NULLIF(bmo.roastingIn,0)) * 100 as [rLoss], " _
    & "AVG((bmo.packingIn - (bmo.roastingOut+bmo.rework))/(NULLIF(bmo.roastingOut,0)+bmo.rework))*100 as [gLoss], " _
    & "AVG((bmo.packingOut - (bmo.packingIn+ISNULL(bmo.reworkAtPacking,0)))/(NULLIF(bmo.packingIn,0)+ISNULL(bmo.reworkAtPacking,0)))*100 as [pLoss], " _
    & "AVG((bmo.warehouseIn - (bmo.roastingOut+bmo.rework+ISNULL(bmo.reworkAtPacking,0)))/(NULLIF(bmo.roastingOut,0)+bmo.rework+ISNULL(bmo.reworkAtPacking,0)))*100 as [gpLoss], " _
    & "AVG((bmo.warehouseIn - (bmo.roastingIn+bmo.rework + ISNULL(bmo.reworkAtPacking,0)))/(NULLIF(bmo.roastingIn,0)+bmo.rework + ISNULL(bmo.reworkAtPacking,0)))*100 as [tLoss], " _
    & "AVG(bmo.rework) as [Rework], " _
    & "AVG(bmo.bomVsRealScrap)*100 as [BOMvsREAL] " _
    & "FROM tbZfin z " _
    & "LEFT JOIN tbBMOverview bmo ON bmo.zfinId = z.zfinId " _
    & "LEFT JOIN tbBM bm ON bm.bmId = bmo.bmId " _
    & "WHERE bm.bmWeek Is Not Null And bm.bmWeek > " & prevWeek & " And bm.bmYear >= " & ToYear(prevDate) & " And bmo.isConsistent = 1 " _
    & "GROUP BY z.zfinIndex, z.zfinName"

Set rs = New ADODB.Recordset

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If inCollection(rs.Fields("zfinIndex").Value, pBlends) Then
            pBlends(CStr(rs.Fields("zfinIndex").Value)).AvgRoastingLoss = rs.Fields("rLoss").Value
            pBlends(CStr(rs.Fields("zfinIndex").Value)).AvgGrindingLoss = rs.Fields("gLoss").Value
            pBlends(CStr(rs.Fields("zfinIndex").Value)).AvgPackingLoss = rs.Fields("pLoss").Value
            pBlends(CStr(rs.Fields("zfinIndex").Value)).AvgGPLoss = rs.Fields("gpLoss").Value
            pBlends(CStr(rs.Fields("zfinIndex").Value)).AvgTotalLoss = rs.Fields("tLoss").Value
            pBlends(CStr(rs.Fields("zfinIndex").Value)).AvgRework = rs.Fields("Rework").Value
            pBlends(CStr(rs.Fields("zfinIndex").Value)).AvgBomVsReal = rs.Fields("BOMvsREAL").Value
        End If
        rs.MoveNext
    Loop
End If

closeConnection
End Sub

Public Sub printHeaders()
Dim b As clsBlend

Debug.Print "Index" & "    " & "Nazwa" & "    " & "    " & "Zielona" & "    " & "Prażona" & "    " & "Mielona" & "    " & "Pakowana" & "    " & "Magazyn"
For Each b In pBlends
    Debug.Print b.index & "    " & b.Name & "    " & b.getGreen & "    " & b.getRoasted & "    " & b.getGround & "    " & b.getPacked & "    " & b.getStocked
Next b
End Sub

Public Function getBlends() As String
Dim b As clsBlend

For Each b In pBlends
    getBlends = getBlends & b.index & vbNewLine
Next b

End Function

Public Sub calculateRework()
Dim SQL As String
Dim zforList As String
Dim zfinList As String
Dim b As clsBlend

zforList = Me.getOrderList("zfor")
zfinList = Me.getOrderList("zfin")

updateConnection

SQL = "SELECT  main.[To order], main.Amount / (SELECT COUNT(rew.RAmount) FROM tbRework rew WHERE rew.RTo=main.RTo AND rew.RDate=main.RDate) as [Amount], main.RDate " _
        & "FROM (SELECT r.RTo , oTo.sapId as [To order], SUM(r.RAmount) as [Amount], r.RDate " _
        & "FROM tbRework r LEFT JOIN tbOrders oTo ON r.RTo=oTo.orderId " _
        & "WHERE oTo.sapId IN (" & zforList & ") " _
        & "GROUP BY oTo.sapId, r.RTo, r.RDate) main "
    
Set rs = New ADODB.Recordset

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If InStr(1, zforList, CStr(rs.Fields("To order").Value), vbTextCompare) > 0 Then
            'InStr(1, zfinList, CStr(rs.Fields("From order").Value), vbTextCompare) = 0 And InStr(1, zforList, CStr(rs.Fields("To order").Value), vbTextCompare) > 0
            'if "From order" IS NOT in zfinList and "To order" IS in zforList
            Set b = Me.findBlendByOrder(rs.Fields("To order").Value, "zfor")
            b.saveOrders rs.Fields("To order").Value, , , , , rs.Fields("Amount").Value
            b.rework = b.rework + rs.Fields("Amount").Value
            pRework = pRework + rs.Fields("Amount").Value
        End If
        rs.MoveNext
    Loop
End If

closeConnection

End Sub

Public Sub calculateReworkAtPacking()
Dim SQL As String
Dim zforList As String
Dim zfinList As String
Dim b As clsBlend
Dim i As Integer

On Error GoTo err_trap

zforList = Me.getOrderList("zfor")
zfinList = Me.getOrderList("zfin")

updateConnection

SQL = "SELECT bFrom.batchId as [From order], oTo.sapId as [To order], SUM(r.RAmount*u.unitWeight) as [Amount] " _
    & "FROM tbReworkWarehouse r LEFT JOIN tbOrders oTo ON r.RTo=oTo.orderId " _
    & "LEFT JOIN tbBatch bFrom ON r.RFrom=bFrom.batchId " _
    & "LEFT JOIN tbUom u ON u.zfinId=bFrom.zfinId " _
    & "WHERE oTo.sapId IN (" & zfinList & ") " _
    & "GROUP BY bFrom.batchId, oTo.sapId"
    
Set rs = New ADODB.Recordset

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        i = i + 1
        If InStr(1, zfinList, CStr(rs.Fields("To order").Value), vbTextCompare) > 0 And Not IsNull(rs.Fields("Amount").Value) Then
            Set b = Me.findBlendByOrder(rs.Fields("To order").Value, "zfin")
            'b.saveOrders rs.Fields("To order").Value, , , , , rs.Fields("Amount").Value
            b.reworkAtPacking = b.reworkAtPacking + rs.Fields("Amount").Value
            If b.IsBeans Then
                pReworkAtPackingBeans = pReworkAtPackingBeans + rs.Fields("Amount").Value
            Else
                pReworkAtPackingGround = pReworkAtPackingGround + rs.Fields("Amount").Value
            End If
        End If
        
        rs.MoveNext
    Loop
End If

exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Error in calculateReworkAtPacking of clsBlendKeeper. Description: " & Err.Description, vbOKOnly + vbCritical, "Błąd"
Resume exit_here

End Sub

Public Function GetRework() As Double
GetRework = pRework
End Function

Public Function GetReworkAtPacking(Optional t As Variant) As Double
If IsMissing(t) Then
    GetReworkAtPacking = pReworkAtPackingBeans + pReworkAtPackingGround
Else
    If t = "beans" Then
        GetReworkAtPacking = pReworkAtPackingBeans
    Else
        GetReworkAtPacking = pReworkAtPackingGround
    End If
End If
End Function

Public Function getOrderList(oType As String) 'zfin or zfor
Dim b As clsBlend
Dim res As String
Dim nOrders As String

res = ""

If oType = "zfor" Then
    For Each b In pBlends
        nOrders = b.getOrdersNumbersString
        If Len(nOrders) > 0 Then
            res = res & nOrders & ","
        End If
    Next b
    If Len(res) > 0 Then res = Left(res, Len(res) - 1)
ElseIf oType = "zfin" Then
    For Each b In pBlends
        nOrders = b.getZfinOrdersNumbersString
        If Len(nOrders) > 0 Then
            res = res & nOrders & ","
        End If
    Next b
    If Len(res) > 0 Then res = Left(res, Len(res) - 1)
End If

getOrderList = res

End Function

Public Function findBlendByOrder(oNumber As Double, oType As String) As clsBlend
Dim b As clsBlend

For Each b In pBlends
    If b.IsOrderInBlend(oNumber, oType) Then
        Set findBlendByOrder = b
        Exit For
    End If
Next b

End Function

Public Function trimOrdersToBlends(oNumbers As String, blends As Variant) As String
'it removes all orders given in oNumbers from oNumbers if their blend is not in blends

On Error GoTo err_trap

Dim nOrders As String

If Len(oNumbers) > 0 Then
    updateConnection

    SQL = "SELECT o.sapId,z.zfinIndex FROM tbOrders o LEFT JOIN tbZfin z ON o.zfinId=z.zfinId WHERE o.sapId IN (" & oNumbers & ")"
        
    Set rs = New ADODB.Recordset
    
    rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable
    
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            If IsInArray(rs.Fields("zfinIndex").Value, blends) Then
                nOrders = nOrders & "'" & rs.Fields("sapId").Value & "',"
            End If
            rs.MoveNext
        Loop
        If Len(nOrders) > 0 Then trimOrdersToBlends = Left(nOrders, Len(nOrders) - 1)
    Else
        trimOrdersToBlends = trimOrdersToBlends
    End If
Else
    trimOrdersToBlends = trimOrdersToBlends
End If

exit_here:
Exit Function

err_trap:
MsgBox "There has been errer in trimOrdersToBlends. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function



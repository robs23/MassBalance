VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public index As Long
Public Name As String
Private greenCoffee As Double
Private roastedCoffee As Double
Private groundCoffee As Double
Private totalPacked As Double
Private totalWarehoused As Double
Private pRoastGroundScrap As Double
Private pPackScrap As Double
Private ords() As Variant
Private bZfins As New Collection
Public IsBeans As Boolean
Public isGroundConsistent As Boolean
Public consistencyStr As String
Public consistent As Boolean 'used only when restoring from sheet
Public cost As Double
Public CostLotSize As Double
Public inScope As Boolean 'is this blend in the scope? If not, don't include it into overviews
Public bomVsReal As Double
Public gpValueLoss As Double
Public rework As Double
Public reworkAtPacking As Double
Public AvgRoastingLoss As Double
Public AvgGrindingLoss As Double
Public AvgPackingLoss As Double
Public AvgTotalLoss As Double
Public AvgGPLoss As Double
Public AvgRework As Double
Public AvgBomVsReal As Double

Public Property Get reworkValue() As Double 'money value of the coffee that had been added to this blend
reworkValue = (rework + reworkAtPacking) * (cost / CostLotSize)
End Property

Public Property Let green(v As Double)
greenCoffee = v
End Property

Public Property Let Roasted(v As Double)
roastedCoffee = v
End Property

Public Property Let Ground(v As Double)
groundCoffee = v
End Property

Public Property Let Packed(v As Double)
totalPacked = v
End Property

Public Property Let Stocked(v As Double)
totalWarehoused = v
End Property

Public Property Let roastGroundScrap(val As Double)
pRoastGroundScrap = val
End Property

Public Property Get roastGroundScrap() As Double
roastGroundScrap = pRoastGroundScrap
End Property

Public Property Let packScrap(val As Double)
pPackScrap = val
End Property

Public Property Get packScrap() As Double
packScrap = pPackScrap
End Property

Public Property Get totalScrap() As Double
totalScrap = pRoastGroundScrap + pPackScrap
End Property

Public Sub addGreen(oNumber As Double, g As Double, roaster As Integer)
If g > 0 Then
    greenCoffee = greenCoffee + g
    saveOrders oNumber, g, , , roaster
End If
End Sub

Public Property Get CostOfKg() As Double

On Error GoTo err_trap

CostOfKg = cost / CostLotSize

exit_here:
Exit Property

err_trap:
CostOfKg = 0
Resume exit_here

End Property

Public Property Get RoastedValue() As Double
'returns how much is worth this blend after roasting
On Error GoTo err_trap

RoastedValue = (cost / CostLotSize) * Me.getRoasted

exit_here:
Exit Property

err_trap:
RoastedValue = 0
Resume exit_here

End Property

Public Property Get PackedValue() As Double
'returns how much is worth this blend after roasting
On Error GoTo err_trap

PackedValue = (cost / CostLotSize) * Me.getPacked

exit_here:
Exit Property

err_trap:
PackedValue = 0
Resume exit_here

End Property

Public Property Get getGreen(Optional ord As Variant, Optional roaster As Variant) As Double
Dim i As Integer

If Not IsMissing(ord) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        If ords(0, i) = o Then
            If IsMissing(roaster) Then
                getGreen = ords(1, i) = 0 'green coffee
            Else
                If roaster = 3000 Then
                    getGreen = ords(4, i)
                Else
                    getGreen = ords(6, i)
                End If
            End If
            Exit For
        End If
    Next i
Else
    If IsMissing(roaster) Then
        getGreen = greenCoffee 'green coffee
    Else
        For i = LBound(ords, 2) To UBound(ords, 2)
             If roaster = 3000 Then
                 getGreen = getGreen + ords(4, i) 'green coffee of RN4000
             Else
                 getGreen = getGreen + ords(6, i) 'green coffee of RN4000
             End If
         Next i
    End If
End If
End Property

Public Property Get getPacked() As Double
getPacked = totalPacked
End Property

Public Property Get getStocked() As Double
getStocked = totalPacked 'totalWarehoused
End Property

Public Sub addRoasted(oNumber As Double, r As Double, roaster As Integer)
If r > 0 Then
    roastedCoffee = roastedCoffee + r
    saveOrders oNumber, , r, , roaster
End If
End Sub

Public Property Get getRoasted(Optional ord As Variant, Optional roaster As Variant) As Double
Dim i As Integer

If Not IsMissing(ord) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        If ords(0, i) = o Then
            If IsMissing(roaster) Then
                getRoasted = ords(2, i) 'roasted coffee
            Else
                If roaster = 3000 Then
                    getRoasted = ords(5, i) 'roasted coffee of RN4000
                Else
                    getRoasted = ords(7, i) 'roasted coffee of RN4000
                End If
            End If
            Exit For
        End If
    Next i
Else
    If IsMissing(roaster) Then
        getRoasted = roastedCoffee 'roasted coffee
    Else
        For i = LBound(ords, 2) To UBound(ords, 2)
             If roaster = 3000 Then
                 getRoasted = getRoasted + ords(5, i) 'roasted coffee of RN4000
             Else
                 getRoasted = getRoasted + ords(7, i) 'roasted coffee of RN4000
             End If
         Next i
    End If
End If
End Property

Public Property Get getGround(Optional ord As Variant) As Double
Dim i As Integer

If Not IsMissing(ord) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        If ords(0, i) = o Then
            If Me.IsBeans Then
                getGround = ords(2, i)
            Else
                getGround = ords(3, i) 'ground coffee
            End If
            Exit For
        End If
    Next i
Else
    If Me.IsBeans Then
        getGround = roastedCoffee
    Else
        getGround = groundCoffee
    End If
End If
End Property

Public Property Get numberOfOrders() As Integer
If isArrayEmpty(ords, 2) Then
    numberOfOrders = 0
Else
    numberOfOrders = UBound(ords, 2) + 1
End If
End Property

Public Property Get numberOfZfinOrders() As Integer
Dim bzfin As clsZfin
Dim o As Variant
Dim tot As Integer

For Each bzfin In bZfins
    o = bzfin.getOrders
    If Not isArrayEmpty(o, 2) Then
        tot = tot + UBound(o, 2) + 1
    End If
Next bzfin

numberOfZfinOrders = tot

End Property

Public Property Get numberOfZfins() As Integer
numberOfZfins = bZfins.Count
End Property

Public Sub saveOrders(o As Double, Optional g As Variant, Optional r As Variant, Optional Ground As Variant, Optional roaster As Variant, Optional rework As Variant)
Dim found As Boolean
Dim i As Integer
found = False

If Not isArrayEmpty(ords, 2) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        If ords(0, i) = o Then
            found = True
            If Not IsMissing(g) Then ords(1, i) = ords(1, i) + g 'green coffee
            If Not IsMissing(r) Then ords(2, i) = ords(2, i) + r 'roasted coffee
            If Not IsMissing(Ground) Then ords(3, i) = Ground 'ground coffee
            If Not IsMissing(roaster) Then
                If roaster = 3000 Then
                    If Not IsMissing(g) Then ords(4, i) = g 'roaster's RN3000 green coffee
                    If Not IsMissing(r) Then ords(5, i) = r 'roaster's RN3000 roasted coffee
                ElseIf roaster = 4000 Then
                    If Not IsMissing(g) Then ords(6, i) = g 'roaster's RN4000 green coffee
                    If Not IsMissing(r) Then ords(7, i) = r 'roaster's RN4000 roasted coffee
                End If
            End If
            If Not IsMissing(rework) Then
                ords(8, i) = rework
            End If
            Exit For
        End If
    Next i
End If

If found = False Then
    If isArrayEmpty(ords, 2) Then
        ReDim ords(8, 0) As Variant
        ords(0, 0) = o 'order number
        If Not IsMissing(g) Then ords(1, 0) = g 'green coffee
        If Not IsMissing(r) Then ords(2, 0) = r 'roasted coffee
        If Not IsMissing(Ground) Then ords(3, 0) = Ground 'ground coffee
        If Not IsMissing(roaster) Then
            If roaster = 3000 Then
                If Not IsMissing(g) Then ords(4, 0) = g 'roaster's RN3000 green coffee
                If Not IsMissing(r) Then ords(5, 0) = r 'roaster's RN3000 roasted coffee
            ElseIf roaster = 4000 Then
                If Not IsMissing(g) Then ords(6, 0) = g 'roaster's RN4000 green coffee
                If Not IsMissing(r) Then ords(7, 0) = r 'roaster's RN4000 roasted coffee
            End If
        End If
        If Not IsMissing(rework) Then
            ords(8, i) = rework
        End If
    Else
        ReDim Preserve ords(8, UBound(ords, 2) + 1) As Variant
        ords(0, UBound(ords, 2)) = o
        If Not IsMissing(g) Then ords(1, UBound(ords, 2)) = g 'green coffee
        If Not IsMissing(r) Then ords(2, UBound(ords, 2)) = r 'roasted coffee
        If Not IsMissing(Ground) Then ords(3, UBound(ords, 2)) = Ground 'ground coffee
        If Not IsMissing(roaster) Then
            If roaster = 3000 Then
                If Not IsMissing(g) Then ords(4, UBound(ords, 2)) = g 'roaster's RN3000 green coffee
                If Not IsMissing(r) Then ords(5, UBound(ords, 2)) = r 'roaster's RN3000 roasted coffee
            ElseIf roaster = 4000 Then
                If Not IsMissing(g) Then ords(6, UBound(ords, 2)) = g 'roaster's RN4000 green coffee
                If Not IsMissing(r) Then ords(7, UBound(ords, 2)) = r 'roaster's RN4000 roasted coffee
            End If
        End If
        If Not IsMissing(rework) Then
            ords(8, i) = rework
        End If
    End If
End If
End Sub

Public Function getOrders() As Variant
If isArrayEmpty(ords, 2) Then
    getOrders = "B/D"
Else
    getOrders = ords
End If

End Function

Private Function newZfin(zNumber As Long, Name As String) As clsZfin
Dim found As Boolean
Dim z As clsZfin

found = False

If bZfins.Count > 0 Then
    For Each z In bZfins
        If z.zfinIndex = zNumber Then
            found = True
            Set newZfin = z
            Exit For
        End If
    Next z
End If

If found = False Then
    Set newZfin = New clsZfin
    newZfin.zfinIndex = zNumber
    newZfin.zfinName = Name
    bZfins.Add newZfin, CStr(zNumber)
End If

End Function

Public Sub setPacked()
Dim n As Integer
Dim i As Integer
Dim z As clsZfin
Dim SQL As String
Dim oNumbers As String
Dim rs As ADODB.Recordset
Dim nzfin As clsZfin

On Error GoTo err_trap

n = bZfins.Count 'clean up
Do While bZfins.Count > 0
    bZfins.Remove n
    n = n - 1
Loop

If Not isArrayEmpty(ords, 2) Then
    updateConnection
    For i = LBound(ords, 2) To UBound(ords, 2)
        'for every zfor order find zfin orders and their data
        oNumbers = oNumbers & ords(0, i) & ","
    Next i
    If Len(oNumbers) > 0 Then
        oNumbers = Left(oNumbers, Len(oNumbers) - 1)
        setGround oNumbers
        'sql = "SELECT DISTINCT o1.sapId as id, z.zfinIndex as zfin, z.zfinName as name,o1.executedSap*u.unitWeight as kg, b.batchNumber as batch, SUM(q.batchSize*u.unitWeight) as pw " _
                & "FROM tbOrders o LEFT JOIN tbOrderDep od ON od.zforOrder=o.orderId LEFT JOIN tbOrders o1 ON o1.orderId = od.zfinOrder LEFT JOIN tbZfin z ON z.zfinId = o1.zfinId LEFT JOIN tbUom u ON u.zfinId = z.zfinId LEFT JOIN tbBatch b ON b.batchId=o1.batchId LEFT JOIN tbQdocData q ON q.batchId= o1.batchId LEFT JOIN tbQdocReconciliation qr ON qr.qReconciliationId=q.qReconciliationId " _
                & "WHERE o.sapId IN (" & oNumbers & ") AND qr.qType = 'WHD_PW' " _
                & "GROUP BY o1.sapId, z.zfinIndex, z.zfinName,o1.executedSap*u.unitWeight, b.batchNumber " _
                & "ORDER BY zfin"
        If pSource = "SAP" Then
            SQL = "SELECT DISTINCT o1.sapId as id, o1.isCancelled, z.zfinIndex as zfin, z.zfinName as name,o1.executedSap*u.unitWeight as kg, b.batchNumber as batch, (SELECT SUM(qd.batchSize*u1.unitWeight) FROM tbQdocData qd INNER JOIN tbQdocReconciliation qr on qr.qReconciliationId=qd.qReconciliationId INNER JOIN tbUom u1 ON u1.zfinId=b.zfinId WHERE qr.qType='WHD_PW' AND qd.batchId=b.batchId) as PW " _
                & "FROM tbOrders o RIGHT JOIN tbOrderDep od ON od.zforOrder=o.orderId LEFT JOIN tbOrders o1 ON o1.orderId = od.zfinOrder LEFT JOIN tbZfin z ON z.zfinId = o1.zfinId LEFT JOIN tbUom u ON u.zfinId = z.zfinId LEFT JOIN tbBatch b ON b.batchId=o1.batchId " _
                & "WHERE o.sapId IN (" & oNumbers & ") AND (od.isRemoved = 0 OR od.isRemoved IS NULL) GROUP BY o1.sapId, z.zfinIndex, o1.isCancelled, z.zfinName,o1.executedSap*u.unitWeight, b.batchNumber,b.batchId,b.zfinId;"
        ElseIf pSource = "MES" Then
            SQL = "SELECT DISTINCT o1.sapId as id, o1.isCancelled, z.zfinIndex as zfin, z.zfinName as name,o1.executedMes*u.unitWeight as kg, b.batchNumber as batch, (SELECT SUM(qd.batchSize*u1.unitWeight) FROM tbQdocData qd INNER JOIN tbQdocReconciliation qr on qr.qReconciliationId=qd.qReconciliationId INNER JOIN tbUom u1 ON u1.zfinId=b.zfinId WHERE qr.qType='WHD_PW' AND qd.batchId=b.batchId) as PW " _
                & "FROM tbOrders o RIGHT JOIN tbOrderDep od ON od.zforOrder=o.orderId LEFT JOIN tbOrders o1 ON o1.orderId = od.zfinOrder LEFT JOIN tbZfin z ON z.zfinId = o1.zfinId LEFT JOIN tbUom u ON u.zfinId = z.zfinId LEFT JOIN tbBatch b ON b.batchId=o1.batchId " _
                & "WHERE o.sapId IN (" & oNumbers & ") AND (od.isRemoved = 0 OR od.isRemoved IS NULL) GROUP BY o1.sapId, z.zfinIndex, o1.isCancelled, z.zfinName,o1.executedMes*u.unitWeight, b.batchNumber,b.batchId,b.zfinId;"
        End If
        Set rs = CreateObject("adodb.recordset")
        rs.Open SQL, adoConn
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                Set nzfin = newZfin(rs.Fields("zfin").Value, rs.Fields("name").Value)
                With nzfin
                    If IsNull(rs.Fields("batch")) Then
                        missingBatches = missingBatches & rs.Fields("id") & ","
                    End If
                    .addPacked rs.Fields("id").Value, rs.Fields("batch").Value, rs.Fields("kg").Value, rs.Fields("pw").Value, rs.Fields("isCancelled").Value
                    If Not IsNull(rs.Fields("kg").Value) Then totalPacked = totalPacked + rs.Fields("kg").Value
                    If Not IsNull(rs.Fields("pw").Value) Then totalWarehoused = totalWarehoused + rs.Fields("pw").Value
                End With
                rs.MoveNext
            Loop
        End If
        rs.Close
    End If
End If

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Sub

err_trap:
MsgBox "Error in setPacked of clsBlend. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub setGround(operStr As String)
Dim rs As ADODB.Recordset
Dim colName As String
Dim SQL As String
Dim bool As Boolean

Set rs = CreateObject("adodb.recordset")

SQL = "SELECT o.sapId, o.executedMes, o.executedSap, o.executedMesGround, zp.[beans?] as isBean " _
    & "FROM tbOrders o RIGHT JOIN tbZfinProperties zp ON zp.zfinId=o.zfinId " _
    & "WHERE o.sapId IN (" & operStr & ")"
    
rs.Open SQL, adoConn
If Not rs.EOF Then
    rs.MoveFirst
    If rs.Fields("isBean") <> 0 Then
        Me.IsBeans = True
    Else
        Me.IsBeans = False
    End If
    If gSource = "MES" Then
        colName = "executedMesGround"
    Else
        colName = "executedSap"
    End If
    bool = True
    Do Until rs.EOF
        If Not IsNull(rs.Fields(colName)) Then
            groundCoffee = groundCoffee + rs.Fields(colName)
            saveOrders rs.Fields("sapId"), , , rs.Fields(colName)
        Else
            bool = False
        End If
        rs.MoveNext
    Loop
Else
    bool = False
End If

isGroundConsistent = bool

rs.Close
Set rs = Nothing

End Sub

Public Function getZfins() As Collection
Set getZfins = bZfins
End Function

Public Function isConsistent() As Boolean
Dim nzfin As clsZfin
Dim bool As Boolean
Dim SQL As String
Dim rs As ADODB.Recordset
Dim i As Integer
Dim operStr As String

On Error GoTo err_trap

updateConnection

bool = False

For Each nzfin In bZfins
    If nzfin.allOrdersConsistent Then
        bool = True
    Else
        bool = False
        Exit For
    End If
Next nzfin

'If bool Then
'    bool = True
'    If isArrayEmpty(ords, 2) Then
'        bool = False
'        consistencyStr = "There are no roasting orders for ZFOR " & index & "."
'    Else
'        For i = LBound(ords, 2) To UBound(ords, 2)
'            operStr = operStr & ords(0, i) & ","
'        Next i
'        If Len(operStr) > 0 Then operStr = Left(operStr, Len(operStr) - 1)
'            'additional check - do we have packing orders for all roasting orders? At least 1 for each should be present to evaluate TRUE
'            sql = "SELECT o.sapId as sapId, COUNT(od.zfinOrder) as Counter " _
'                & "FROM tbOrders o LEFT JOIN tbOrderDep od ON od.zforOrder=o.orderId " _
'                & "WHERE o.sapId IN (" & operStr & ") GROUP BY o.sapId"
'
'            Set rs = CreateObject("adodb.recordset")
'            rs.Open sql, adoConn
'            If rs.EOF Then
'                bool = False
'            Else
'                rs.MoveFirst
'                Do Until rs.EOF
'                    If rs.Fields("counter") = 0 Then
'                        bool = False
'                        If Len(consistencyStr) = 0 Then
'                            consistencyStr = "Can't find packing order(s) for roasting order(s) " & rs.Fields("sapId")
'                        Else
'                            consistencyStr = consistencyStr & ", " & rs.Fields("sapId")
'                        End If
'                    End If
'                    rs.MoveNext
'                Loop
'            End If
'        End If
'End If

isConsistent = bool

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Function

err_trap:
MsgBox "Error in isConsistent of clsBlend. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Function getLoss(processType As String, valueType As String, Optional roaster As Variant, Optional vsAvg As Variant) As Double
    Dim difference As Double

On Error GoTo err_trap

If IsMissing(vsAvg) Then
    vsAvg = False
End If

Select Case processType
    Case Is = "r"
        If vsAvg Then
            difference = Me.getRoasted - Me.getGreen
            getLoss = difference / Me.getGreen
            getLoss = getLoss * 100 - AvgRoastingLoss
        Else
            If IsMissing(roaster) Then
                difference = Me.getRoasted - Me.getGreen
                If valueType = "kg" Then
                    getLoss = difference
                Else
                    getLoss = difference / Me.getGreen
                End If
            Else
                difference = Me.getRoasted(, roaster) - Me.getGreen(, roaster)
                If valueType = "kg" Then
                    getLoss = difference
                Else
                    If difference = Me.getRoasted(, roaster) Then
                        getLoss = 0
                    Else
                        getLoss = difference / Me.getGreen(, roaster)
                    End If
                End If
            End If
        End If
        
    Case Is = "g"
        If vsAvg Then
            difference = Me.getGround - (Me.getRoasted + Me.rework)
            getLoss = difference / (Me.getRoasted + Me.rework)
            getLoss = getLoss * 100 - AvgGrindingLoss
        Else
            difference = Me.getGround - (Me.getRoasted + Me.rework)
            If valueType = "kg" Then
                getLoss = difference
            Else
                getLoss = difference / (Me.getRoasted + Me.rework)
            End If
        End If
        
    Case Is = "p"
        If vsAvg Then
            difference = Me.getPacked - (Me.getGround + Me.reworkAtPacking)
            getLoss = difference / (Me.getGround + Me.reworkAtPacking)
            getLoss = getLoss * 100 - AvgPackingLoss
        Else
            difference = Me.getPacked - (Me.getGround + Me.reworkAtPacking)
            If valueType = "kg" Then
                getLoss = difference
            Else
                getLoss = difference / (Me.getGround + Me.reworkAtPacking)
            End If
        End If
    Case Is = "t"
        If vsAvg Then
            difference = Me.getStocked - (Me.getGreen + Me.rework + Me.reworkAtPacking)
            getLoss = difference / (Me.getGreen + Me.rework + Me.reworkAtPacking)
            getLoss = getLoss * 100 - AvgTotalLoss
        Else
            difference = Me.getStocked - (Me.getGreen + Me.rework + Me.reworkAtPacking)
            If valueType = "kg" Then
                getLoss = difference
            Else
                getLoss = difference / (Me.getGreen + Me.rework + Me.reworkAtPacking)
            End If
        End If
    Case Is = "g+p"
        If vsAvg Then
            difference = Me.getPacked - (Me.getRoasted + Me.rework + Me.reworkAtPacking)
            getLoss = difference / (Me.getRoasted + Me.rework + Me.reworkAtPacking)
            getLoss = getLoss * 100 - AvgGPLoss
        Else
            difference = Me.getPacked - (Me.getRoasted + Me.rework + Me.reworkAtPacking)
            If valueType = "kg" Then
                getLoss = difference
            ElseIf valueType = "$" Then
                getLoss = difference * Me.CostOfKg
            Else
                getLoss = difference / (Me.getRoasted + Me.rework + Me.reworkAtPacking)
            End If
        End If
    Case Is = "r+g"
        difference = Me.getGround - Me.rework - (Me.getGreen + Me.rework)
        If valueType = "kg" Then
            getLoss = difference
        Else
            getLoss = difference / (Me.getGreen + Me.rework)
        End If
End Select

exit_here:
Exit Function

err_trap:
If Err.Number = 6 Or Err.Number = 11 Then
    getLoss = 0
Else
    MsgBox "Error in ""getLoss"" of clsBlend. Error number: " & Err.Number & ", " & Err.Description
End If
Resume exit_here

End Function

Public Function getRealVsBom(processType As String, valueType As String) As Double
Dim difference As Double

Select Case processType
    Case Is = "r+g"
        If valueType = "%" Then
            getRealVsBom = Me.getLoss(processType, "%") - (-1 * Me.roastGroundScrap / 100)
        End If
    Case Is = "p"
        If valueType = "%" Then
            getRealVsBom = Me.getLoss(processType, "%") - (-1 * Me.packScrap / 100)
        End If
    Case Is = "t"
        If valueType = "%" Then
            getRealVsBom = Me.getLoss(processType, "%") - (-1 * Me.totalScrap / 100)
        End If
End Select
End Function

Public Function getOrdersNumbersString() As String
' gives you string with order numbers within the blend e.g. "12000012, 12000013, 12000014"
Dim i As Integer
Dim res As String

res = ""

If Not isArrayEmpty(ords, 2) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        res = res & ords(0, i) & ","
    Next i
    If Len(res) > 0 Then res = Left(res, Len(res) - 1)
End If

getOrdersNumbersString = res

End Function

Public Function getZfinOrdersNumbersString() As String
' gives you string with order numbers within the zfins of this blend e.g. "12000012, 12000013, 12000014"
Dim i As Integer
Dim res As String
Dim z As clsZfin
Dim zords As Variant

res = ""

For Each z In bZfins
    zords = z.getOrders
    If IsArray(zords) Then
        For i = LBound(zords, 2) To UBound(zords, 2)
        res = res & zords(0, i) & ","
        Next i
    End If
Next z
If Len(res) > 0 Then res = Left(res, Len(res) - 1)

getZfinOrdersNumbersString = res

End Function

Public Function IsOrderInBlend(oNumber As Double, oType As String) As Boolean
Dim res As Boolean
Dim i As Integer
Dim zfinOrders As String

res = False

If oType = "zfor" Then
    If Not isArrayEmpty(ords, 2) Then
        For i = LBound(ords, 2) To UBound(ords, 2)
            If ords(0, i) = oNumber Then
                res = True
                Exit For
            End If
        Next i
    End If
ElseIf oType = "zfin" Then
    zfinOrders = getZfinOrdersNumbersString
    If Len(zfinOrders) > 0 Then
        If InStr(1, zfinOrders, CStr(oNumber), vbTextCompare) > 0 Then
            res = True
        End If
    End If
End If

IsOrderInBlend = res

End Function



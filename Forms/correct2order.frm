VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} correct2order 
   Caption         =   "Correct to order"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   OleObjectBlob   =   "correct2order.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "correct2order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private order As Long
Private batch As Long
Private challenger As Long

Private Sub UserForm_Activate()
Dim sht As Worksheet
Dim rs As ADODB.Recordset
Dim SQL As String
Dim o As clsOrder

Set sht = ThisWorkbook.Sheets("Operations sequence")
order = 0
batch = 0

updateConnection

challenger = ActiveCell
SQL = "SELECT * FROM tbOrders o WHERE o.sapId = " & challenger

Set rs = New ADODB.Recordset
rs.Open SQL, adoConn

If Not rs.EOF Then
    Set o = New clsOrder
    o.orderId = rs.Fields("orderId").Value
    o.sapId = rs.Fields("sapId").Value
    o.orderType = rs.Fields("type").Value
    If o.orderType = "r" Then
        'roasting/grinding order
        o.mesRoasted = rs.Fields("executedMes").Value
        o.mesGround = rs.Fields("executedMesGround").Value
        o.sapGround = rs.Fields("executedSap").Value
        connectScada
        rs.Close
        SQL = "SELECT rD.OrderNumber as theOrder, rd.MaterialNumber as zfor, rd.NAZWARECEPT as name, sum(rd.SUMA_ZIELONEJ) as sumaZielonej,sum(rd.ILOSC_PALONA) As sumaPalonej " _
            & "FROM (select DISTINCT z.NUMERPIECA, z.SUMA_ZIELONEJ, z.ILOSC_PALONA, z.DTZAPIS, zl.OrderNumber, zl.MaterialNumber, zl.NAZWARECEPT from ZLECENIA_PALONA z Join ZLECENIAWARTOSCI w ON (z.IDZLECENIE = w.IDZLECENIE) JOIN ZLECENIA zl on (w.IDZLECENIE = zl.IDZLECENIE)) as rD " _
            & "WHERE rd.OrderNumber IN (" & o.sapId & ") " _
            & "GROUP BY rd.OrderNumber,rd.MaterialNumber,rd.NAZWARECEPT " _
            & "ORDER BY zfor"
        rs.Open SQL, conn
        If Not rs.EOF Then
            o.scadaGreen = Round(rs.Fields("sumaZielonej").Value, 1)
            o.scadaRoasted = Round(rs.Fields("sumaPalonej").Value, 1)
        End If
        rs.Close
        Set rs = Nothing
    Else
        'packing order
    End If
    
    Me.txtRoastedMes = o.mesRoasted
    Me.txtGroundMes = o.mesGround
    Me.txtGroundSap = o.sapGround
    Me.txtGreenScada = o.scadaGreen
    Me.txtRoastedScada = o.scadaRoasted
Else
    rs.Close
    SQL = "SELECT * FROM tbBatch b WHERE b.batchNumber = " & challenger
    rs.Open SQL, adoConn
    If Not rs.EOF Then
    
    Else
        MsgBox "There's no order/batch " & challenger & ". Can't proceed..", vbCritical + vbOKOnly, "The number doesn't exist"
    End If
    rs.Close
    Set rs = Nothing
End If

End Sub


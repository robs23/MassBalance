VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rStart As Date
Public rEnd As Date
Public greenCoffee As Double
Public mksLoss As Double
Public mksReceipt As Double
Public mksPurge As Double
Public mksDiff As Double
Public rBeansIn As Double
Public rGroundIn As Double
Public rBeansOut As Double
Public rGroundOut As Double
Public pGroundIn As Double
Public pBeansOut As Double
Public pGroundOut As Double
Public pwBeans As Double
Public pwGround As Double
Public rework As Double
Public reworkAtPacking As Double
Public reworkAtPackingBeans As Double
Public reworkAtPackingGround As Double
Public countRework As Boolean
Public groundStr As String
Public blends As New Collection

Public Sub deployResults()
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("BM")
If sht.ProtectContents Then sht.Unprotect

If Len(period) = 0 Then
    sht.Range("P2") = "Bilans mas za okres od " & rStart & " do " & rEnd
    sht.Range("P3") = ""
Else
    If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
        sht.Range("P2") = "Bilans mas za tydzień " & period
    Else
        sht.Range("P2") = "Bilans mas za miesiąc " & period
    End If
    sht.Range("P3") = rStart & " - " & rEnd
End If

sht.Range("B8") = greenCoffee
sht.Range("D12") = mksReceipt
sht.Range("D14") = mksPurge
sht.Range("D16") = mksDiff
sht.Range("G12") = rBeansIn
sht.Range("G14") = rGroundIn
sht.Range("M12") = rBeansOut
sht.Range("M14") = rGroundOut
sht.Range("S12") = rBeansOut
sht.Range("S14") = pGroundIn
sht.Range("Y12") = pBeansOut
sht.Range("Y14") = pGroundOut
sht.Range("AE12") = pwBeans
sht.Range("AE14") = pwGround
sht.Range("P10") = rework
sht.Range("T10") = reworkAtPacking
sht.Range("V12") = (sht.Range("S12") + reworkAtPackingBeans - sht.Range("Y12")) * -1
sht.Range("V14") = (sht.Range("S14") + reworkAtPackingGround - sht.Range("Y14")) * -1
sht.Range("P18") = 1 - ((sht.Range("AE14") - sht.Range("P10") - reworkAtPackingGround) / sht.Range("G14"))
sht.Range("S18") = 1 - ((sht.Range("AE12") - reworkAtPackingBeans) / sht.Range("G12"))
sht.Protect
End Sub

Public Sub addRoast(amount As Double, IsBeans As Boolean)
If IsBeans Then
    'it's beans
    rBeansOut = rBeansOut + amount
Else
    rGroundOut = rGroundOut + amount
End If
End Sub

Public Sub addGreen(amount As Double, IsBeans As Boolean)
If IsBeans Then
    'it's beans
    rBeansIn = rBeansIn + amount
Else
    rGroundIn = rGroundIn + amount
End If
End Sub

Public Sub addGround(amount As Double, IsBeans As Boolean)
If IsBeans Then
    'it's beans
    pBeansIn = pBeansIn + amount
Else
    pGroundIn = pGroundIn + amount
End If
End Sub

Public Sub addRework(amount As Double)
rework = rework + amount
End Sub

Public Sub addReworkAtPacking(amount As Double, Optional t As Variant)
If IsMissing(t) Then
    reworkAtPacking = reworkAtPacking + amount
Else
    If t = "beans" Then
        reworkAtPackingBeans = reworkAtPackingBeans + amount
    Else
        reworkAtPackingGround = reworkAtPackingGround + amount
    End If
End If
End Sub

Public Sub addPacked(amount As Double, Optional IsBeans As Boolean)
If IsBeans Then
    'it's beans
    pBeansOut = pBeansOut + amount
Else
    pGroundOut = pGroundOut + amount
End If
End Sub

Public Sub addWarehoused(amount As Double, IsBeans As Boolean)
If IsBeans Then
    'it's beans
    pwBeans = pwBeans + amount
Else
    pwGround = pwGround + amount
End If
End Sub

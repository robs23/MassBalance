VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsZfin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public zfinId As Long
Public zfinIndex As Long
Public zfinName As String
Public IsBeans As Boolean
Private Packed As Double
Private pewu As Double
Private ords() As Variant

Public Sub addPacked(oNumber As Double, Optional batchNumber As Variant, Optional p As Variant, Optional pw As Variant, Optional isCancelled As Variant)
Dim i As Integer
Dim found As Boolean

found = False

If Not isArrayEmpty(ords, 2) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        If ords(0, i) = oNumber Then
            found = True
            If Not IsMissing(p) And Not IsNull(p) Then ords(2, i) = p 'packed
            If Not IsMissing(batchNumber) And Not IsNull(batchNumber) Then ords(1, i) = batchNumber 'number of batch
            If Not IsMissing(pw) And Not IsNull(pw) Then ords(3, i) = pw
            If Not IsMissing(isCancelled) And Not IsNull(isCancelled) Then ords(4, i) = isCancelled Else ords(4, 0) = False
            Exit For
        End If
    Next i
End If

If found = False Then
    If isArrayEmpty(ords, 2) Then
        ReDim ords(4, 0) As Variant
        ords(0, 0) = oNumber 'order number
        If Not IsMissing(batchNumber) And Not IsNull(batchNumber) Then ords(1, 0) = batchNumber 'batch
        If Not IsMissing(p) And Not IsNull(p) Then ords(2, 0) = p 'packed
        If Not IsMissing(pw) And Not IsNull(pw) Then ords(3, 0) = pw 'pw
        If Not IsMissing(isCancelled) And Not IsNull(isCancelled) Then ords(4, 0) = isCancelled Else ords(4, 0) = False
    Else
        ReDim Preserve ords(4, UBound(ords, 2) + 1) As Variant
        ords(0, UBound(ords, 2)) = oNumber
        If Not IsMissing(batchNumber) And Not IsNull(batchNumber) Then ords(1, UBound(ords, 2)) = batchNumber 'batch
        If Not IsMissing(p) And Not IsNull(p) Then ords(2, UBound(ords, 2)) = p 'packed
        If Not IsMissing(pw) And Not IsNull(pw) Then ords(3, UBound(ords, 2)) = pw 'pw
        If Not IsMissing(isCancelled) And Not IsNull(isCancelled) Then ords(4, UBound(ords, 2)) = isCancelled Else ords(4, UBound(ords, 2)) = False
    End If
End If
If Not IsMissing(p) And Not IsNull(p) Then Packed = Packed + p
If Not IsMissing(pw) And Not IsNull(pw) Then pewu = pewu + pw
End Sub

Public Property Get getPacked(Optional ord As Variant) As Double
Dim i As Integer

If Not IsMissing(ord) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        If ords(0, i) = o Then
            getPacked = ords(2, i) = 0 'packed coffee
            Exit For
        End If
    Next i
Else
    getPacked = Packed
End If
End Property

Public Property Get getStocked(Optional ord As Variant) As Double
Dim i As Integer

If Not IsMissing(ord) Then
    For i = LBound(ords, 2) To UBound(ords, 2)
        If ords(0, i) = o Then
            getStocked = ords(2, i) = 0 'packed coffee
            Exit For
        End If
    Next i
Else
    getStocked = Packed 'pewu
End If
End Property

Public Property Get numberOfOrders() As Integer
If isArrayEmpty(ords, 2) Then
    numberOfOrders = 0
Else
    numberOfOrders = UBound(ords, 2) + 1
End If
End Property

'Private Sub saveOrders(o As Double, Optional b As Variant, Optional p As Variant, Optional pw As Variant)
'Dim found As Boolean
'Dim i As Integer
'found = False
'
''If Not isArrayEmpty(ords, 2) Then
''    For i = LBound(ords, 2) To UBound(ords, 2)
''        If ords(0, i) = o Then
''            found = True
''            If Not IsMissing(p) Then ords(2, i) = p 'packed
''            Exit For
''        End If
''    Next i
''End If
'
'If found = False Then
'    If isArrayEmpty(ords, 2) Then
'        ReDim ords(3, 0) As Variant
'        ords(0, 0) = o 'order number
'        If Not IsMissing(b) Then ords(1, 0) = b 'batch
'        If Not IsMissing(p) Then ords(2, 0) = p 'packed
'        If Not IsMissing(pw) Then ords(3, 0) = pw 'pw
'    Else
'        ReDim Preserve ords(3, UBound(ords, 2) + 1) As Variant
'        ords(0, UBound(ords, 2)) = o
'        If Not IsMissing(b) Then ords(1, UBound(ords, 2)) = b 'batch
'        If Not IsMissing(p) Then ords(2, UBound(ords, 2)) = p 'packed
'        If Not IsMissing(pw) Then ords(3, UBound(ords, 2)) = pw 'pw
'    End If
'End If
'End Sub

Public Function getOrders() As Variant
If isArrayEmpty(ords, 2) Then
    getOrders = "B/D"
Else
    getOrders = ords
End If

End Function

Public Function allOrdersConsistent() As Boolean
Dim bool As Boolean
Dim i As Integer
bool = False

If Not isArrayEmpty(ords, 2) Then
    bool = True
    'All orders are valid
'    For i = LBound(ords, 2) To UBound(ords, 2)
'        If ords(4, i) = False Then
'            If (ords(2, i) = 0 Or ords(3, i) = 0) And ords(4, i) = False Then
'                bool = False
'                Exit For
'
'            ElseIf Abs(ords(2, i) - ords(3, i)) / ords(3, i) > inProgress / 100 Then
'                Debug.Print Abs(ords(2, i) - ords(3, i)) / ords(3, i)
'                bool = False
'                Exit For
'            End If
'        End If
'    Next i
End If

allOrdersConsistent = bool

End Function



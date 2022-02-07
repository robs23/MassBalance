VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsZlecenie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pIndex As Long
Private pName As String
Private pZlecenie As Long
Private pAmount As Double
Private pType As String 'r (roasting), g (grinding) or p (packing)
Private pMachine As Integer


Public Property Let index(ind As Long)
pIndex = ind
End Property

Public Property Let Name(n As String)
pName = n
End Property

Public Property Let order(o As Long)
pZlecenie = o
End Property

Public Property Let amount(val As Double)
pAmount = val
End Property

Public Property Let processType(t As String)
pType = t
End Property

Public Property Let Machine(m As Variant)
Dim i As Integer

For i = 1 To Len(m)
    'look for first non-zero digit in string starting from right
    If Not IsNumeric(Right(m, i)) Then
        pMachine = CInt(Right(m, i - 1))
    End If
Next i
End Property

Public Property Get index() As Long
index = pIndex
End Property

Public Property Get Name() As String
Name = pName
End Property

Public Property Get order() As Long
order = pZlecenie
End Property

Public Property Get amount() As Double
amount = pAmount
End Property

Public Property Get processType() As String
processType = pType
End Property

Public Property Get Machine() As Variant
Machine = pMachine
End Property


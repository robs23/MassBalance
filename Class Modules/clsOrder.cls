VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pOrderId As Long
Private pSapId As Long
Private pIsPresent As Boolean
Public scadaGreen As Variant
Public scadaRoasted As Variant
Public mesRoasted As Variant
Public mesGround As Variant
Public mesPacked As Variant
Public mesWarehoused As Variant
Public sapGround As Variant
Public sapPacked As Variant
Public orderType As String

Public Property Let orderId(val As Long)
pOrderId = val
End Property

Public Property Let sapId(val As Long)
pSapId = val
End Property

Public Property Get orderId() As Long
orderId = pOrderId
End Property

Public Property Get sapId() As Long
sapId = pSapId
End Property

Public Property Let isPresent(val As Boolean)
pIsPresent = val
End Property

Public Property Get isPresent() As Boolean
isPresent = pIsPresent
End Property



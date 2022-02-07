VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFunctionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ChosenFunction As clsFunction
Public Functions As New Collection
Public ReturnedOrders As New Collection

Private Sub Class_Initialize()
    InitializeFunctions
End Sub

Private Sub InitializeFunctions()
    Dim cFunction As clsFunction
    
    Set cFunction = New clsFunction
    cFunction.Id = 0
    cFunction.Name = "Numer sesji wg zleceń"
    cFunction.SQL = "USE npd " _
                & "SELECT o.sapId, op.SessionNumber " _
                & "FROM tbOrders o LEFT JOIN tbOperations op ON op.orderId=o.orderId " _
                & "WHERE o.sapId IN (params)"
    cFunction.Hint = "Wpisz numery zleceń oddzielone przecinkiem np. 1234567, 8901234:"
    Me.Functions.Add cFunction
    
    Set cFunction = New clsFunction
    cFunction.Id = 1
    cFunction.Name = "ZFOR -> ZFIN"
    cFunction.SQL = "Use npd " _
                & "SELECT DISTINCT zfin.sapId, op.SessionNumber, od.isRemoved " _
                & "FROM tbOrders zfor LEFT JOIN tbOrderDep od ON od.zforOrder=zfor.orderId " _
                & "LEFT JOIN tbOrders zfin ON zfin.orderId=od.zfinOrder " _
                & "LEFT JOIN tbOperations op ON op.orderId=zfin.orderId " _
                & "WHERE zfor.sapId IN (params)"
    cFunction.Hint = "Wpisz numery zleceń oddzielone przecinkiem np. 1234567, 8901234:"
    Me.Functions.Add cFunction
    
    Set cFunction = New clsFunction
    cFunction.Id = 2
    cFunction.Name = "ZFIN -> ZFOR"
    cFunction.SQL = "Use npd " _
                & "SELECT DISTINCT zfor.sapId, op.SessionNumber, od.isRemoved " _
                & "FROM tbOrders zfin LEFT JOIN tbOrderDep od ON od.zfinOrder=zfin.orderId " _
                & "LEFT JOIN tbOrders zfor ON zfor.orderId=od.zforOrder " _
                & "LEFT JOIN tbOperations op ON op.orderId=zfor.orderId " _
                & "WHERE zfin.sapId IN (params)"
    cFunction.Hint = "Wpisz numery zleceń oddzielone przecinkiem np. 1234567, 8901234:"
    Me.Functions.Add cFunction
    
    Set cFunction = New clsFunction
    cFunction.Id = 3
    cFunction.Name = "Numery zleceń wg ZFORa i sesji"
    cFunction.SQL = "Use npd " _
                & "SELECT DISTINCT oZfor.sapId, op.SessionNumber " _
                & "FROM tbZfin z LEFT JOIN tbOperations op ON op.zfinId=z.zfinId " _
                & "LEFT JOIN tbOrders oZfor ON oZfor.orderId=op.orderId " _
                & "WHERE op.SessionNumber=(SELECT TOP(1) value FROM dbo.SplitString('params',',') ORDER BY value DESC) AND z.zfinIndex=(SELECT TOP(1) value FROM dbo.SplitString('params',','))"
    cFunction.Hint = "Wpisz numer ZFORa i numer sesji odzielone przecinkiem np. 34005471, 612:"
    Me.Functions.Add cFunction
    
    
End Sub

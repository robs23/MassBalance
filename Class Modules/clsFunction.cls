VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Id As Integer
Public Name As String
Public SQL As String
Public params As String
Public Output As String
Public Hint As String

Public Function Execute(params As String) As Collection
Dim o As clsOrder
Dim i As Integer
Dim SQL As String
Dim rs As ADODB.Recordset
Dim orders As New Collection

On Error GoTo err_trap

updateConnection

Me.params = params

SQL = Replace(Me.SQL, "params", Me.params)
Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    
    Do Until rs.EOF
        Set o = New clsOrder
        o.sapId = rs.Fields("sapId")
        If inCollection(o.sapId, orders) = False Then
            orders.Add o, CStr(o.sapId)
        End If
        For i = 0 To rs.Fields.Count - 1
            Output = Output & vbTab & rs.Fields(i)
        Next i
        Output = Output & vbCr
        rs.MoveNext
    Loop
    
Else
    MsgBox "Żaden rekord nie spełnia określonych warunków. Warto sprawdzić dane wejściowe", vbInformation + vbOKOnly, "Brak danych"
End If

exit_here:
Set Execute = orders
Exit Function

err_trap:
MsgBox "Błąd: " + Err.Description, vbCritical + vbOKOnly, "Wystąpił błąd"
Resume exit_here

End Function



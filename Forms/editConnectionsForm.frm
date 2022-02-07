VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} editConnectionsForm 
   Caption         =   "Modify ZFIN - ZFOR orders' connections"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   OleObjectBlob   =   "editConnectionsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "editConnectionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private t As String
Private oId As Long
Private orders As New Collection

Private Sub btnSave_Click()
If verify Then
    editConnections
End If
End Sub

Private Sub txtIn_AfterUpdate()
Dim oIn As Long

If Not IsNull(Me.txtIn) Then
    Me.txtOut = ""
    If IsNumeric(Me.txtIn) Then
        oIn = Me.txtIn
        bringConnections oIn
    Else
        MsgBox "Please input only one order number into the first field!", vbExclamation + vbOKOnly, "Not numeric"
    End If
End If
End Sub

Private Sub bringConnections(oIn As Long)
'brings connected order numbers
'and sets lblIn type (ZFOR or ZFIN)
Dim rs As ADODB.Recordset
Dim SQL As String

updateConnection
SQL = "SELECT orderId, type FROM tbOrders WHERE sapId=" & oIn

Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic

If rs.EOF Then
    MsgBox "Order number " & oIn & " doesn't exist!", vbCritical + vbOKOnly, "Order doesn't exist"
Else
    Me.btnSave.Enabled = True
    Me.txtOut.Locked = False
    rs.MoveFirst
    If rs.Fields("type") = "p" Then
        Me.lblIn.Caption = "ZFIN"
        Me.lblOut.Caption = "ZFOR"
    Else
        Me.lblIn.Caption = "ZFOR"
        Me.lblOut.Caption = "ZFIN"
    End If
    t = rs.Fields("type")
    oId = rs.Fields("orderId")
    Me.lblIn.Visible = True
    Me.lblOut.Visible = True
    If t = "p" Then
        SQL = "SELECT oZfor.sapId " _
            & "FROM tbOrders oZfin LEFT JOIN tbOrderDep od ON od.zfinOrder=oZfin.orderId LEFT JOIN tbOrders oZfor ON oZfor.orderId=od.zforOrder " _
            & "WHERE oZfin.orderId=" & oId & " AND (od.isRemoved IS NULL OR od.isRemoved = 0);"
    Else
        SQL = "SELECT oZfin.sapId " _
            & "FROM tbOrders oZfin LEFT JOIN tbOrderDep od ON od.zfinOrder=oZfin.orderId LEFT JOIN tbOrders oZfor ON oZfor.orderId=od.zforOrder " _
            & "WHERE oZfor.orderId=" & oId & " AND (od.isRemoved IS NULL OR od.isRemoved = 0);"
    End If
    Set rs = New ADODB.Recordset
    rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        MsgBox "Currently no " & Me.lblOut.Caption & " orders are connected to the " & Me.lblIn.Caption & " order", vbInformation + vbOKOnly, "No orders connected"
    Else
        rs.MoveFirst
        Do Until rs.EOF
            Me.txtOut.Text = Me.txtOut.Text & rs.Fields("sapId") & ","
            rs.MoveNext
        Loop
        Me.txtOut.Text = Left(Me.txtOut.Text, Len(Me.txtOut.Text) - 1)
    End If
    rs.Close
    Set rs = Nothing
End If
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

End Sub

Private Sub UserForm_Initialize()
Me.btnSave.Enabled = False
Me.lblIn.Visible = False
Me.lblOut.Visible = False
Me.txtOut.Locked = True
End Sub


Private Sub editConnections()
Dim SQL As String
Dim found As Boolean
Dim o As clsOrder
Dim rs As ADODB.Recordset

updateConnection

If t = "p" Then
    SQL = "UPDATE tbOrderDep SET isRemoved=1, RemovedOn=CURRENT_TIMESTAMP WHERE zfinOrder=" & oId
Else
    SQL = "UPDATE tbOrderDep SET isRemoved=1, RemovedOn=CURRENT_TIMESTAMP WHERE zforOrder=" & oId
End If

adoConn.Execute SQL

Me.txtOut.Value = ""

For Each o In orders
    If t = "p" Then
        Set rs = New ADODB.Recordset
        SQL = "SELECT * " _
        & "FROM tbOrderDep od " _
        & "WHERE zfinOrder=" & oId & " AND zforOrder=" & o.orderId
        rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            rs.MoveFirst
            found = True
            rs.Fields("isRemoved").Value = Null
            rs.Update
            rs.Close
            Set rs = Nothing
        Else
            rs.Close
            Set rs = Nothing
            adoConn.Execute "INSERT INTO tbOrderDep (zfinOrder, zforOrder) VALUES (" & oId & "," & o.orderId & ")"
        End If
    Else
        Set rs = New ADODB.Recordset
        SQL = "SELECT * " _
        & "FROM tbOrderDep od " _
        & "WHERE zforOrder=" & oId & " AND zfinOrder=" & o.orderId
        rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            rs.MoveFirst
            found = True
            rs.Fields("isRemoved").Value = Null
            'rs.Close
            rs.Update
            rs.Close
            Set rs = Nothing
        Else
            rs.Close
            Set rs = Nothing
            adoConn.Execute "INSERT INTO tbOrderDep (zfinOrder, zforOrder) VALUES (" & o.orderId & "," & oId & ")"
        End If
    End If
    adoConn.Execute SQL
    Me.txtOut.Text = Me.txtOut.Text & o.sapId & ","
Next o
Me.txtOut.Text = Left(Me.txtOut.Text, Len(Me.txtOut.Text) - 1)

MsgBox "Connections has been updated!", vbOKOnly + vbInformation, "Success"

End Sub

Private Function verify() As Boolean
Dim bool As Boolean
Dim v() As String
Dim o As clsOrder
Dim i As Integer
Dim SQL As String
Dim rs As ADODB.Recordset

bool = False

updateConnection

If oId > 0 Then
    For Each o In orders
        orders.Remove CStr(o.sapId)
    Next o
    v() = Split(Me.txtOut.Text, ",", , vbTextCompare)
    For i = LBound(v) To UBound(v)
        If IsNumeric(v(i)) Then
            SQL = "SELECT orderId FROM tbOrders WHERE sapId=" & v(i)
            Set rs = New ADODB.Recordset
            rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                rs.MoveFirst
                Set o = New clsOrder
                o.sapId = CLng(v(i))
                o.orderId = rs.Fields("orderId")
                orders.Add o, CStr(o.sapId)
            Else
                MsgBox "Order " & v(i) & " doesn't exist in database and will be omitted", vbInformation + vbOKOnly, "Order doesn't exist"
            End If
        Else
            MsgBox "Order " & v(i) & " isn't valid number and will be omitted", vbInformation + vbOKOnly, "Wrong format"
        End If
    Next i
    If orders.Count = 0 Then
        MsgBox "There is no valid order numbers in " & Me.lblOut.Caption & ". There must be at least 1 valid order number to continue", vbInformation + vbOKOnly, "Wrong format"
    Else
        bool = True
    End If
Else
    MsgBox "No order has been provided as input", vbCritical + vbOKOnly, "No input order"
End If

verify = bool
End Function

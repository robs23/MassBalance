VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} logger 
   Caption         =   "Login"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
   OleObjectBlob   =   "logger.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnLogin_Click()
Dim user As Integer

On Error GoTo err_trap

user = Me.lstUsers.List(Me.lstUsers.ListIndex, 0)
logUser (user)

exit_here:
Exit Sub

err_trap:
MsgBox "Please first select your user in the list above", vbOKOnly + vbInformation, "Error"
Resume exit_here

End Sub

Private Sub updateUsers()
Dim i As Integer
Dim rs As ADODB.Recordset
Dim SQL As String


On Error GoTo err_trap

For i = Me.lstUsers.ListCount - 1 To 0 Step -1
    Me.lstUsers.RemoveItem i
Next i

updateConnection

SQL = "SELECT userId, userName, userSurname FROM tbUsers"

Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If rs.Fields("userId") <> 43 Then
            Me.lstUsers.AddItem rs.Fields("userId")
            Me.lstUsers.List(Me.lstUsers.ListCount - 1, 1) = rs.Fields("userName") & " " & rs.Fields("userSurname")
        End If
        rs.MoveNext
    Loop
End If
rs.Close

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""updateUsers"", error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub lblPassRemind_Click()
Dim user As Integer

On Error GoTo err_trap

user = Me.lstUsers.List(Me.lstUsers.ListIndex, 0)
remindPassword (user)

exit_here:
Exit Sub

err_trap:
MsgBox "Please first select your user in the list above", vbOKOnly + vbInformation, "Error"
Resume exit_here

End Sub

Private Sub lstUsers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim user As Integer
user = Me.lstUsers.List(Me.lstUsers.ListIndex, 0)
logUser (user)
End Sub


Private Sub logUser(user As Integer)
updateProperty "userId", user
password.Show
End Sub

Private Sub UserForm_Initialize()
updateUsers
End Sub

Private Sub remindPassword(user As Integer)
Dim password As String
Dim Mail As String
Dim title As String
Dim body As String
Dim rs As ADODB.Recordset
Dim SQL As String

On Error GoTo err_trap

updateConnection

Set rs = New ADODB.Recordset
SQL = "SELECT userMail, userPassword FROM tbUsers WHERE userId = " & user

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    password = rs.Fields("userPassword")
    Mail = rs.Fields("userMail")
    title = "[NPD] Przypomnienie hasła"
    body = toHtml("Tego maila dostajesz ponieważ poprosiłeś o przypomnienie hasła. Twoje hasło to: ") & toHtml(password, True)
    body = body & "<br><br>" & toHtml("Wiadomość wysłana automatycznie, prosimy nie odpowiadać", True)
    sendMail body, title, Mail, , True
    MsgBox "You're password has been sent to " & Mail, vbOKOnly + vbInformation, "Password sent"
End If
rs.Close

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""RemindPassword"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

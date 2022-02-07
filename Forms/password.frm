VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} password 
   Caption         =   "Wprowadź hasło"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3210
   OleObjectBlob   =   "password.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LogMe()
Dim sht As Worksheet
Dim rs As ADODB.Recordset


If Len(Me.txtPassword) = 0 Then
    MsgBox "Please enter your password first, then hit ENTER", vbInformation + vbOKOnly, "Passowrd missing"
Else
    updateConnection

    SQL = "SELECT userPassword FROM tbUsers WHERE userId = " & ThisWorkbook.CustomDocumentProperties("userId")

    Set rs = New ADODB.Recordset
    rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveFirst
        If rs.Fields("userPassword") = Me.txtPassword Then
            updateProperty "isUserLogged", True
            Unload Me
            Unload logger
        Else
            MsgBox "The password you entered doesn't match selected user's password. You can try one more time", vbOKOnly + vbExclamation, "Wrong password"
            Me.txtPassword = ""
        End If
    End If
    rs.Close
    
    closeConnection
End If
End Sub


Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
    LogMe
End If
End Sub


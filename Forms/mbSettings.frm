VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mbSettings 
   Caption         =   "Mass balance settings"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   OleObjectBlob   =   "mbSettings.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mbSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCreate_Click()
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Dim ctrl As MSForms.control
Dim bool As Boolean
Dim bmId As Integer

bool = True
initializeObjects

If validate Then
    updateConnection
    Set rs = New ADODB.Recordset
    'Set rs = Conn.Execute("SELECT * FROM tbBM WHERE bmWeek = " & week, , adCmdText)
    If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
        rs.Open "SELECT * FROM tbBM WHERE bmWeek = " & w & " AND bmYear = " & y & ";", adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable
    Else
        rs.Open "SELECT * FROM tbBM WHERE bmMonth = " & m & " AND bmYear = " & y & ";", adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable
    End If
    
    If rs.EOF Then
        With mbSettings
            rs.Close
'            w.Value = CInt(mbSettings.Controls("WeekNo").Value)
'            y.Value = CInt(mbSettings.Controls("cmbYear").Value)
            rs.Open "tbBm;", adoConn, adOpenKeyset, adLockOptimistic ', adCmdTable
            rs.AddNew
            If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
                rs.Fields("bmWeek").Value = w
            Else
                rs.Fields("bmMonth").Value = m
            End If
            rs.Fields("bmYear").Value = y
            rs.Fields("bmCreatedOn").Value = Now
            rs.Fields("roastingFrom").Value = getDate(.Controls("dRoastingFrom").Value, .Controls("cmbRoastingFrom").Value)
            rs.Fields("roastingTo").Value = getDate(.Controls("dRoastingTo").Value, .Controls("cmbRoastingTo").Value)
            rs.Fields("createdBy").Value = ThisWorkbook.CustomDocumentProperties("userId")
            rs.Update
            rs.Close
            rs.Open "SELECT @@identity AS bmId FROM tbBm;", adoConn, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then
                rs.MoveFirst
                bmId = rs.Fields("bmId")
            End If
            rs.Close
            rs.Open "tbBMDetails;", adoConn, adOpenKeyset, adLockOptimistic, adCmdTable
            rs.AddNew
            rs.Fields("bmId") = bmId
            rs.Fields("receiptLoss") = Me.txtReceipt
            rs.Fields("cleaningLoss") = Me.txtClean
            rs.Fields("zLoss") = Me.txtDiff
            If Not Me.txtBefore = Empty Then rs.Fields("greenCoffeeReceipt") = Me.txtBefore
            If Not Me.txtAfter = Empty Then rs.Fields("bmEndState") = Me.txtAfter
            rs.Update
            rs.Close
            Set rs = Nothing
            MsgBox "Mass balance for " & w & "||" & y & " has been added successfully.", vbOKOnly + vbInformation, "Success"
            Unload Me
'                clear
        End With
    Else
        If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
            MsgBox "Period " & w & "||" & y & " already exists.", vbOKOnly + vbExclamation, "Data duplication"
        Else
            MsgBox "Period " & m & "||" & y & " already exists.", vbOKOnly + vbExclamation, "Data duplication"
        End If
    End If
    closeConnection
End If


End Sub

Private Sub cmbYear_AfterUpdate()
setRanges
End Sub

Private Sub optMonthly_Click()
'change to monthly
ThisWorkbook.CustomDocumentProperties("PeriodType") = "monthly"
changePeriod
End Sub

Private Sub optWeekly_Click()
'change to weekly
ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly"
changePeriod
End Sub

Private Sub changePeriod()
If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
    Me.lblPeriodType.Caption = "Week"
Else
    Me.lblPeriodType.Caption = "Month"
End If
setRanges
End Sub

Private Sub UserForm_Initialize()
Dim i As Integer
Dim year As Integer
For i = 1 To 10
    year = 2015 + i
    Me.cmbYear.AddItem CStr(year)
Next i
setTime
initializeObjects
Me.optWeekly = True
End Sub

Private Sub WeekNo_AfterUpdate()
setRanges
End Sub

Sub setRanges()
Dim mn As Date

y.Value = 0

If Not Me.cmbYear.Value = "" Then
    y.Value = CInt(Me.cmbYear.Value)
End If

If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
    If Me.WeekNo.Value <> "" Then w.Value = Me.WeekNo.Value
    If Not w = 0 And Not y = 0 Then
        Me.dRoastingFrom.Value = 7 * (w - 1) + DateSerial(y, 1, 4) - Weekday(DateSerial(y, 1, 4), 2) - 1
        Me.dRoastingTo.Value = 7 * (w - 1) + DateSerial(y, 1, 4) - Weekday(DateSerial(y, 1, 4), 2) + 6
        Me.cmbRoastingFrom.ListIndex = 44
        Me.cmbRoastingTo.ListIndex = 28
    End If
Else
    If Me.WeekNo.Value <> "" Then m.Value = Me.WeekNo.Value
    If Not m = 0 And Not y = 0 Then
        Me.dRoastingFrom.Value = DateSerial(y, m, 1)
        Me.dRoastingTo.Value = DateAdd("d", -1, DateSerial(y, m + 1, 1))
        Me.cmbRoastingFrom.ListIndex = 12
        Me.cmbRoastingTo.ListIndex = 47
    End If
End If

End Sub


Sub setTime()
Dim i As Integer
Dim h As String
Dim m As String
Dim n As Integer
Dim cmb As ComboBox

For n = 1 To 6
    Select Case n
    Case 1
        Set cmb = Me.cmbRoastingFrom
    Case 2
        Set cmb = Me.cmbRoastingTo
    End Select
    h = -1
    For i = 0 To 48
        If i Mod 2 = 0 Then
            h = h + 1
            m = "00"
        Else
            m = "30"
        End If
        cmb.AddItem h & ":" & m
    Next i
Next n
End Sub

Private Function validate() As Boolean
Dim bool As Boolean
Dim ctrl As MSForms.control

bool = True
If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
    For Each ctrl In mbSettings.Controls
        If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then
            If ctrl.Value = "" Then
                bool = False
                Exit For
            End If
        End If
    Next ctrl
Else
    For Each ctrl In mbSettings.Controls
        If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then
            If ctrl.Value = "" And ctrl.Name <> "txtBefore" And ctrl.Name <> "txtAfter" Then
                bool = False
                Exit For
            End If
        End If
    Next ctrl
End If

If bool = True Then
    bool = False
    If IsNumeric(Me.WeekNo) And IsNumeric(Me.cmbYear) Then
        w = Me.WeekNo
        y = Me.cmbYear
        If ((w > 53 Or w < 1 Or y > 2025 Or y < 2015) And ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly") Or ((m < 1 Or m > 12) And ThisWorkbook.CustomDocumentProperties("PeriodType") = "monthly") Then
            MsgBox "You put wrong parameters in. Week must be between 1 - 53 and year between 2015 - 2025. Correct parameters and try again.", vbOKOnly + vbExclamation, "Wrong input data"
        Else
            If Not IsDate(Me.dRoastingFrom) Or Not IsDate(Me.dRoastingTo) Then
                MsgBox "Either ""From"" or ""To"" is in wrong format. Correct parameters and try again.", vbOKOnly + vbExclamation, "Wrong input data"
            Else
                If Me.dRoastingFrom > Me.dRoastingTo Then
                    MsgBox """From"" date can't be later than ""To"" date. Correct parameters and try again.", vbOKOnly + vbExclamation, "Wrong input data"
                Else
                    If IsNumeric(Me.txtClean) And IsNumeric(Me.txtDiff) And IsNumeric(Me.txtReceipt) Then
                        bool = True
                    Else
                        MsgBox "You put non-numeric value in ""Loss on receipt""/""Loss on cleaning""/""Differences"". Correct parameters and try again.", vbOKOnly + vbExclamation, "Wrong input data"
                    End If
                End If
            End If
        End If
    Else
        MsgBox "You put non-numeric value in week/month/year. Correct parameters and try again.", vbOKOnly + vbExclamation, "Wrong input data"
    End If
Else
    MsgBox "All fields must be filled in to continue.", vbOKOnly + vbExclamation, "Missing data"
End If

validate = bool
End Function


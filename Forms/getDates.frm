VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} getDates 
   Caption         =   "Choose roasting dates to start with"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   OleObjectBlob   =   "getDates.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "getDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
getDates.Hide
End Sub

Private Sub btnOK_Click()
Dim dFrom As Date
Dim dTo As Date
Dim ex As Variant
Dim fa As Variant
Dim var As Variant
Dim exp As Variant 'array of expansion blends
Dim beanOption As Integer '0-all, 1-ground, 2-beans
Dim bool As Boolean
Dim firstDate As Date

firstDate = #10/2/2016#

Me.MultiPage1.Value = 0

bool = True
period = ""

With getDates
    If Me.oWeekly = True Then
        ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly"
        var = validate
        If IsArray(var) Then
            dFrom = var(0)
            dTo = var(1)
            updateProperty "week", Me.cmbWeek
            updateProperty "year", Me.cmbYear
            period = Me.cmbWeek & "|" & Me.cmbYear
            Application.Caption = "Loaded week " & period
        Else
            bool = False
            MsgBox "No data for chosen period yet. You can try custom mode for similar period of time", vbOKOnly + vbExclamation, "No data"
        End If
    ElseIf Me.oMonthly = True Then
        ThisWorkbook.CustomDocumentProperties("PeriodType") = "monthly"
        var = validate
        If IsArray(var) Then
            dFrom = var(0)
            dTo = var(1)
            updateProperty "Month", Me.cmbWeek
            updateProperty "year", Me.cmbYear
            period = Me.cmbWeek & "|" & Me.cmbYear
            Application.Caption = "Loaded month " & period
        Else
            bool = False
            MsgBox "No data for chosen period yet. You can try custom mode for similar period of time", vbOKOnly + vbExclamation, "No data"
        End If
    Else
        If Not IsDate(txtFrom.Value) Or Not IsDate(txtTo) Then
            bool = False
            MsgBox "Both ""from"" and ""to"" fields are required", vbExclamation + vbOKOnly, "Error in dates range"
        Else
            ThisWorkbook.CustomDocumentProperties("PeriodType") = "custom"
            dFrom = getDate(Me.txtFrom.Value, Me.cmbRoastingFrom.Value)
            dTo = getDate(Me.txtTo.Value, Me.cmbRoastingTo.Value)
            updateProperty "week", 0
            updateProperty "year", 0
            updateProperty "Month", 0
            Application.Caption = "Loaded period " & DateSerial(year(dFrom), month(dFrom), Day(dFrom)) & " - " & DateSerial(year(dTo), month(dTo), Day(dTo))
        End If
    End If
    If dFrom < firstDate Then
        bool = False
        MsgBox "Currently there's no data earlier than " & firstDate & ". Please change desired time span", vbOKOnly + vbExclamation, "No data"
    End If
    If bool Then
        If Len(Me.txtExclude) > 0 And Len(Me.cmbLimiter) = 0 Then
            MsgBox "You have put some blends to be excluded/limited to but you haven't chosen if they should be excluded or limited to. Please choose one of options from drop-down field", vbOKOnly + vbExclamation, "Error"
        Else
            If dFrom > dTo Then
                MsgBox "Date ""from"" can't be later than ""to""", vbExclamation + vbOKOnly, "Error in dates range"
            Else
                If Not IsNumeric(Me.txtProgress) Or Me.txtProgress > 100 Or Me.txtProgress < 0 Then
                    MsgBox "In progress value needs to be number in range from 0 to 100. Please correct it (tab ""Options"")", vbExclamation + vbOKOnly, "Wrong data type"
                Else
                    If Len(Me.cmbGsource) = 0 Or Len(Me.cmbPsource) = 0 Or Len(Me.cmbRsource) = 0 Then
                        MsgBox "Choose data source for all production stages from drop-down list (tab ""Data source""). Each stage (roasting, grinding, packing) must have data source assigned in order to continue", vbOKOnly + vbExclamation, "Data source missing"
                    Else
                        If IsNull(Me.cmbBeans) Then
                            beanOption = 0
                        Else
                            beanOption = Me.cmbBeans.ListIndex
                        End If
                        ex = getExcludedBlends
                        fa = getFormating
                        exp = getExpansionBlends
                        If IsArray(ex) And IsArray(fa) Then
                            If IsArray(exp) Then
                                If Me.cmbLimiter.ListIndex = 0 And Me.cmbBlends2Expand.ListIndex = 0 Then 'exclude blends & expansion for certain blends
                                    scadaSummary dFrom, dTo, beanOption, ex, , fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource, exp
                                ElseIf Me.cmbLimiter.ListIndex = 0 And Me.cmbBlends2Expand.ListIndex = 1 Then 'exclude blends & expansion NOT for certain blends
                                    scadaSummary dFrom, dTo, beanOption, ex, , fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource, , exp
                                ElseIf Me.cmbLimiter.ListIndex = 1 And Me.cmbBlends2Expand.ListIndex = 0 Then 'limit to blends & expansion for certain blends
                                    scadaSummary dFrom, dTo, beanOption, , ex, fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource, exp
                                ElseIf Me.cmbLimiter.ListIndex = 1 And Me.cmbBlends2Expand.ListIndex = 1 Then 'limit to blends & expansion NOT for certain blends
                                    scadaSummary dFrom, dTo, beanOption, , ex, fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource, , exp
                                End If
                            Else
                                If Me.cmbLimiter.ListIndex = 0 Then
                                    scadaSummary dFrom, dTo, beanOption, ex, , fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource
                                Else
                                    scadaSummary dFrom, dTo, beanOption, , ex, fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource
                                End If
                            End If
                        ElseIf IsArray(ex) = False And IsArray(fa) Then
                            If IsArray(exp) Then
                                If Me.cmbBlends2Expand.ListIndex = 0 Then 'expansion for certain blends
                                    scadaSummary dFrom, dTo, beanOption, , , fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource, exp
                                Else
                                    scadaSummary dFrom, dTo, beanOption, , , fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource, , exp
                                End If
                            Else
                                scadaSummary dFrom, dTo, beanOption, , , fa, Me.txtProgress, Me.cmbGsource, Me.cmbPsource
                            End If
                        ElseIf IsArray(ex) And IsArray(fa) = False Then
                            If IsArray(exp) Then
                                If Me.cmbLimiter.ListIndex = 0 And Me.cmbBlends2Expand.ListIndex = 0 Then 'exclude blends & expansion for certain blends
                                    scadaSummary dFrom, dTo, beanOption, ex, , , Me.txtProgress, Me.cmbGsource, Me.cmbPsource, exp
                                ElseIf Me.cmbLimiter.ListIndex = 0 And Me.cmbBlends2Expand.ListIndex = 1 Then 'exclude to blends & expansion NOT for certain blends
                                    scadaSummary dFrom, dTo, beanOption, ex, , , Me.txtProgress, Me.cmbGsource, Me.cmbPsource, , exp
                                ElseIf Me.cmbLimiter.ListIndex = 1 And Me.cmbBlends2Expand.ListIndex = 0 Then 'limit to blends & expansion for certain blends
                                    scadaSummary dFrom, dTo, beanOption, , ex, , Me.txtProgress, Me.cmbGsource, Me.cmbPsource, exp
                                ElseIf Me.cmbLimiter.ListIndex = 1 And Me.cmbBlends2Expand.ListIndex = 1 Then 'limit to blends & expansion NOT for certain blends
                                    scadaSummary dFrom, dTo, beanOption, , ex, , Me.txtProgress, Me.cmbGsource, Me.cmbPsource, , exp
                                End If
                            Else
                                If Me.cmbLimiter.ListIndex = 0 Then
                                    scadaSummary dFrom, dTo, beanOption, ex, , , Me.txtProgress, Me.cmbGsource, Me.cmbPsource
                                Else
                                    scadaSummary dFrom, dTo, beanOption, , ex, , Me.txtProgress, Me.cmbGsource, Me.cmbPsource
                                End If
                            End If
                        Else
                            If IsArray(exp) Then 'expansion for certain blends
                                If Me.cmbBlends2Expand.ListIndex = 0 Then
                                    scadaSummary dFrom, dTo, beanOption, , , , Me.txtProgress, Me.cmbGsource, Me.cmbPsource, exp
                                Else
                                    scadaSummary dFrom, dTo, beanOption, , , , Me.txtProgress, Me.cmbGsource, Me.cmbPsource, , exp
                                End If
                            Else
                                scadaSummary dFrom, dTo, beanOption, , , , Me.txtProgress, Me.cmbGsource, Me.cmbPsource
                            End If
                        End If
                        Me.Hide
                    End If
                End If
            End If
        End If
    End If
End With

End Sub


Private Sub oCustom_Click()
Me.cmbWeek.Enabled = False
Me.cmbYear.Enabled = False
Me.txtFrom.Enabled = True
Me.txtTo.Enabled = True
Me.cmbRoastingFrom.Enabled = True
Me.cmbRoastingTo.Enabled = True
End Sub

Private Sub oMonthly_Click()
Me.cmbWeek.Enabled = True
Me.cmbYear.Enabled = True
Me.txtFrom.Enabled = False
Me.txtTo.Enabled = False
Me.cmbRoastingFrom.Enabled = False
Me.cmbRoastingTo.Enabled = False
generatePeriods
Me.lblPeriod.Caption = "Month"
End Sub

Private Sub oWeekly_Click()
Me.cmbWeek.Enabled = True
Me.cmbYear.Enabled = True
Me.txtFrom.Enabled = False
Me.txtTo.Enabled = False
Me.cmbRoastingFrom.Enabled = False
Me.cmbRoastingTo.Enabled = False
generatePeriods
Me.lblPeriod.Caption = "Week"
End Sub

Private Sub UserForm_Initialize()
Dim sDate As Date
Dim eDate As Date
Dim sTime As Variant
Dim eTime As Variant
Dim i As Integer

On Error GoTo err_trap

Me.MultiPage1.Value = 0
Me.oWeekly = True
setCustom
fillCmbs

For i = Me.cmbBeans.ListCount - 1 To 0 Step -1
    Me.cmbBeans.RemoveItem i
Next i

For i = Me.cmbLimiter.ListCount To 1 Step -1
    Me.cmbLimiter.RemoveItem i
Next i

Me.txtProgress = "10"
Me.cmbBeans.AddItem "All included"
Me.cmbBeans.AddItem "Ground only"
Me.cmbBeans.AddItem "Beans only"

Me.cmbLimiter.AddItem "Exlude"
Me.cmbLimiter.AddItem "Limit to"

Me.cmbBlends2Expand.AddItem "allowed"
Me.cmbBlends2Expand.AddItem "not allowed"

Me.cmbBeans.ListIndex = 0 'select the first item

Me.MultiPage1.Value = 0 'this stops windows date & time picker control error

Me.cboxAllowExpansion.Value = False

deployDsources

exit_here:
Exit Sub

err_trap:
If Err.Number = 5 Then
    Resume Next
Else
    MsgBox "Error in user form ""getDates"". Error number: " & Err.Number & ", " & Err.Description
    Resume exit_here
End If

End Sub

Sub setTime()
Dim i As Integer
Dim h As String
Dim m As String
Dim n As Integer
Dim cmb As ComboBox

For n = Me.cmbRoastingFrom.ListCount To 1 Step -1
    Me.cmbRoastingFrom.RemoveItem n
    Me.cmbRoastingTo.RemoveItem n
Next n


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

Private Function timeToList(time As String) As Integer
'takes time in "23:59" format and converts it to list item
Dim h As Integer
Dim m As Integer
h = CInt(Left(time, 2))
If CInt(Right(time, 2)) = 0 Then m = 0 Else m = 1
timeToList = (h * 2) + m
End Function

Private Function getExcludedBlends() As Variant
Dim v() As String
Dim i As Integer

If Len(Me.txtExclude) = 0 Then
    getExcludedBlends = False
Else
    v = Split(Me.txtExclude, ",")
    getExcludedBlends = v
End If
End Function

Private Function getExpansionBlends() As Variant
Dim v() As String
Dim i As Integer

If Len(Me.txtExpansionFor) = 0 Then
    getExpansionBlends = False
Else
    v = Split(Me.txtExpansionFor, ",")
    getExpansionBlends = v
End If
End Function

Private Function getFormating() As Variant
Dim arr(9) As Variant

If Me.oKg = True Or Me.oPercent = True Then
    'we've found some formating
    If Me.oKg = True Then
        arr(0) = 0 'kg
    Else
        arr(0) = 1 '%
    End If
    If Me.cboxAbs = True Then
        arr(1) = True
    Else
        arr(1) = False
    End If
    If Me.txtHR = "" Then
        arr(2) = Null
    Else
        arr(2) = CDbl(Me.txtHR)
    End If
    If Me.txtLR = "" Then
        arr(3) = Null
    Else
        arr(3) = CDbl(Me.txtLR)
    End If
    If Me.txtHG = "" Then
        arr(4) = Null
    Else
        arr(4) = CDbl(Me.txtHG)
    End If
    If Me.txtLG = "" Then
        arr(5) = Null
    Else
        arr(5) = CDbl(Me.txtLG)
    End If
    If Me.txtHP = "" Then
        arr(6) = Null
    Else
        arr(6) = CDbl(Me.txtHP)
    End If
    If Me.txtLP = "" Then
        arr(7) = Null
    Else
        arr(7) = CDbl(Me.txtLP)
    End If
    If Me.txtHE = "" Then
        arr(8) = Null
    Else
        arr(8) = CDbl(Me.txtHE)
    End If
    If Me.txtLE = "" Then
        arr(9) = Null
    Else
        arr(9) = CDbl(Me.txtLE)
    End If
    getFormating = arr
Else
    getFormating = False
End If

End Function

Private Sub fillCmbs()
Dim i As Integer

For i = Me.cmbSortType.ListCount To 1 Step -1
    Me.cmbSortType.RemoveItem i
Next i

For i = Me.cmbSortOrder.ListCount To 1 Step -1
    Me.cmbSortOrder.RemoveItem i
Next i

Me.cmbSortType.AddItem "Roasting loss in kg"
Me.cmbSortType.AddItem "Roasting loss in %"
Me.cmbSortType.AddItem "Grinding loss in kg"
Me.cmbSortType.AddItem "Grinding loss in %"
Me.cmbSortType.AddItem "Packing loss in kg"
Me.cmbSortType.AddItem "Packing loss in %"
Me.cmbSortType.AddItem "Grinding + packing loss in kg"
Me.cmbSortType.AddItem "Grinding + packing loss in %"
Me.cmbSortType.AddItem "Total loss in kg"
Me.cmbSortType.AddItem "Total loss in %"
Me.cmbSortType.AddItem "Real vs BOM for roasting + grinding in %"
Me.cmbSortType.AddItem "Real vs BOM for packing in %"
Me.cmbSortType.AddItem "Real vs BOM for total in %"
Me.cmbSortType.AddItem "Roasting loss on RN3000 in %"
Me.cmbSortType.AddItem "Roasting loss on RN3000 in kg"
Me.cmbSortType.AddItem "Roasting loss on RN4000 in %"
Me.cmbSortType.AddItem "Roasting loss on RN4000 in kg"
Me.cmbSortType.AddItem "Roasted coffe value"
Me.cmbSortType.AddItem "Packed Coffee value"
Me.cmbSortType.AddItem "Lost value on grinding + packing"
Me.cmbSortType.AddItem "Roasting loss vs average"
Me.cmbSortType.AddItem "Grinding loss vs average"
Me.cmbSortType.AddItem "Packing loss vs average"
Me.cmbSortType.AddItem "Total loss vs average"
Me.cmbSortType.AddItem "Grinding+Packing loss vs average"
Me.cmbSortOrder.AddItem "ASC"
Me.cmbSortOrder.AddItem "DESC"

End Sub

Private Sub deployDsources()
Dim i As Integer

For i = Me.cmbRsource.ListCount To 1 Step -1
    Me.cmbRsource.RemoveItem i
Next i

For i = Me.cmbGsource.ListCount To 1 Step -1
    Me.cmbGsource.RemoveItem i
Next i

For i = Me.cmbPsource.ListCount To 1 Step -1
    Me.cmbPsource.RemoveItem i
Next i

Me.cmbRsource.AddItem "SCADA"
Me.cmbRsource.ListIndex = 0
Me.cmbGsource.AddItem "SAP"
Me.cmbGsource.AddItem "MES"
Me.cmbGsource.ListIndex = 0
Me.cmbPsource.AddItem "SAP"
Me.cmbPsource.AddItem "MES"
Me.cmbPsource.ListIndex = 0

End Sub

Private Sub generatePeriods()
Dim i As Integer

For i = Me.cmbWeek.ListCount - 1 To 0 Step -1
    Me.cmbWeek.RemoveItem i
Next i

For i = Me.cmbYear.ListCount - 1 To 0 Step -1
    Me.cmbYear.RemoveItem i
Next i

If Me.oWeekly Then
    For i = 1 To 53
        Me.cmbWeek.AddItem i
    Next i
Else
    For i = 1 To 12
        Me.cmbWeek.AddItem i
    Next i
End If

For i = 2016 To 2025
    Me.cmbYear.AddItem i
Next i
End Sub

Private Sub setCustom()
Dim sDate As Date
Dim eDate As Date
Dim rs As ADODB.Recordset
Dim SQL As String

On Error GoTo err_trap

Set rs = New ADODB.Recordset

updateConnection

SQL = "SELECT TOP 1 roastingFrom,roastingTo FROM tbBM ORDER BY bmCreatedOn DESC"
Set rs = adoConn.Execute(SQL)
If Not rs.EOF Then
    rs.MoveFirst
    sDate = rs.Fields("roastingFrom")
    eDate = rs.Fields("roastingTo")
End If
rs.Close

sTime = getTime(sDate)
eTime = getTime(eDate)

If Not IsDate(sDate) Or year(sDate) < 2010 Then sDate = DateAdd("d", -7, Date)
If Not IsDate(eDate) Or year(eDate) < 2010 Then eDate = Date
Me.txtFrom.Value = DateValue(sDate)
Me.txtTo.Value = DateValue(eDate)
setTime
If IsNull(sTime) Then
    Me.cmbRoastingFrom.ListIndex = 0
Else
    Me.cmbRoastingFrom.ListIndex = timeToList(CStr(sTime))
End If

If IsNull(eTime) Then
    Me.cmbRoastingTo.ListIndex = 0
Else
    Me.cmbRoastingTo.ListIndex = timeToList(CStr(eTime))
End If

exit_here:
Set rs = Nothing
closeConnection
Exit Sub

err_trap:
MsgBox "Error in setCustom of form ""getDates"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Function validate() As Variant
Dim bool As Variant
Dim rs As ADODB.Recordset
Dim SQL As String
Dim arr(1) As Date

On Error GoTo err_trap

bool = False

If Me.oWeekly Or Me.oMonthly Then
    Set rs = New ADODB.Recordset

    updateConnection
    If Me.oWeekly Then
        SQL = "SELECT roastingFrom,roastingTo FROM tbBM WHERE bmWeek = " & Me.cmbWeek & " AND bmYear = " & Me.cmbYear
    Else
        SQL = "SELECT roastingFrom,roastingTo FROM tbBM WHERE bmMonth = " & Me.cmbWeek & " AND bmYear = " & Me.cmbYear
    End If
    Set rs = adoConn.Execute(SQL)
    If Not rs.EOF Then
        rs.MoveFirst
        arr(0) = rs.Fields("roastingFrom")
        arr(1) = rs.Fields("roastingTo")
        bool = arr
    End If
    rs.Close
End If

exit_here:
validate = bool
Set rs = Nothing
closeConnection
Exit Function

err_trap:
MsgBox "Error in validate of form ""getDates"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

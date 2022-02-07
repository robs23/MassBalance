VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} graphSettings 
   Caption         =   "Chart update settings"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3390
   OleObjectBlob   =   "graphSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "graphSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnUpdate_Click()
Dim w0 As Integer
Dim y0 As Integer
Dim w1 As Integer
Dim y1 As Integer
Dim x As Integer
Dim SQL As String

SQL = ""

If Me.cmbOptions.ListIndex = 0 Then
    If Me.cmbType.ListIndex = 0 Then
        'weekly
        SQL = "SELECT * " _
            & "FROM tbBM JOIN tbBMDetails ON tbBM.bmId = tbBMDetails.bmId " _
            & "WHERE tbBm.bmMonth IS NULL AND tbBm.bmYear = " & year(Date) _
            & " ORDER BY tbbm.bmYear DESC, tbBM.bmWeek DESC"
    Else
        'monthly
        SQL = "SELECT * " _
            & "FROM tbBM JOIN tbBMDetails ON tbBM.bmId = tbBMDetails.bmId " _
            & "WHERE tbBm.bmMonth IS NOT NULL AND tbBm.bmYear = " & year(Date) _
            & " ORDER BY tbbm.bmYear DESC, tbBM.bmMonth DESC"
    End If
ElseIf Me.cmbOptions.ListIndex = 1 Then
    If IsNumeric(Me.txtX.Value) Then
        If Me.txtX.Value < 1 Or Me.txtX.Value > 200 Then
            MsgBox "Please provide number in range 1 - 200", vbOKOnly + vbInformation, "Wrong value"
        Else
            x = Me.txtX.Value
            If Me.cmbType.ListIndex = 0 Then
                'weekly
                SQL = "SELECT TOP(" & x & ") * " _
                    & "FROM tbBM JOIN tbBMDetails ON tbBM.bmId = tbBMDetails.bmId " _
                    & "WHERE tbBM.bmMonth IS NULL " _
                    & "ORDER BY tbbm.bmYear DESC, tbBM.bmWeek DESC"
            Else
                'monthly
                SQL = "SELECT TOP(" & x & ") * " _
                    & "FROM tbBM JOIN tbBMDetails ON tbBM.bmId = tbBMDetails.bmId " _
                    & "WHERE tbBM.bmMonth IS NOT NULL " _
                    & "ORDER BY tbbm.bmYear DESC, tbBM.bmMonth DESC"
            End If
        End If
    Else
        MsgBox "You have to provide numeric value in range 1 - 200", vbOKOnly + vbInformation, "Non-numeric value"
    End If
ElseIf Me.cmbOptions.ListIndex = 2 Then
    If IsDate(Me.dFrom.Value) And IsDate(Me.dTo.Value) Then
        If Me.dFrom.Value > Me.dTo.Value Then
            MsgBox "Start date must be earlier than finish date", vbOKOnly + vbInformation, "Date error"
        Else
            x = DateDiff("ww", Me.dFrom.Value, Me.dTo.Value)
            w1 = IsoWeekNumber(Me.dTo.Value)
            y1 = year(Me.dTo.Value)
            If Me.cmbType.ListIndex = 0 Then
                'weekly
                SQL = "SELECT TOP(" & x & ") * " _
                    & "FROM tbBM JOIN tbBMDetails ON tbBM.bmId = tbBMDetails.bmId " _
                    & "WHERE tbbm.bmMonth IS NULL AND tbbm.bmWeek <= " & w0 & " And tbbm.bmYear <= " & y0 _
                    & " ORDER BY tbbm.bmYear DESC, tbBM.bmWeek DESC"
            Else
                'monthly
                SQL = "SELECT TOP(" & x & ") * " _
                    & "FROM tbBM JOIN tbBMDetails ON tbBM.bmId = tbBMDetails.bmId " _
                    & "WHERE tbbm.bmMonth IS NOT NULL AND tbbm.bmWeek <= " & w0 & " And tbbm.bmYear <= " & y0 _
                    & " ORDER BY tbbm.bmYear DESC, tbBM.bmMonth DESC"
            End If
        End If
    Else
        MsgBox "Both fields must be filled with value in date format", vbOKOnly + vbInformation, "Wrong value"
    End If
End If

If Len(SQL) > 0 Then
    bringResults SQL
    Me.Hide
End If

End Sub

Private Sub cmbOptions_Change()
If Me.cmbOptions.ListIndex = 0 Then
    Me.lX.Visible = False
    Me.txtX.Visible = False
    Me.dFrom.Visible = False
    Me.dTo.Visible = False
    Me.lFrom.Visible = False
    Me.lTo.Visible = False
ElseIf Me.cmbOptions.ListIndex = 1 Then
    Me.lX.Visible = True
    Me.txtX.Visible = True
    Me.dFrom.Visible = False
    Me.dTo.Visible = False
    Me.lFrom.Visible = False
    Me.lTo.Visible = False
ElseIf Me.cmbOptions.ListIndex = 2 Then
    Me.lX.Visible = False
    Me.txtX.Visible = False
    Me.dFrom.Visible = True
    Me.dTo.Visible = True
    Me.lFrom.Visible = True
    Me.lTo.Visible = True
End If
End Sub

Private Sub UserForm_Initialize()
Me.cmbOptions.clear
Me.cmbOptions.AddItem "This year"
Me.cmbOptions.AddItem "Last X weeks"
Me.cmbOptions.AddItem "Date range"
Me.cmbOptions.ListIndex = 0
Me.cmbType.clear
Me.cmbType.AddItem "Weekly"
Me.cmbType.AddItem "Monthly"
Me.cmbType.ListIndex = 0
End Sub

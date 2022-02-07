Attribute VB_Name = "Module1"
Public w As DocumentProperty
Public y As DocumentProperty
Public m As DocumentProperty
Public currentWeekID As Variant
Public currentYearID As Variant
Public weekCtl As IRibbonControl
Public yearCtl As IRibbonControl
Public TotalRecords As Integer
Public weekLoaded As Integer
Public rib As IRibbonUI

Sub GetSelectedWeekID(control As IRibbonControl, ByRef itemID As Variant)
initializeObjects
    Set weekCtl = control
    If IsEmpty(currentWeekID) Then
        If Not IsEmpty(w) Then
            currentWeekID = "ddWeek" & w
        Else
            currentWeekID = "ddWeek" & IsoWeekNumber(Date)
            w.Value = IsoWeekNumber(Date)
        End If
    End If
    itemID = currentWeekID
End Sub

Sub GetSelectedYearID(control As IRibbonControl, ByRef itemID As Variant)
initializeObjects
    Dim i As Integer
    
    If IsEmpty(currentYearID) Then
        If IsEmpty(y) Then
            i = y - 2014
        Else
             i = year(Date) - 2014
        End If
        currentYearID = "ddYear" & i
        y.Value = i + 2014
    End If
    itemID = currentYearID
End Sub


Public Sub OnRibbonLoad(objRibbon As IRibbonUI)
    Set rib = objRibbon
End Sub

Sub changeWeek(control As IRibbonControl, selectedID As String, _
             selectedIndex As Integer)
initializeObjects
    currentWeekID = selectedID
    w.Value = CInt(Right(currentWeekID, Len(currentWeekID) - 6))
End Sub

Public Sub nextWeek(control As IRibbonControl)
initializeObjects

If Not IsEmpty(currentWeekID) And w < 53 Then
    w.Value = CInt(Right(currentWeekID, Len(currentWeekID) - 6)) + 1
    Call changeWeek(weekCtl, "ddWeek" & w, w - 1)
    rib.InvalidateControl "ddWeeks"
'    rib.InvalidateControl "btnSave"
End If
End Sub

Public Sub nextYear(control As IRibbonControl)
Dim i As Integer
initializeObjects

If Not IsEmpty(currentYearID) And y < 2020 Then
    i = CInt(Right(currentYearID, Len(currentYearID) - 6)) + 1
    y.Value = y + 1
    Call changeYear(yearCtl, "ddYear" & i, i - 1)
    rib.InvalidateControl "ddYears"
'    rib.InvalidateControl "btnSave"
End If
    End Sub

Public Sub prevYear(control As IRibbonControl)
Dim i As Integer
initializeObjects

If Not IsEmpty(currentYearID) And y > 2015 Then
    i = CInt(Right(currentYearID, Len(currentYearID) - 6)) - 1
    y.Value = y - 1
    Call changeYear(yearCtl, "ddYear" & i, i - 1)
    rib.InvalidateControl "ddYears"
'    rib.InvalidateControl "btnSave"
End If
End Sub

Public Sub prevWeek(control As IRibbonControl)
initializeObjects
If Not IsEmpty(currentWeekID) And w > 1 Then
    w.Value = CInt(Right(currentWeekID, Len(currentWeekID) - 6)) - 1
    Call changeWeek(weekCtl, "ddWeek" & w, w - 1)
    rib.InvalidateControl "ddWeeks"
'    rib.InvalidateControl "btnSave"
End If
End Sub

Sub changeYear(control As IRibbonControl, selectedID As String, _
             selectedIndex As Integer)
initializeObjects
    currentYearID = selectedID
    y.Value = CInt(Right(currentYearID, Len(currentYearID) - 6)) + 2014
End Sub

Public Sub toSettings(control As IRibbonControl)
mbOtherSettings.Show
End Sub

Public Sub prepare(control As IRibbonControl)

If propertyExists("userId") And propertyExists("isUserLogged") Then
    If ThisWorkbook.CustomDocumentProperties("userId") > 0 And ThisWorkbook.CustomDocumentProperties("isUserLogged") = True Then
        If authorize(60) Then
            mbSettings.Show
        End If
    Else
        MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
        logger.Show
    End If
Else
    MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
    logger.Show
End If

End Sub

Public Sub warehouseOverview(control As IRibbonControl)
getStockData
End Sub

Public Sub clear()
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("BM")

sht.Unprotect
With sht
    .Range("B8") = ""
    .Range("D12") = ""
    .Range("D14") = ""
    .Range("D16") = ""
    .Range("G12") = ""
    .Range("G14") = ""
    .Range("M12") = ""
    .Range("M14") = ""
    .Range("S12") = ""
    .Range("S14") = ""
    .Range("Y12") = ""
    .Range("Y14") = ""
    .Range("AE12") = ""
    .Range("AE14") = ""
End With
sht.Protect
End Sub

Public Sub save(control As IRibbonControl)
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String
Dim bmId As Long
Dim res As VbMsgBoxResult
Dim blend As clsBlend
Dim bool As Boolean
Dim saved As Boolean
Dim p As Integer
Dim bk As clsBlendKeeper
Dim reworkAtPackingBeans As Double
Dim reworkAtPackingGround As Double

Set bk = New clsBlendKeeper

On Error GoTo err_trap
initializeObjects

bool = False

Set bk = New clsBlendKeeper

'If w > 0 And y > 0 Then
'    bool = True
'End If
If propertyExists("userId") And propertyExists("isUserLogged") Then
    If ThisWorkbook.CustomDocumentProperties("userId") > 0 And ThisWorkbook.CustomDocumentProperties("isUserLogged") = True Then
        If authorize(60) Then
            bool = True
        End If
    Else
        MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
        logger.Show
    End If
Else
    MsgBox "This function requires that you log in first", vbOKOnly + vbInformation, "User not logged in"
    logger.Show
End If

updateConnection
saved = True

If bool Then
    If (w = 0 And m = 0) Or y = 0 Then
        MsgBox "You can't upload results for periods other than week or month. Currently no week/month has been chosen. In order to choose a week/month click ""Summary"", choose a ""weekly""/""monthly"", specify week/month & year. Afterwards press OK to prepare the report", vbOKOnly + vbInformation, "Period error"
    Else
        
        bk.restoreFromSheet
        downloadZfins "'zfor'"
        For Each blend In bk.blends
            If zfins(CStr(blend.index)).IsBeans Then
                reworkAtPackingBeans = reworkAtPackingBeans + blend.reworkAtPacking
            Else
                reworkAtPackingGround = reworkAtPackingGround + blend.reworkAtPacking
            End If
        Next blend
        Set rs = New ADODB.Recordset
        'Set rs = Conn.Execute("SELECT * FROM tbBM WHERE bmWeek = " & week, , adCmdText)
        If ThisWorkbook.CustomDocumentProperties("PeriodType") = "weekly" Then
            p = w
            rs.Open "SELECT * FROM tbBM WHERE bmWeek = " & w & " AND bmYear = " & y & ";", adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable
        ElseIf ThisWorkbook.CustomDocumentProperties("PeriodType") = "monthly" Then
            p = m
            rs.Open "SELECT * FROM tbBM WHERE bmMonth = " & m & " AND bmYear = " & y & ";", adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable
        End If
        If Not rs.EOF Then
            With ThisWorkbook.Sheets("BM")
                rs.MoveFirst
                bmId = rs.Fields("bmId").Value
                rs.Fields("bmLastModifiedOn").Value = Now
                rs.Update
                rs.Close
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                rs.Open "SELECT * FROM tbBMDetails WHERE bmId = " & bmId & ";", adoConn, adOpenDynamic, adLockOptimistic ', adCmdTable
                If Not rs.EOF Then
                    res = MsgBox("There is already data for period " & p & "|" & y & ". Do you want to overwrite it?", vbOKCancel + vbExclamation, "Confirm overwriting")
                    If res = vbOK Then
                        saved = True
                        rs.MoveFirst
                        rs.Fields("inRoastBeans").Value = .Range("G12")
                        rs.Fields("inRoastGround").Value = .Range("G14")
                        rs.Fields("outRoastBeans").Value = .Range("M12")
                        rs.Fields("outRoastGround").Value = .Range("M14")
                        rs.Fields("receiptLoss").Value = .Range("D12")
                        rs.Fields("cleaningLoss").Value = .Range("D14")
                        rs.Fields("zLoss").Value = .Range("D16")
                        rs.Fields("inPackGround").Value = .Range("S14")
                        rs.Fields("inPackBeans").Value = .Range("S12")
                        rs.Fields("outPackGround").Value = .Range("Y14")
                        rs.Fields("outPackBeans").Value = .Range("Y12")
                        rs.Fields("pwGround").Value = .Range("AE14")
                        rs.Fields("pwBeans").Value = .Range("AE12")
                        rs.Fields("gpValueLoss").Value = ThisWorkbook.Sheets("Operations sequence").Range("AM3")
                        rs.Fields("rework").Value = ThisWorkbook.Sheets("Operations sequence").Range("AN3")
                        rs.Fields("reworkAtPackingBeans").Value = reworkAtPackingBeans
                        rs.Fields("reworkAtPackingGround").Value = reworkAtPackingGround
                        rs.Fields("bmId").Value = bmId
                        rs.Update
                    Else
                        saved = False
                    End If
                Else
                    rs.Close
                    rs.Open "tbBMDetails", conn, adOpenKeyset, adLockOptimistic, adCmdTable
                    rs.AddNew
                    rs.Fields("inRoastBeans").Value = .Range("G12")
                    rs.Fields("inRoastGround").Value = .Range("G14")
                    rs.Fields("outRoastBeans").Value = .Range("M12")
                    rs.Fields("outRoastGround").Value = .Range("M14")
                    rs.Fields("receiptLoss").Value = .Range("D12")
                    rs.Fields("cleaningLoss").Value = .Range("D14")
                    rs.Fields("zLoss").Value = .Range("D16")
                    rs.Fields("inPackGround").Value = .Range("S14")
                    rs.Fields("inPackBeans").Value = .Range("S12")
                    rs.Fields("outPackGround").Value = .Range("Y14")
                    rs.Fields("outPackBeans").Value = .Range("Y12")
                    rs.Fields("pwGround").Value = .Range("AE14")
                    rs.Fields("pwBeans").Value = .Range("AE12")
                    rs.Fields("gpValueLoss").Value = ThisWorkbook.Sheets("Operations sequence").Range("AM3")
                    rs.Fields("rework").Value = ThisWorkbook.Sheets("Operations sequence").Range("AN3")
                    rs.Fields("reworkAtPackingBeans").Value = reworkAtPackingBeans
                    rs.Fields("reworkAtPackingGround").Value = reworkAtPackingGround
                    rs.Fields("bmId").Value = bmId
                    rs.Update
                End If
                rs.Close
                Set rs = Nothing
                saveOverview CInt(bmId), bk
                If saved Then MsgBox "Data saved successfully!", vbOKOnly + vbInformation, "Success"
            End With
        End If
    End If
End If

exit_here:
closeConnection
Set rs = Nothing
Set conn = Nothing
Exit Sub

err_trap:
MsgBox "Error no " & Err.Number & ", description: " & Err.Description
Resume exit_here

End Sub

Public Sub CreateOverview(control As IRibbonControl)
Dim bk As clsBlendKeeper

Set bk = New clsBlendKeeper
bk.ToOverview

End Sub

Public Sub saveOverview(bmId As Integer, bk As clsBlendKeeper)
Dim b As clsBlend
Dim iSql As String

On Error GoTo err_trap

bk.downloadCost

updateConnection

adoConn.Execute "DELETE FROM tbBMOverview WHERE bmId=" & bmId

For Each b In bk.blends
    iSql = "INSERT INTO tbBMOverview (bmId, zfinId, roastingIn, roastingOut, packingIn, packingOut, warehouseIn,  bomVsRealScrap, gpValueLoss, isConsistent, rework, reworkAtPacking) VALUES ("
    iSql = iSql & bmId & "," & zfins(CStr(b.index)).zfinId & "," & Replace(CStr(b.getGreen), ",", ".") & "," & Replace(CStr(b.getRoasted), ",", ".")
    iSql = iSql & "," & Replace(CStr(b.getGround), ",", ".") & "," & Replace(CStr(b.getPacked), ",", ".") & "," & Replace(CStr(b.getStocked), ",", ".")
    iSql = iSql & "," & Replace(CStr(b.bomVsReal), ",", ".") & "," & Replace(CStr(b.gpValueLoss), ",", ".")
    If b.consistent Then
        iSql = iSql & ",1,"
    Else
        iSql = iSql & ",0,"
    End If
    iSql = iSql & Replace(CStr(b.rework), ",", ".") & "," & Replace(CStr(b.reworkAtPacking), ",", ".") & ")"
    adoConn.Execute iSql
Next b

exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""saveOverview"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Private Sub downloadZfins(typeStr As String)
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newZfin As clsZfin

On Error GoTo err_trap

n = zfins.Count
Do While zfins.Count > 0
    zfins.Remove n
    n = n - 1
Loop

sSql = "SELECT z.zfinId, zfinIndex, zp.[beans?] as IsBeans FROM tbZfin z LEFT JOIN tbZfinProperties zp ON z.zfinId=zp.zfinId WHERE zfinType IN (" & typeStr & ");"
Set rs = New ADODB.Recordset
rs.Open sSql, adoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newZfin = New clsZfin
        With newZfin
            .zfinId = rs.Fields("zfinId").Value
            .zfinIndex = rs.Fields("zfinIndex").Value
            If rs.Fields("IsBeans").Value = 1 Then
                .IsBeans = True
            Else
                .IsBeans = False
            End If
            zfins.Add newZfin, CStr(rs.Fields("zfinIndex").Value)
        End With
        rs.MoveNext
    Loop
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadZfins. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub

Public Sub bringWeek(Optional control As IRibbonControl)
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim SQL As String
Dim bool As Boolean

On Error GoTo err_trap

initializeObjects

If w > 0 Then
    clear
    Set conn = New ADODB.Connection
    conn.Open ConnectionString
    conn.CommandTimeout = 90
    ThisWorkbook.Sheets("BM").Unprotect
    bool = False
    If w > 1 Then
        SQL = "SELECT bmEndState, leftover FROM tbBMDetails LEFT JOIN tbBM ON tbBM.bmId = tbBMDetails.bmId WHERE tbBM.bmWeek = " & w - 1 & " AND tbBM.bmYear = " & y
    Else
        SQL = "SELECT bmEndState, leftover FROM tbBMDetails LEFT JOIN tbBM ON tbBM.bmId = tbBMDetails.bmId WHERE tbBM.bmYear = " & y + 1 & " ORDER BY tbBM.bmWeek DESC"
    End If
    Set rs = conn.Execute(SQL)
    If Not rs.EOF Then
        rs.MoveFirst
        ThisWorkbook.Sheets("BM").Range("B2") = rs.Fields("bmEndState").Value
        ThisWorkbook.Sheets("BM").Range("K43") = rs.Fields("leftover").Value
    Else
        bool = True
    End If
    rs.Close
    Set rs = Nothing
    
    SQL = "SELECT * FROM tbBM WHERE bmWeek = " & w & " AND bmYear = " & y
    Set rs = conn.Execute(SQL)
    If rs.EOF Then
        MsgBox "Brak danych dla wybranego okresu!", vbOKOnly + vbExclamation, "Brak danych"
        rs.Close
        updateProperty "weekLoaded", 0
        updateProperty "yearLoaded", 0
        ThisWorkbook.Sheets("BM").Range("G1").Value = "TYDZIEŃ -|-"
        Application.Caption = "Period has not been loaded"
    Else
        updateProperty "roastingFrom", rs.Fields("roastingFrom")
        updateProperty "roastingTo", rs.Fields("roastingTo")
        updateProperty "grindingFrom", rs.Fields("grindingFrom")
        updateProperty "grindingTo", rs.Fields("grindingTo")
        updateProperty "packingFrom", rs.Fields("packingFrom")
        updateProperty "packingTo", rs.Fields("packingTo")
        With ThisWorkbook.Sheets("BM")
            .Range("B42") = "Tydzień " & w
            .Range("C44") = rs.Fields("roastingFrom").Value
            .Range("D44") = rs.Fields("roastingTo").Value
            .Range("C45") = rs.Fields("grindingFrom").Value
            .Range("D45") = rs.Fields("grindingTo").Value
            .Range("C46") = rs.Fields("packingFrom").Value
            .Range("D46") = rs.Fields("packingTo").Value
        End With
        rs.Close
        SQL = "SELECT * FROM tbBMDetails JOIN tbBM ON tbBM.bmId = tbBMDetails.bmId WHERE tbBM.bmWeek = " & w & " AND tbBM.bmYear = " & y
        Set rs = conn.Execute(SQL)
        If Not rs.EOF Then
            With ThisWorkbook.Sheets("BM")
                If bool Then .Range("b2") = rs.Fields("bmStartState").Value
                 .Range("B6") = rs.Fields("greenCoffeeReceipt").Value
'                .Range("B10") = rs.Fields("inOutDifference").Value
                .Range("F4") = rs.Fields("inRoastBeans").Value
                .Range("F6") = rs.Fields("inRoastGround").Value
                .Range("O5") = rs.Fields("outRoastBeans").Value
                .Range("b14") = rs.Fields("bmEndState").Value
                .Range("O11") = rs.Fields("outRoastGround").Value
                .Range("D23") = rs.Fields("zforMovBeans").Value
                .Range("D29") = rs.Fields("zforMoveGround").Value
                .Range("J23") = rs.Fields("outPackBeans").Value
                .Range("J29") = rs.Fields("outGroundCoffee").Value
                .Range("H45") = rs.Fields("pw").Value
                .Range("H46") = rs.Fields("pwMes").Value
                .Range("H48") = rs.Fields("PreviousPw").Value
                If bool Then .Range("K43") = rs.Fields("previousLeftover").Value
                .Range("K44") = rs.Fields("leftover").Value
                .Range("K45") = rs.Fields("addedZfin").Value
                .Range("F13") = rs.Fields("receiptLoss").Value
                .Range("F14") = rs.Fields("cleaningLoss").Value
                .Range("F15") = rs.Fields("silosLoss").Value
                .Range("F16") = rs.Fields("zLoss").Value
                .Range("J14") = rs.Fields("rocksLoss").Value
                .Range("J15") = rs.Fields("huskLoss").Value
                .Range("J17") = rs.Fields("reworkRoasting").Value
                .Range("M36") = rs.Fields("reworkGrinding").Value
                .Range("M37") = rs.Fields("reworkPiabs").Value
                .Range("M33") = rs.Fields("coffeeGrindingLoss").Value
                .Range("M34") = rs.Fields("coffeePiabsLoss").Value
                .Range("G42") = rs.Fields("zfinOverweight").Value
                .Range("G43") = rs.Fields("trashedCoffee").Value
                .Range("J19") = rs.Fields("roastingMess").Value
                .Range("J33") = rs.Fields("grindingMess").Value
                .Range("L4") = rs.Fields("rn3000").Value
                .Range("L5") = rs.Fields("finezja").Value
            End With
'            updateProperty "weekLoaded", w
'            updateProperty "yearLoaded", y
'            ThisWorkbook.Sheets("BM").Range("G1").value = "TYDZIEŃ " & w & "|" & y
'            updateProperty "import path", "K:\Common\Bilans Mas\Tygodniowo\" & ThisWorkbook.CustomDocumentProperties("yearLoaded").value & "\Tydzień " & ThisWorkbook.CustomDocumentProperties("weekLoaded").value
'            rib.InvalidateControl "btnSave"
'            highlightEmpty
        Else
            clear
        End If
        updateProperty "weekLoaded", w
        updateProperty "yearLoaded", y
        ThisWorkbook.Sheets("BM").Unprotect
        ThisWorkbook.Sheets("BM").Range("G1").Value = "TYDZIEŃ " & w & "|" & y
        Application.Caption = "Loaded week " & w & "|" & y
        highlightEmpty
        ThisWorkbook.Sheets("BM").Protect
        updateProperty "import path", "K:\Common\Bilans Mas\Tygodniowo\" & ThisWorkbook.CustomDocumentProperties("yearLoaded").Value & "\Tydzień " & ThisWorkbook.CustomDocumentProperties("weekLoaded").Value
        rs.Close
        Set rs = Nothing
    End If
End If


exit_here:
ThisWorkbook.Sheets("BM").Protect
Set rs = Nothing
Set conn = Nothing
Exit Sub

err_trap:
MsgBox "Error no " & Err.Number & ", description: " & Err.Description
Resume exit_here

End Sub

Public Function IsoWeekNumber(InDate As Date) As Long
    IsoWeekNumber = DatePart("ww", InDate, vbMonday, vbFirstFourDays)
End Function

Sub updateCharts(control As IRibbonControl)
graphSettings.Show
End Sub

Public Sub bringResults(SQL As String)
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Dim i As Integer
Dim period As String
Dim dFrom As Date
Dim dTo As Date
Dim total As Double
Dim green As Double
Dim roastLossG As Double
Dim roastLossB As Double
Dim grindLoss As Double
Dim packLossG As Double
Dim packLossB As Double
Dim rework As Double
Dim reworkAtPackingBeans As Double
Dim reworkAtPackingGround As Double

On Error GoTo err_trap
initializeObjects

updateConnection

clearResults

Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
    MsgBox "There is no data for chosen period!", vbOKOnly + vbExclamation, "No data"
    rs.Close
Else
'    rs.Sort = "ORDER BY tbbm.bmYear ASC, tbBM.bmWeek ASC"
    rs.MoveLast
    Do Until rs.BOF
        green = rs.Fields("receiptLoss").Value + rs.Fields("cleaningLoss").Value + rs.Fields("zLoss").Value
        total = rs.Fields("inRoastBeans").Value + rs.Fields("inRoastGround") + Abs(green) + rs.Fields("rework") + rs.Fields("reworkAtPackingBeans") + rs.Fields("reworkAtPackingGround")
        For i = 65 To 1000
            period = ThisWorkbook.Sheets("Results").Range("A" & i)
            If period = "" Then
                'save results
                With ThisWorkbook.Sheets("Results")
                    If IsNull(rs.Fields("rework")) Then rework = 0 Else rework = rs.Fields("rework")
                    If IsNull(rs.Fields("reworkAtPackingBeans")) Then reworkAtPackingBeans = 0 Else reworkAtPackingBeans = rs.Fields("reworkAtPackingBeans")
                    If IsNull(rs.Fields("reworkAtPackingGround")) Then reworkAtPackingGround = 0 Else reworkAtPackingGround = rs.Fields("reworkAtPackingGround")
                    .Range("A" & i) = rs.Fields("bmYear").Value & "|" & rs.Fields("bmWeek").Value
                    dFrom = rs.Fields("roastingFrom").Value
                    dTo = rs.Fields("roastingTo").Value
                    .Range("B" & i) = Day(dFrom) & "." & month(dFrom) & "-" & Day(dTo) & "." & month(dTo)
                    .Range("C" & i) = Round(rs.Fields("pwGround").Value + rs.Fields("pwBeans").Value, 1)
                    .Range("D" & i) = rs.Fields("receiptLoss").Value
                    .Range("E" & i) = rs.Fields("receiptLoss").Value / total
                    .Range("F" & i) = rs.Fields("cleaningLoss").Value
                    .Range("G" & i) = rs.Fields("cleaningLoss").Value / total
                    .Range("H" & i) = rs.Fields("zLoss").Value
                    .Range("I" & i) = rs.Fields("zLoss").Value / total
                    .Range("J" & i) = green
                    .Range("K" & i) = green / total
                    .Range("L" & i) = Round(rs.Fields("inRoastBeans"), 1)
                    .Range("m" & i) = Round(rs.Fields("inRoastGround"), 1)
                    .Range("n" & i) = Round(rs.Fields("outRoastBeans"), 1)
                    .Range("O" & i) = Round(rs.Fields("outRoastGround"), 1)
                    roastLossB = rs.Fields("inRoastBeans") - rs.Fields("outRoastBeans")
                    roastLossG = rs.Fields("inRoastGround") - rs.Fields("outRoastGround")
                    .Range("p" & i) = roastLossB
                    .Range("Q" & i) = roastLossG
                    .Range("R" & i) = -1 * (roastLossB / rs.Fields("inRoastBeans"))
                    .Range("S" & i) = -1 * (roastLossG / rs.Fields("inRoastGround"))
                    .Range("T" & i) = -1 * ((roastLossB + roastLossG) / (rs.Fields("inRoastBeans") + rs.Fields("inRoastGround")))
                    .Range("U" & i) = Round(-1 * (rs.Fields("outRoastGround") + rework - rs.Fields("inPackGround")), 1)
                    .Range("V" & i) = -1 * ((rs.Fields("outRoastGround") + rework - rs.Fields("inPackGround")) / (rs.Fields("outRoastGround") + rework))
                    .Range("W" & i) = Round(rs.Fields("inPackBeans") + rs.Fields("reworkAtPackingBeans"), 1)
                    .Range("X" & i) = Round(rs.Fields("inPackGround") + rs.Fields("reworkAtPackingGround"), 1)
                    .Range("Y" & i) = Round(rs.Fields("outPackBeans"), 1)
                    .Range("z" & i) = Round(rs.Fields("outPackGround"), 1)
                    .Range("AA" & i) = Round(-1 * (rs.Fields("inPackBeans") - rs.Fields("outPackBeans")), 1)
                    .Range("AB" & i) = Round(-1 * (rs.Fields("inPackGround") - rs.Fields("outPackGround")), 1)
                    .Range("AC" & i) = Round(-1 * ((rs.Fields("inPackBeans") + rs.Fields("reworkAtPackingBeans") - rs.Fields("outPackBeans")) / (rs.Fields("inPackBeans") + rs.Fields("reworkAtPackingBeans"))), 4)
                    .Range("AD" & i) = Round(-1 * ((rs.Fields("inPackGround") + rs.Fields("reworkAtPackingGround") - rs.Fields("outPackGround")) / (rs.Fields("inPackGround") + rs.Fields("reworkAtPackingGround"))), 4)
                    .Range("AE" & i) = Round(-1 * ((rs.Fields("inPackGround") - rs.Fields("outPackGround")) + (rs.Fields("inPackBeans") - rs.Fields("outPackBeans")) + reworkAtPackingGround + reworkAtPackingBeans) / (rs.Fields("inPackGround") + rs.Fields("inPackBeans") + reworkAtPackingBeans + reworkAtPackingGround), 4)
                    .Range("AF" & i) = Round(rs.Fields("pwBeans"), 1)
                    .Range("AG" & i) = Round(rs.Fields("pwGround"), 1)
                    .Range("AH" & i) = -1 * Round(rs.Fields("inRoastBeans") + rs.Fields("reworkAtPackingBeans") - rs.Fields("pwBeans"), 1)
                    .Range("AI" & i) = -1 * Round((rs.Fields("inRoastBeans") + rs.Fields("reworkAtPackingBeans") - rs.Fields("pwBeans")) / (rs.Fields("inRoastBeans") + rs.Fields("reworkAtPackingBeans")), 4)
                    .Range("AJ" & i) = -1 * Round(rs.Fields("inRoastGround") + rework + rs.Fields("reworkAtPackingGround") - rs.Fields("pwGround"), 1)
                    .Range("AK" & i) = -1 * Round((rs.Fields("inRoastGround") + rework + rs.Fields("reworkAtPackingGround") - rs.Fields("pwGround")) / (rs.Fields("inRoastGround") + rework + rs.Fields("reworkAtPackingGround")), 4)
                    .Range("AL" & i) = -1 * Round((total - (rs.Fields("pwGround") + rs.Fields("pwBeans"))), 1)
                    .Range("AM" & i) = -1 * Round((total - (rs.Fields("pwGround") + rs.Fields("pwBeans"))) / total, 4)
                    .Range("AN" & i) = -1 * Round((rs.Fields("outRoastBeans") + rs.Fields("outRoastGround")) - (rs.Fields("pwGround") + rs.Fields("pwBeans")), 1)
                    .Range("AO" & i) = -1 * Round(Round((rs.Fields("outRoastBeans") + rs.Fields("outRoastGround")) - (rs.Fields("pwGround") + rs.Fields("pwBeans")), 1) / Round(rs.Fields("outRoastBeans") + rs.Fields("outRoastGround"), 1), 4)
                    .Range("AP" & i) = rs.Fields("gpValueLoss")
                    .Range("AQ" & i) = rs.Fields("rework")
                End With
                Exit For
            End If
        Next i
        rs.MovePrevious
    Loop
Restart:
    TotalRecords = i
    formatResults
    With ThisWorkbook.Sheets("Results").Range("A63:AQ" & i).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With ThisWorkbook.Sheets("Results").Range("A63:AQ" & i).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With ThisWorkbook.Sheets("Results").Range("A63:AQ" & i).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With ThisWorkbook.Sheets("Results").Range("A63:AQ" & i).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With ThisWorkbook.Sheets("Results").Range("A63:AQ" & i).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    updateChart
End If

exit_here:
Set rs = Nothing
closeConnection
Exit Sub

err_trap:
If Err.Number = 94 Then
    Resume Restart
Else
    MsgBox "Error in ""BringResults"". Error number: " & Err.Number & ", description: " & Err.Description
End If
Resume exit_here

End Sub

Public Sub clearResults()

ThisWorkbook.Sheets("Results").Range("A65:AX265").clear

End Sub

Public Sub formatResults()
With ThisWorkbook.Sheets("Results")
    .Range("E64:E263").NumberFormat = "0.00%"
    .Range("G64:G263").NumberFormat = "0.00%"
    .Range("I64:I263").NumberFormat = "0.00%"
    .Range("K64:K263").NumberFormat = "0.00%"
    .Range("R64:R263").NumberFormat = "0.00%"
    .Range("S64:S263").NumberFormat = "0.00%"
    .Range("T64:T263").NumberFormat = "0.00%"
    .Range("V64:V263").NumberFormat = "0.00%"
    .Range("AC64:AC263").NumberFormat = "0.00%"
    .Range("AD64:AD263").NumberFormat = "0.00%"
    .Range("AE64:AE263").NumberFormat = "0.00%"
    .Range("AI64:AI263").NumberFormat = "0.00%"
    .Range("AK64:AK263").NumberFormat = "0.00%"
    .Range("AM64:AM263").NumberFormat = "0.00%"
    .Range("AO64:AO263").NumberFormat = "0.00%"
    .Range("AM64:AM263").NumberFormat = "0.00%"
End With
End Sub

Public Sub updateChart()
Dim lineChart As ChartObject
Dim srs As Series
'ThisWorkbook.Worksheets("Results").ChartObjects("Wykres 4").Name = "grpGreenCoffee"
For Each lineChart In ThisWorkbook.Worksheets("Results").ChartObjects
    'Debug.Print "Wszystkie serie danych " & lineChart.Name
    With Worksheets("Results")
        Select Case lineChart.Name
        Case Is = "grpRoasting"
            For Each srs In lineChart.Chart.SeriesCollection
                Select Case srs.Name
                Case Is = "Loss on beans"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("R65:R" & TotalRecords)
                Case Is = "Loss on ground"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("S65:S" & TotalRecords)
                Case Is = "Total loss"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("T65:T" & TotalRecords)
                Case Is = "Production BN"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AF65:AF" & TotalRecords)
                Case Is = "Production GD"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AG65:AG" & TotalRecords)
                End Select
            Next srs
        Case Is = "grpGrinding"
            For Each srs In lineChart.Chart.SeriesCollection
                Select Case srs.Name
                Case Is = "Loss on grinding"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("V65:V" & TotalRecords)
                Case Is = "Production volume"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("C65:C" & TotalRecords)
                End Select
            Next srs
        Case Is = "grpPacking"
            For Each srs In lineChart.Chart.SeriesCollection
                Select Case srs.Name
                Case Is = "Loss on beans"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AC65:AC" & TotalRecords)
                Case Is = "Loss on ground"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AD65:AD" & TotalRecords)
                Case Is = "Total loss"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AE65:AE" & TotalRecords)
                Case Is = "Production BN"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AF65:AF" & TotalRecords)
                Case Is = "Production GD"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AG65:AG" & TotalRecords)
                End Select
            Next srs
        Case Is = "grpTotal"
            For Each srs In lineChart.Chart.SeriesCollection
                Select Case srs.Name
                Case Is = "Loss on beans"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AI65:AI" & TotalRecords)
                Case Is = "Loss on ground"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AK65:AK" & TotalRecords)
                Case Is = "Total loss"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AM65:AM" & TotalRecords)
                Case Is = "Production BN"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AF65:AF" & TotalRecords)
                Case Is = "Production GD"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("AG65:AG" & TotalRecords)
                End Select
            Next srs
        Case Is = "grpGreenCoffee"
             For Each srs In lineChart.Chart.SeriesCollection
                Select Case srs.Name
                Case Is = "Production volume"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("C65:C" & TotalRecords)
                Case Is = "Loss on receipt"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("E65:E" & TotalRecords)
                Case Is = "Loss on cleaning"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("G65:G" & TotalRecords)
                Case Is = "Total loss"
                    srs.xValues = .Range("A65:A" & TotalRecords)
                    srs.values = .Range("K65:K" & TotalRecords)
                End Select
            Next srs
        End Select
    End With
Next lineChart

End Sub

Sub yy()
Dim lineChart As ChartObject
Dim srs As Series
'ThisWorkbook.Worksheets("Results").ChartObjects("Wykres 4").Name = "grpGreenCoffee"
For Each lineChart In ThisWorkbook.Worksheets("Results").ChartObjects
    Debug.Print "Wszystkie serie danych " & lineChart.Name
    For Each srs In lineChart.Chart.SeriesCollection
        'Debug.Print srs.Values = worksheets("Results").Range("
    Next srs
Next lineChart
End Sub

Public Sub highlightEmpty()
ThisWorkbook.Sheets("BM").Unprotect
With ThisWorkbook.Sheets("BM")
    With .Range("B6:B7").Borders
        If ThisWorkbook.Sheets("BM").Range("B6").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("B14:B15").Borders
        If ThisWorkbook.Sheets("BM").Range("B14").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("F4").Borders
        If ThisWorkbook.Sheets("BM").Range("F4").Value <> "" Then
            .LineStyle = xlLineStyleNone
            
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("F6").Borders
        If ThisWorkbook.Sheets("BM").Range("F6").Value <> "" Then
            .LineStyle = xlLineStyleNone
            
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("L4").Borders
        If ThisWorkbook.Sheets("BM").Range("L4").Value <> "" Then
            .LineStyle = xlLineStyleNone
            
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("L5").Borders
        If ThisWorkbook.Sheets("BM").Range("L5").Value <> "" Then
            .LineStyle = xlLineStyleNone
            
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("O5:O6").Borders
        If ThisWorkbook.Sheets("BM").Range("O5").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("O11:O12").Borders
        If ThisWorkbook.Sheets("BM").Range("O11").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("F13").Borders
        If ThisWorkbook.Sheets("BM").Range("F13").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("F14").Borders
        If ThisWorkbook.Sheets("BM").Range("F14").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("F16").Borders
        If ThisWorkbook.Sheets("BM").Range("F16").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("J19").Borders
        If ThisWorkbook.Sheets("BM").Range("J19").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("J23:J24").Borders
        If ThisWorkbook.Sheets("BM").Range("J23").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("J29:J30").Borders
        If ThisWorkbook.Sheets("BM").Range("J29").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("J33").Borders
        If ThisWorkbook.Sheets("BM").Range("J33").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("M33").Borders
        If ThisWorkbook.Sheets("BM").Range("M33").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("M34").Borders
        If ThisWorkbook.Sheets("BM").Range("M34").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("M36").Borders
        If ThisWorkbook.Sheets("BM").Range("M36").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("M37").Borders
        If ThisWorkbook.Sheets("BM").Range("M37").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("G42").Borders
        If ThisWorkbook.Sheets("BM").Range("G42").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("G43").Borders
        If ThisWorkbook.Sheets("BM").Range("G43").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("K44").Borders
        If ThisWorkbook.Sheets("BM").Range("K44").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("K45").Borders
        If ThisWorkbook.Sheets("BM").Range("K45").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("H45").Borders
        If ThisWorkbook.Sheets("BM").Range("H45").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("H46").Borders
        If ThisWorkbook.Sheets("BM").Range("H46").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
    With .Range("H48").Borders
        If ThisWorkbook.Sheets("BM").Range("H48").Value <> "" Then
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        Else
            .LineStyle = xlContinuous
            .Weight = 3
            .Color = vbRed
        End If
    End With
End With

ThisWorkbook.Sheets("BM").Protect
End Sub

Sub GetSaveEnabled(control As IRibbonControl, ByRef returnedVal)
'    If weekLoaded = w Then
'        returnedVal = True
'    Else
'        returnedVal = False
'    End If
returnedVal = True
End Sub

Public Sub createCustomProperty(theName As String, theValue As Variant)
Dim theType As Variant

Select Case VarType(theValue)
    Case 0 To 1
    theType = Null
    Case 2 To 3
    theType = msoPropertyTypeNumber
    Case 4 Or 5 Or 14
    theType = msoPropertyTypeFloat
    Case 7
    theType = msoPropertyTypeDate
    Case 8
    theType = msoPropertyTypeString
    Case 11
    theType = msoPropertyTypeBoolean
    Case Else
    theType = Null
End Select

If theType = Null Then
    MsgBox "Type of variable ""theValue"" passed to ""createCustomProperty"" could not be determined or is unsuported. No custom property has been created", vbOKOnly + vbExclamation
Else
    ThisWorkbook.CustomDocumentProperties.Add Name:=theName, LinkToContent:=False, Type:=theType, Value:=theValue
    'MsgBox "Property " & theName & " has been created successfully and set to " & theValue, vbOKOnly + vbInformation


End If

End Sub

Public Function propertyExists(Name As String) As Boolean
Dim prop As DocumentProperty
propertyExists = False
For Each prop In ThisWorkbook.CustomDocumentProperties
    If prop.Name = Name Then
        propertyExists = True
        Exit For
    End If
Next prop
End Function

Public Sub updateProperty(propName As String, propValue As Variant)

With ThisWorkbook.CustomDocumentProperties
    If propertyExists(propName) Then
        .item(propName).Value = propValue
    Else
        createCustomProperty propName, propValue
    End If
End With

End Sub

Public Sub debugCustomProperties()
Dim prop As DocumentProperty
For Each prop In ThisWorkbook.CustomDocumentProperties
    Debug.Print prop.Name
Next prop
End Sub


Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Function FolderExists(strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Public Function filePicker(Optional startDirectory As Variant, Optional title As Variant) As String
Dim f As Object
Dim varFile As Variant
Set f = Application.FileDialog(3)
f.AllowMultiSelect = False
If Not IsMissing(title) Then
    f.title = title
Else
    f.title = "Wybierz plik"
End If
If Not IsMissing(startDirectory) Then
    f.InitialFileName = startDirectory & "\"
End If

With f
    .Filters.clear
    .Filters.Add "Excel files", "*.xls; *.xlsx; *.xlsm"
    .Filters.Add "All Files", "*.*"
   ' Show the dialog box. If the .Show method returns True, the '
   ' user picked at least one file. If the .Show method returns '
   ' False, the user clicked Cancel. '
   If .Show = True Then

      'Loop through each file selected and add it to our list box. '
        filePicker = f.SelectedItems(1)
   End If
End With



End Function

Public Function folderPicker(Optional title As Variant) As String
Dim f As Object
Dim varFile As Variant
Set f = Application.FileDialog(msoFileDialogFolderPicker)
f.AllowMultiSelect = False
If Not IsMissing(title) Then
    f.title = title
Else
    f.title = "Wybierz folder"
End If

With f
    .Filters.clear
    
   ' Show the dialog box. If the .Show method returns True, the '
   ' user picked at least one file. If the .Show method returns '
   ' False, the user clicked Cancel. '
   If .Show = True Then

      'Loop through each file selected and add it to our list box. '
        folderPicker = f.SelectedItems(1)
   End If
End With



End Function

Public Sub initializeObjects()
If w Is Nothing Then
    Set w = ThisWorkbook.CustomDocumentProperties("week")
End If
If y Is Nothing Then
    Set y = ThisWorkbook.CustomDocumentProperties("year")
End If
If m Is Nothing Then
    Set m = ThisWorkbook.CustomDocumentProperties("month")
End If
End Sub

Public Function getDate(ctrDate As String, ctrHour As String) As Date
Dim v() As String
Dim h As Integer
Dim m As Integer
v() = Split(ctrHour, ":")
h = CInt(v(0))
m = CInt(v(1))
getDate = DateAdd("n", (h * 60) + m, CDate(ctrDate))

End Function

Public Function getTime(fullDate As Date) As Variant
Dim v() As String

On Error GoTo err_trap

v() = Split(CStr(fullDate), " ")
getTime = Left(v(1), 5)

exit_here:
Exit Function

err_trap:
getTime = Null
Resume exit_here

End Function

Public Sub importScada(control As IRibbonControl)

If ThisWorkbook.CustomDocumentProperties("weekLoaded") <> 0 And ThisWorkbook.CustomDocumentProperties("yearLoaded") <> 0 Then
    importFromScada ThisWorkbook.CustomDocumentProperties("roastingFrom"), ThisWorkbook.CustomDocumentProperties("roastingTo")
Else
    MsgBox "No week/year has been selected. Please reload data for desired week/year first", vbOKOnly + vbExclamation, "Period could not be determined"
End If

End Sub


Public Function toBeans(blends As Variant) As Variant
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim prodStr As String
Dim i As Integer
Dim dbPath As String

On Error GoTo exit_here

Set conn = New ADODB.Connection
conn.Open ConnectionString
conn.CommandTimeout = 90

If IsArray(blends) Then
    If Not isArrayEmpty(blends) Then
        For i = LBound(blends) To UBound(blends)
            If i = UBound(blends) Then
                prodStr = prodStr & blends(i)
            Else
                prodStr = prodStr & blends(i) & ", "
            End If
        Next i
        
        sSql = "SELECT [beans?] FROM tbZfinProperties JOIN tbZfin on tbZfin.zfinId = tbZfinProperties.zfinId WHERE tbZfin.zfinIndex IN (" & prodStr & ");"
        Set rs = New ADODB.Recordset
        rs.Open sSql, conn, adOpenStatic, adLockOptimistic, adCmdText
        
        Set toBeans = rs
        rs.Close
    End If
End If

exit_here:
Set rs = Nothing
conn.Close
Set conn = Nothing
End Function

Public Function range2array(rng As Range) As Variant
Dim c As Range
Dim var() As Variant

For Each c In rng
    If isArrayEmpty(var) Then
        ReDim var(0) As Variant
        var(0) = c.Value
    Else
        ReDim Preserve var(UBound(var) + 1) As Variant
        var(UBound(var)) = c.Value
    End If
Next c

range2array = var

End Function

Public Function importExcelData(filePath As String, Optional sheetName As Variant, Optional im As Variant) As Variant
Dim cnn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim conStr As String
Dim strSQL  As String
Dim allSheets As Variant
Dim Worksheet As String

On Error GoTo err_trap

If Right(filePath, 1) = "s" Then
    ver = "8"
Else
    ver = "12"
End If


If IsMissing(sheetName) Then
    allSheets = getExcelSheetName(filePath)
    
    If Not IsNull(allSheets) Then
        If UBound(allSheets) > 0 Then MsgBox "There's more than 1 worksheet in the source file. As it's not set which worksheet contains data, the first one was chosen (""" & allSheets(0) & """). In case there's no data imported or you suspect errors, please remove from the source file all sheets but the desired one and try again", vbInformation + vbOKOnly, "Possible errors"
        Worksheet = allSheets(0)
    End If
Else
    Worksheet = sheetName & "$"
End If

If Not Worksheet = "" Then
    If Not IsMissing(im) Then
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & filePath & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=" & im & ";"";"
    Else
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & filePath & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=0;"";"
    End If
    strSQL = "SELECT * FROM [" & Worksheet & "];"
    cnn.Open conStr
    rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
    Set importExcelData = rs
Else
    MsgBox "There's problem reading worksheet's name " & filePath & ". Check ""importExcelData"" function"
    importExcelData = Null
End If



exit_here:
    Set rs = Nothing
    Exit Function
   
err_trap:
    MsgBox "Error in Sub ""importExcelData""." & vbNewLine & "Err no. " & Err.Number & ", description: " & Err.Description
    Resume exit_here
End Function

Public Function getExcelSheetName(path As String) As Variant
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim dollarInt As Long
Dim lengthInt As Long
Dim ver As Integer
Dim conStr As String
Dim arr() As String

If Right(path, 1) = "s" Then
    ver = "8"
Else
    ver = "12"
End If

conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=YES;IMEX=1;"";"

Set cn = New ADODB.Connection
cn.Open conStr
           
Set rs = cn.OpenSchema(20)

If Not rs.EOF Then
    rs.MoveFirst
    If InStr(1, rs.Fields("TABLE_NAME"), "$", vbTextCompare) >= 0 Then
        ReDim arr(0) As String
        arr(0) = rs.Fields("TABLE_NAME").Value
    ElseIf rs.Fields("TABLE_NAME") = "'$'" Then
        ReDim arr(0) As String
        arr(UBound(arr)) = rs.Fields("TABLE_NAME").Value
    End If
    rs.MoveNext
    Do Until rs.EOF
        If InStr(1, rs.Fields("TABLE_NAME"), "$", vbTextCompare) > 0 Then
            If isArrayEmpty(arr) Then
                ReDim arr(0) As String
                arr(UBound(arr)) = rs.Fields("TABLE_NAME").Value
            Else
                ReDim Preserve arr(UBound(arr) + 1) As String
                arr(UBound(arr)) = rs.Fields("TABLE_NAME").Value
            End If
        ElseIf rs.Fields("TABLE_NAME") = "'$'" Then
            If isArrayEmpty(arr) Then
                ReDim arr(0) As String
                arr(UBound(arr)) = rs.Fields("TABLE_NAME").Value
            Else
                ReDim Preserve arr(UBound(arr) + 1) As String
                arr(UBound(arr)) = rs.Fields("TABLE_NAME").Value
            End If
        End If
        rs.MoveNext
    Loop
End If

If isArrayEmpty(arr) Then
    getExcelSheetName = Null
Else
    getExcelSheetName = arr
End If
'While Not rs.EOF
'
'    dollarInt = InStr(rs.Fields("TABLE_NAME").value, "$")
'    lengthInt = Len(rs.Fields("TABLE_NAME").value)
'    If (dollarInt + 1) = lengthInt Then
'       List1.AddItem rs.Fields("TABLE_NAME").value
'    End If
'    rs.MoveNext
'Wend

rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing
End Function

Public Function authorize(fun As Integer, Optional user As Variant) As Boolean
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim temp As Boolean
Dim SQL As String
Dim funString As String
Dim i As Integer

On Error GoTo err_trap

updateConnection

If IsMissing(user) Then
    user = ThisWorkbook.CustomDocumentProperties("userId")
End If

temp = False

SQL = "SELECT functionId FROM tbPrivilages WHERE userId = " & user

Set rs = New ADODB.Recordset

rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If rs.Fields("functionId") = fun Then
            bool = True
            Exit Do
        End If
        rs.MoveNext
    Loop
End If

rs.Close

If bool Then
    authorize = True
Else
    SQL = "SELECT functionString FROM tbFunctions WHERE functionId = " & fun
    Set rs = New ADODB.Recordset
    rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveFirst
        funString = rs.Fields("functionString")
    End If
    rs.Close
    MsgBox "You are not authorized to use this function. Missing authorization: " & funString, vbExclamation + vbOKOnly, "Access denied"
    authorize = False
End If

exit_here:
If Not rs Is Nothing Then Set rs = Nothing
closeConnection
Exit Function

err_trap:
MsgBox "Error in ""Authorize"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Sub editConnections(control As IRibbonControl)
editConnectionsForm.Show
End Sub

Public Function inCollection(ind As String, col As Collection) As Boolean
Dim v As Variant
Dim isError As Boolean

isError = False

On Error GoTo err_trap

Set v = col(ind)

exit_here:
If isError Then
    inCollection = False
Else
    inCollection = True
End If
Exit Function

err_trap:
isError = True
Resume exit_here


End Function

Public Sub checkMissingConnections()
Dim sht As Worksheet
Dim rng As Range
Dim c As Range
Dim consistent As Boolean
Dim x As Integer
Dim y As Integer
Dim index As Long
Dim zfors As New Collection
Dim zfins As New Collection
Dim ord As clsOrder
Dim zfinStr As String
Dim zforStr As String
Dim SQL As String
Dim rs As ADODB.Recordset
Dim missingZfors As String
Dim missingZfins As String
Dim mZfins As String
Dim mZfors As String
Dim Msg As String

Set sht = ThisWorkbook.Sheets("Operations sequence")

Set rng = Application.Selection

consistent = True
index = 0

For Each c In rng
    If c.Column <> 1 Then
        consistent = False
        Exit For
    Else
        If index = 0 Then index = c.Value
        If c.Value <> index And (c.Value <> Empty) Then
            consistent = False
            Exit For
        End If
    End If
Next c

If consistent Then
    For Each c In rng
        If sht.Range("C" & c.row).Value <> Empty Then
            index = sht.Range("C" & c.row).Value
            If Not inCollection(CStr(index), zfors) Then
                Set ord = New clsOrder
                ord.sapId = index
                zfors.Add ord, CStr(index)
                zforStr = zforStr & index & ","
            End If
        End If
        If sht.Range("O" & c.row).Value <> Empty Then
            index = sht.Range("O" & c.row).Value
            If Not inCollection(CStr(index), zfins) Then
                Set ord = New clsOrder
                ord.sapId = index
                zfins.Add ord, CStr(index)
                zfinStr = zfinStr & index & ","
            End If
        End If
    Next c
    If Len(zfinStr) > 0 Then zfinStr = Left(zfinStr, Len(zfinStr) - 1)
    If Len(zforStr) > 0 Then zforStr = Left(zforStr, Len(zforStr) - 1)
    
    missingZfins = checkMissingZfins(zforStr, zfins)
    missingZfors = checkMissingZfors(zfinStr, zfors)
    
    If Len(missingZfins) = 4 And Len(missingZfors) = 4 Then
        MsgBox "No ZFIN nor ZFOR order is missing. Everything is OK", vbInformation + vbOKOnly, "Orders OK"
    Else
        Do Until Len(missingZfins) <= 4 And Len(missingZfors) <= 4
            'go deeper and deeper to find what will be missing if we add missing orders from previous step
            If Len(missingZfins) > 4 Then
                If Len(mZfins) > 0 Then
                    mZfins = mZfins & "," & missingZfins
                Else
                    mZfins = missingZfins
                End If
                missingZfins = ""
                zfinStr = mZfins
                missingZfors = checkMissingZfors(zfinStr, zfors)
                missingZfors = validateMissing(missingZfors, mZfors)
            End If
            If Len(missingZfors) > 4 Then
                If Len(mZfors) > 0 Then
                    mZfors = mZfors & "," & missingZfors
                Else
                    mZfors = missingZfors
                End If
                mZfors = mZfors & "," & missingZfors
                missingZfors = ""
                zforStr = mZfors
                missingZfins = checkMissingZfins(zforStr, zfins)
                missingZfins = validateMissing(missingZfins, mZfins)
            End If
            If Len(missingZfins) > 4 Then
                If Len(mZfins) = 0 Then
                    mZfins = missingZfins
                Else
                    mZfins = mZfins & "," & missingZfins
                End If
            End If
            If Len(missingZfors) > 4 Then
                If Len(mZfors) = 0 Then
                    mZfors = missingZfors
                Else
                    mZfors = mZfors & "," & missingZfors
                End If
            End If
        Loop
        MsgBox "Missing ZFOR orders: " & mZfors & vbNewLine & "Missing ZFIN orders: " & mZfins
    End If
Else
    MsgBox "Selected range is inconsistent (selected are more than one blends) or  there's no blend number in selection. Please select only blend number of single blend (e.g. 34005471)", vbExclamation + vbOKOnly, "Inconsistent selection"
End If

End Sub

Public Sub MissingConnectionsBySessionNumber()
Dim sht As Worksheet
Dim rng As Range
Dim consistent As Boolean
Dim rs As ADODB.Recordset
Dim SQL As String
Dim inp As String
Dim index As Long
Dim nFile As Workbook
Dim nSht As Worksheet
Dim i As Integer
Dim DepStr As String

On Error GoTo err_trap

updateConnection

Set sht = ThisWorkbook.Sheets("Operations sequence")

Set rng = Application.Selection

consistent = True
index = 0

For Each c In rng
    If c.Column <> 1 Then
        consistent = False
        Exit For
    Else
        If index = 0 Then index = c.Value
        If c.Value <> index And (c.Value <> Empty) Then
            consistent = False
            Exit For
        End If
    End If
Next c

If consistent Then
    inp = InputBox("Provide MES session number", "Parameter needed")
    If CInt(inp) > 0 Then
        SQL = "SELECT DISTINCT oZfor.sapId as [Zfor order], oZfin.sapId as [Zfin order], op2.SessionNumber, od.isRemoved " _
        & "FROM tbZfin z LEFT JOIN tbOperations op ON op.zfinId=z.zfinId " _
        & "LEFT JOIN tbOrders oZfor ON oZfor.orderId=op.orderId " _
        & "LEFT JOIN tbOrderDep od ON od.zforOrder=oZfor.orderId " _
        & "LEFT JOIN tbOrders oZfin on od.zfinOrder=oZfin.orderId " _
        & "LEFT JOIN tbOperations op2 On op2.orderId=oZfin.orderId " _
        & "WHERE op.SessionNumber = " & CInt(inp) & " And z.zfinIndex = " & index

        Set rs = New ADODB.Recordset
        rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Set nFile = Workbooks.Add
            Set nSht = nFile.Worksheets.Add
            nSht.Name = "List of ZFOR orders"
            nSht.Range("A1") = "Zfor order"
            nSht.Range("B1") = "Zfin order"
            nSht.Range("C1") = "Session number"
            nSht.Range("D1") = "Is removed?"
            i = 2
            Do Until rs.EOF
                nSht.Range("A" & i) = rs.Fields("Zfor order").Value
                nSht.Range("B" & i) = rs.Fields("Zfin order").Value
                nSht.Range("C" & i) = rs.Fields("SessionNumber").Value
                nSht.Range("D" & i) = rs.Fields("IsRemoved").Value
                DepStr = DepStr & rs.Fields("Zfin order").Value & ","
                i = i + 1
                rs.MoveNext
            Loop
            If Len(DepStr) > 0 Then
                DepStr = Left(DepStr, Len(DepStr) - 1)
                SQL = "SELECT DISTINCT oZfin.sapId as [Zfin order], oZfor.sapId as [Zfor order], op.SessionNumber, od.isRemoved " _
                    & "FROM tbOrders oZfin LEFT JOIN tbOrderDep od ON od.zfinOrder=oZfin.orderId " _
                    & "LEFT JOIN tbOrders oZfor ON od.zforOrder=oZfor.orderId " _
                    & "LEFT JOIN tbOperations op ON op.orderId=oZfor.orderId " _
                    & "WHERE oZfin.sapId IN (" & DepStr & ")"
                Set rs = New ADODB.Recordset
                rs.Open SQL, adoConn, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    Set nSht = nFile.Worksheets.Add
                    nSht.Name = "List of ZFIN orders"
                    nSht.Range("A1") = "Zfin order"
                    nSht.Range("B1") = "Zfor order"
                    nSht.Range("C1") = "Session number"
                    nSht.Range("D1") = "Is removed?"
                    i = 2
                    Do Until rs.EOF
                        nSht.Range("A" & i) = rs.Fields("Zfin order").Value
                        nSht.Range("B" & i) = rs.Fields("Zfor order").Value
                        nSht.Range("C" & i) = rs.Fields("SessionNumber").Value
                        nSht.Range("D" & i) = rs.Fields("IsRemoved").Value
                        i = i + 1
                        rs.MoveNext
                    Loop
                End If
            End If
        End If
    End If
    
Else
    MsgBox "Selected range is inconsistent (selected are more than one blends) or  there's no blend number in selection. Please select only blend number of single blend (e.g. 34005471)", vbExclamation + vbOKOnly, "Inconsistent selection"
End If

exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Error in MissingConnectionsBySessionNumber. Description: " & Err.Description
Resume exit_here

End Sub

Public Function validateMissing(newStr As String, exStr As String) As String
'checks witch orders from newStr are already in exStr. Output is string of newStr not existent in exStr
Dim v() As String
Dim i As Integer
Dim nStr As String

If Len(newStr) > 4 Then
    v = Split(newStr, ",", , vbTextCompare)
    For i = LBound(v) To UBound(v)
        If InStr(1, exStr, Trim(v(i)), vbTextCompare) = 0 Then
            If Len(nStr) > 0 Then
                nStr = nStr & "," & Trim(v(i))
            Else
                nStr = Trim(v(i))
            End If
        End If
    Next i
    validateMissing = nStr
Else
    validateMissing = ""
End If

End Function

Public Function checkMissingZfins(zforStr As String, zfins As Collection) As String
'checks missing zfin order data based on provided zfor order sting
Dim missingZfins As String

updateConnection
SQL = "SELECT DISTINCT oZfin.sapId " _
        & "FROM tbOrders oZfor LEFT JOIN tbOrderDep od ON od.zforOrder=oZfor.orderId LEFT JOIN tbOrders oZfin ON oZfin.orderId=od.zfinOrder " _
        & "WHERE oZfor.sapId IN (" & zforStr & ")"
Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If Not inCollection(CStr(rs.Fields("sapId")), zfins) Then
            missingZfins = missingZfins & rs.Fields("sapId") & ", "
        End If
        rs.MoveNext
    Loop
    If Len(missingZfins) > 0 Then missingZfins = Left(missingZfins, Len(missingZfins) - 2) Else missingZfins = "none"
End If
rs.Close
Set rs = Nothing

checkMissingZfins = missingZfins

End Function

Public Function checkMissingZfors(zfinStr As String, zfors As Collection) As String
'checks missing zfin order data based on provided zfor order sting
Dim missingZfors As String

updateConnection
SQL = "SELECT DISTINCT oZfor.sapId " _
    & "FROM tbOrders oZfin LEFT JOIN tbOrderDep od ON od.zfinOrder=oZfin.orderId LEFT JOIN tbOrders oZfor ON oZfor.orderId=od.zforOrder " _
    & "WHERE oZfin.sapId IN (" & zfinStr & ") AND (od.isRemoved IS NULL OR od.isRemoved = 0)"
Set rs = New ADODB.Recordset
rs.Open SQL, adoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If Not inCollection(CStr(rs.Fields("sapId")), zfors) Then
            missingZfors = missingZfors & rs.Fields("sapId") & ", "
        End If
        rs.MoveNext
    Loop
    If Len(missingZfors) > 0 Then missingZfors = Left(missingZfors, Len(missingZfors) - 2) Else missingZfors = "none"
End If
rs.Close
Set rs = Nothing

checkMissingZfors = missingZfors

End Function

Public Sub showAllColumns(control As IRibbonControl)
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("Operations sequence")

sht.Range("C:AM").EntireColumn.Hidden = False

End Sub

Public Sub showRoastingColumns(control As IRibbonControl)
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("Operations sequence")

sht.Range("C:AM").EntireColumn.Hidden = True
sht.Range("C:H,AA:AA,AD:AD,AG:AJ").EntireColumn.Hidden = False

End Sub

Public Sub showGrindingColumns(control As IRibbonControl)
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("Operations sequence")

sht.Range("C:AM").EntireColumn.Hidden = True
sht.Range("E:F,I:L,Y:Z,AA:AA,AD:AD").EntireColumn.Hidden = False

End Sub

Public Sub showPackingColumns(control As IRibbonControl)
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("Operations sequence")

sht.Range("C:AM").EntireColumn.Hidden = True
sht.Range("I:J,M:V,Y:Z,AB:AB,AE:AE").EntireColumn.Hidden = False

End Sub

Public Sub showLossesColumns(control As IRibbonControl)
Dim sht As Worksheet

Set sht = ThisWorkbook.Sheets("Operations sequence")

sht.Range("C:AM").EntireColumn.Hidden = True
sht.Range("G:H,K:L,M:N,Q:R,W:AF,AM:AM").EntireColumn.Hidden = False

sht.Range("G1") = "Roasting Loss"
sht.Range("K1") = "Grinding Loss"
sht.Range("Q1") = "Packing Loss"

End Sub

Public Function IsInArray(valueToBeFound As Variant, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valueToBeFound Or arr(i) = CStr(valueToBeFound) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Public Sub showDbFunctions(control As IRibbonControl)
dbFunctions.Show
End Sub

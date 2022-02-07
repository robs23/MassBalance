Attribute VB_Name = "Blendy"
Public Sub formatBlends()

With ThisWorkbook.Sheets("Roasting history")
    .Cells.clear
    .Range("B1:D1").Merge
    .Range("E1:G1").Merge
    .Range("H1:J1").Merge
    .Range("A1:A2").Merge
    .Range("A1") = "Period"
    .Range("B1") = "RN3000"
    .Range("E1") = "RN4000"
    .Range("H1") = "Total"
    .Range("B2") = "Green [kg]"
    .Range("E2") = "Green [kg]"
    .Range("H2") = "Green [kg]"
    .Range("C2") = "Roasted [kg]"
    .Range("F2") = "Roasted [kg]"
    .Range("I2") = "Roasted [kg]"
    .Range("D2") = "Loss [%]"
    .Range("G2") = "Loss [%]"
    .Range("J2") = "Loss [%]"
    .Range("A1:J2").Font.Bold = True
    .Range("A1:J2").HorizontalAlignment = xlCenter
End With

End Sub

Public Sub updateRoastingHistory(control As IRibbonControl)
roastingHistory.Show
End Sub

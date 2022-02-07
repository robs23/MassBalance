Attribute VB_Name = "Meser"
Public zlecenia As New Collection

Public Sub importMes(control As IRibbonControl)

With ThisWorkbook.Sheets("MES")
    .Cells.clear
    .Range("B1") = "Roasting"
    .Range("B1").Font.Bold = True
    .Range("B1").Font.Size = 18
    .Range("H1") = "Grinding"
    .Range("H1").Font.Bold = True
    .Range("H1").Font.Size = 18
    .Range("N1") = "Packing"
    .Range("N1").Font.Bold = True
    .Range("N1").Font.Size = 18
End With

roastingMes
grindingMes
packingMes
formatMES
End Sub

Public Sub grindingMes()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim oCol As String
Dim pCol As String
Dim nCol As String
Dim aCol As String
Dim mCol As String
Dim kgCol As String
Dim theSum As Double
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim prop As String
'Dim zlec As clsZlecenie
Dim ver As String

prop = "grinding mes file"

If propertyExists(prop) And ThisWorkbook.CustomDocumentProperties(prop) <> "" Then
    If ThisWorkbook.CustomDocumentProperties("import path") <> "" Then
        If FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls"
        ElseIf FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx"
        Else
            MsgBox "Source file """ & ThisWorkbook.CustomDocumentProperties(prop) & """ could not be found in " & ThisWorkbook.CustomDocumentProperties("import path") & "\ . Check in settings if both file name and path are correct.", vbOKOnly + vbExclamation, "Error"
        End If
        
        If path <> "" Then
            Set rs = importExcelData(path, "Zestawienie ilości wyprodukowan", 1)
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If Len(oCol) = 0 And Len(pCol) = 0 And Len(nCol) = 0 And Len(aCol) = 0 And Len(mCol) = 0 And Len(kgCol) = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            Select Case rs.Fields(i)
                            Case Is = "Nr zlecenia"
                                oCol = rs.Fields(i).Name
                            Case Is = "Nr produktu"
                                pCol = rs.Fields(i).Name
                            Case Is = "Nazwa produktu"
                                nCol = rs.Fields(i).Name
                            Case Is = "Ilość"
                                aCol = rs.Fields(i).Name
                            Case Is = "Nr maszyny"
                                mCol = rs.Fields(i).Name
                            Case Is = "Ilość kg"
                                kgCol = rs.Fields(i).Name
                            End Select
                        Next i
                    Else
                        rs.Close
                        Set rs = Nothing
                        Exit Do
                    End If
                    rs.MoveNext
                Loop
                If Right(path, 1) = "s" Then
                    ver = "8"
                Else
                    ver = "12"
                End If
                Worksheet = "Zestawienie ilości wyprodukowan$"
                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=0;"";"
                'by order number
                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sub." & nCol & ", sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY " & oCol & ", " & pCol & ", " & nCol & ";"
                'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
                'strSQL = "SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100;"
                cnn.Open conStr
                Set rs = New ADODB.Recordset
                rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs.EOF Then
                    rs.MoveFirst
                    With ThisWorkbook.Sheets("MES")
                        u = 2
                        .Range("G2") = "Order number"
                        .Range("H2") = "ZFOR"
                        .Range("I2") = "Description"
                        .Range("J2") = "Amount [kg]"
                        Do Until rs.EOF
                            u = u + 1
                            .Range("G" & u) = rs.Fields(oCol).Value
                            .Range("H" & u) = rs.Fields(pCol).Value
                            .Range("I" & u) = rs.Fields(nCol).Value
                            .Range("J" & u) = rs.Fields("amount").Value
                            theSum = theSum + rs.Fields("amount").Value
                            rs.MoveNext
                        Loop
                        If ThisWorkbook.Sheets("BM").Range("J33") = "" Then ThisWorkbook.Sheets("BM").Range("J33") = theSum
                    End With
                End If
                rs.Close
                Set rs = Nothing
'                        If Not orderExists(rs.Fields(pCol).value) Then
'                            'add new order
'                            Set zlec = New clsZlecenie
'                            rs.filter = pCol & " = " & rs.Fields(pCol).value
'                            With zlec
'                                .index = rs.Fields(pCol).value
'                                .Name = rs.Fields(nCol).value
'                                .Order = rs.Fields(oCol).value
'                            End With
'                            zlecenia.Add zlec, CStr(rs.Fields(pCol))
'                        End If
            End If
        End If
    End If
End If

Set rs = Nothing
End Sub

Public Sub roastingMes()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim oCol As String
Dim pCol As String
Dim nCol As String
Dim aCol As String
Dim mCol As String
Dim kgCol As String
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim theSum As Double
'Dim zlec As clsZlecenie
Dim ver As String

If propertyExists("roasting mes file") And ThisWorkbook.CustomDocumentProperties("roasting mes file") <> "" Then
    If ThisWorkbook.CustomDocumentProperties("import path") <> "" Then
        If FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties("roasting mes file") & ".xls") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties("roasting mes file") & ".xls"
        ElseIf FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties("roasting mes file") & ".xlsx") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties("roasting mes file") & ".xlsx"
        Else
            MsgBox "Source file """ & ThisWorkbook.CustomDocumentProperties("roasting mes file") & """ could not be found in " & ThisWorkbook.CustomDocumentProperties("import path") & "\ . Check in settings if both file name and path are correct.", vbOKOnly + vbExclamation, "Error"
        End If
        
        If path <> "" Then
            Set rs = importExcelData(path, "Zestawienie ilości wyprodukowan", 1)
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If Len(oCol) = 0 And Len(pCol) = 0 And Len(nCol) = 0 And Len(aCol) = 0 And Len(mCol) = 0 And Len(kgCol) = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            Select Case rs.Fields(i)
                            Case Is = "Nr zlecenia"
                                oCol = rs.Fields(i).Name
                            Case Is = "Nr produktu"
                                pCol = rs.Fields(i).Name
                            Case Is = "Nazwa produktu"
                                nCol = rs.Fields(i).Name
                            Case Is = "Ilość"
                                aCol = rs.Fields(i).Name
                            Case Is = "Nr maszyny"
                                mCol = rs.Fields(i).Name
                            Case Is = "Ilość kg"
                                kgCol = rs.Fields(i).Name
                            End Select
                        Next i
                    Else
                        rs.Close
                        Set rs = Nothing
                        Exit Do
                    End If
                    rs.MoveNext
                Loop
                If Right(path, 1) = "s" Then
                    ver = "8"
                Else
                    ver = "12"
                End If
                Worksheet = "Zestawienie ilości wyprodukowan$"
                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=0;"";"
                'by order number
                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sub." & nCol & ", sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY " & oCol & ", " & pCol & ", " & nCol & ";"
                'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
                'strSQL = "SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100;"
                cnn.Open conStr
                Set rs = New ADODB.Recordset
                rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs.EOF Then
                    rs.MoveFirst
                    With ThisWorkbook.Sheets("MES")
                        u = 2
                        .Range("A2") = "Order number"
                        .Range("B2") = "ZFOR"
                        .Range("C2") = "Description"
                        .Range("d2") = "Amount [kg]"
                        Do Until rs.EOF
                            u = u + 1
                            .Range("A" & u) = rs.Fields(oCol).Value
                            .Range("B" & u) = rs.Fields(pCol).Value
                            .Range("C" & u) = rs.Fields(nCol).Value
                            .Range("D" & u) = rs.Fields("amount").Value
                            theSum = theSum + rs.Fields("amount").Value
                            rs.MoveNext
                        Loop
                        If ThisWorkbook.Sheets("BM").Range("J19") = "" Then ThisWorkbook.Sheets("BM").Range("J19") = theSum
                    End With
                End If
                rs.Close
                Set rs = Nothing
'                        If Not orderExists(rs.Fields(pCol).value) Then
'                            'add new order
'                            Set zlec = New clsZlecenie
'                            rs.filter = pCol & " = " & rs.Fields(pCol).value
'                            With zlec
'                                .index = rs.Fields(pCol).value
'                                .Name = rs.Fields(nCol).value
'                                .Order = rs.Fields(oCol).value
'                            End With
'                            zlecenia.Add zlec, CStr(rs.Fields(pCol))
'                        End If
            End If
        End If
    End If
End If

Set rs = Nothing
End Sub

Public Sub packingMes()
Dim rs As ADODB.Recordset
Dim path As String
Dim i As Integer
Dim oCol As String
Dim pCol As String
Dim nCol As String
Dim aCol As String
Dim mCol As String
Dim kgCol As String
Dim strSQL As String
Dim cnn As New ADODB.Connection
Dim conStr As String
Dim u As Integer
Dim Worksheet As String
Dim prop As String
'Dim zlec As clsZlecenie
Dim ver As String

prop = "packaging mes file"

If propertyExists(prop) And ThisWorkbook.CustomDocumentProperties(prop) <> "" Then
    If ThisWorkbook.CustomDocumentProperties("import path") <> "" Then
        If FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xls"
        ElseIf FileExists(ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx") Then
            path = ThisWorkbook.CustomDocumentProperties("import path") & "\" & ThisWorkbook.CustomDocumentProperties(prop) & ".xlsx"
        Else
            MsgBox "Source file """ & ThisWorkbook.CustomDocumentProperties(prop) & """ could not be found in " & ThisWorkbook.CustomDocumentProperties("import path") & "\ . Check in settings if both file name and path are correct.", vbOKOnly + vbExclamation, "Error"
        End If
        
        If path <> "" Then
            Set rs = importExcelData(path, "Zestawienie ilości wyprodukowan", 1)
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If Len(oCol) = 0 And Len(pCol) = 0 And Len(nCol) = 0 And Len(aCol) = 0 And Len(mCol) = 0 And Len(kgCol) = 0 Then
                        For i = 0 To rs.Fields.Count - 1
                            Select Case rs.Fields(i)
                            Case Is = "Nr zlecenia"
                                oCol = rs.Fields(i).Name
                            Case Is = "Nr produktu"
                                pCol = rs.Fields(i).Name
                            Case Is = "Nazwa produktu"
                                nCol = rs.Fields(i).Name
                            Case Is = "Ilość"
                                aCol = rs.Fields(i).Name
                            Case Is = "Nr maszyny"
                                mCol = rs.Fields(i).Name
                            Case Is = "Ilość kg"
                                kgCol = rs.Fields(i).Name
                            End Select
                        Next i
                    Else
                        rs.Close
                        Set rs = Nothing
                        Exit Do
                    End If
                    rs.MoveNext
                Loop
                If Right(path, 1) = "s" Then
                    ver = "8"
                Else
                    ver = "12"
                End If
                Worksheet = "Zestawienie ilości wyprodukowan$"
                conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path & "';" & "Extended Properties=""Excel " & ver & ".0;HDR=NO;IMEX=0;"";"
                'by order number
                strSQL = "SELECT sub." & oCol & ", sub." & pCol & ", sub." & nCol & ", sum(sub." & kgCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY " & oCol & ", " & pCol & ", " & nCol & ";"
                'strSQL = "SELECT sub." & oCol & " as order, sub." & pCol & " as product, sub." & nCol & " as description, sum(sub." & aCol & ") as amount FROM ( SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100) sub GROUP BY order, product, description;"
                'strSQL = "SELECT * FROM [" & Worksheet & "]  WHERE F2 > 100;"
                cnn.Open conStr
                Set rs = New ADODB.Recordset
                rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs.EOF Then
                    rs.MoveFirst
                    With ThisWorkbook.Sheets("MES")
                        u = 2
                        .Range("M2") = "Order number"
                        .Range("N2") = "ZFIN"
                        .Range("O2") = "Description"
                        .Range("P2") = "Amount [kg]"
                        Do Until rs.EOF
                            u = u + 1
                            .Range("M" & u) = rs.Fields(oCol).Value
                            .Range("N" & u) = rs.Fields(pCol).Value
                            .Range("O" & u) = rs.Fields(nCol).Value
                            .Range("P" & u) = rs.Fields("amount").Value
                            theSum = theSum + rs.Fields("amount").Value
                            rs.MoveNext
                        Loop
                        If ThisWorkbook.Sheets("BM").Range("H46") = "" Then ThisWorkbook.Sheets("BM").Range("H46") = theSum
                    End With
                End If
                rs.Close
                Set rs = Nothing
'                        If Not orderExists(rs.Fields(pCol).value) Then
'                            'add new order
'                            Set zlec = New clsZlecenie
'                            rs.filter = pCol & " = " & rs.Fields(pCol).value
'                            With zlec
'                                .index = rs.Fields(pCol).value
'                                .Name = rs.Fields(nCol).value
'                                .Order = rs.Fields(oCol).value
'                            End With
'                            zlecenia.Add zlec, CStr(rs.Fields(pCol))
'                        End If
            End If
        End If
    End If
End If

Set rs = Nothing
End Sub

Public Function orderExists(orderNumber As Long) As Boolean
Dim z As clsZlecenie

If zlecenia.Count = 0 Then
    orderExists = False
Else
    orderExists = False
    For Each z In zlecenia
        If z.index = orderNumber Then
            orderExists = True
            Exit For
        End If
    Next z
End If
End Function


Public Sub formatMES()
Dim rng As Range
Dim lastRow As Long

With ThisWorkbook.Sheets("MES")
    lastRow = .Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("A2:D" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("A2:D2")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
    
    lastRow = .Range("G:G").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("G2:J" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("G2:J2")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
    
    lastRow = .Range("M:M").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).row
    Set rng = .Range("M2:P" & lastRow)
    rng.BorderAround 1, xlMedium, xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).LineStyle = 1
    rng.Borders(xlInsideHorizontal).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = 1
    rng.Borders(xlInsideVertical).ColorIndex = xlColorIndexAutomatic
    rng.Borders(xlInsideVertical).Weight = xlThin
    Set rng = .Range("M2:P2")
    rng.Interior.ColorIndex = 15
    rng.Font.Bold = True
End With

Set rng = Nothing

End Sub


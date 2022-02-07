Attribute VB_Name = "handleMails"

Public Function produceTable(ByRef tHeader() As String) As String
Dim htmlText As String
Dim strTableBeg As String
Dim strTableHeader As String
    'Define format for output
    strTableBeg = "<br><br><table border=1 cellpadding=3 cellspacing=0><font size=3 face=" & Chr(34) & "Arial" & Chr(34) & ">"
    
        strTableHeader = ""
        For x = LBound(tHeader, 2) To UBound(tHeader, 2)
            If tHeader(1, x) <> "" Then
                strTableHeader = strTableHeader & "<tr>" & TD(tHeader(1, x), True) & TD(tHeader(2, x)) & "</tr>"
            End If
        Next x
        produceTable = strTableBeg & strTableHeader & "</font></table>"
 

End Function

Function TD(strIn As String, Optional header As Variant) As String
    If Not IsMissing(header) Then
        If header Then
            TD = "<TD nowrap bgcolor=lightblue><b>" & strIn & "</b></TD>"
        Else
        TD = "<TD nowrap >" & strIn & "</TD>"
        End If
    Else
        TD = "<TD nowrap >" & strIn & "</TD>"
    End If
End Function


Public Sub sendMail(mailBody As String, mailSubject As String, sendTo As String, Optional sendCC As Variant, Optional isImportant As Variant, Optional attachmentPath As Variant)

Dim Mail As New Message
Dim Config As Configuration
Set Config = Mail.Configuration

Config(cdoSendUsingMethod) = cdoSendUsingPort
Config(cdoSMTPServer) = "smtp.gmail.com"
Config(cdoSMTPServerPort) = 465
Config(cdoSMTPAuthenticate) = cdoBasic
Config(cdoSMTPUseSSL) = True
Config(cdoSendUserName) = MailLogin
Config(cdoSendPassword) = MailPassword
Config.Fields.Update

Mail.To = sendTo
Mail.From = Config(cdoSendUserName)
Mail.Subject = mailSubject
Mail.htmlBody = mailBody

Mail.Send

End Sub


Sub xx()
Dim x As Integer
Dim htmlBody As String
ReDim v(2, 1) As String
For x = 1 To 5
    
    v(1, x) = "naglowek" & x
    v(2, x) = "wartosc" & x
    ReDim Preserve v(2, x + 1) As String
Next x

htmlBody = produceTable(v)
Debug.Print htmlBody
'Call SendMail(htmlBody, "nowy temat", "robert.roszak@gmail.com, robert.roszak@demb.com")
End Sub

Public Function toHtml(str As String, Optional isBold As Variant, Optional attributes As Variant) As String
Dim openMark As String
Dim endMark As String

If Not IsMissing(attributes) Then
    Select Case attributes
    Case Is = "i"
        openMark = "<i>"
        endMark = "</i>"
    Case Is = "u"
        openMark = "<u>"
        endMark = "</u>"
    End Select
    str = openMark & str & endMark
End If
If Not IsMissing(isBold) Then
    If isBold Then
        toHtml = "<font size=3 face=" & Chr(34) & "Arial" & Chr(34) & "><b>" & str & "</b></font>"
    Else
        toHtml = "<font size=3 face=" & Chr(34) & "Arial" & Chr(34) & ">" & str & "</font>"
    End If
Else
    toHtml = "<font size=3 face=" & Chr(34) & "Arial" & Chr(34) & ">" & str & "</font>"
End If
End Function

Sub cc()
Call sendMail(toHtml("łęką późno"), "subject", "robert.roszak@demb.com")
End Sub


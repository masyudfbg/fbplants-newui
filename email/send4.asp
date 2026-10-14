<%
Sub sendmail(email, subject, message)

Dim i
if email="" or isnull(email) then
        strMessage = "Email address is empty. Cannot send email."
        Exit Sub
end if

'On Error Resume Next

        Version = Request.ServerVariables("SERVER_SOFTWARE")
        If InStr(Version, "Microsoft-IIS") > 0 Then
                i = InStr(Version, "/")
                If i > 0 Then
                        IISVer = Trim(Mid(Version, i+1))
                End If
        End If

        If IISVer <= "5.0" or IISVer = "7.0" Then
                ' Windows NT / 2000 
                'Set myMail = Server.CreateObject("CDONTS.NewMail")
                'myMail.From = cfrom
                'myMail.To = email

                'myMail.Subject = subject 
                'myMail.Body = message 
                'myMail.Send
                'Set myMail = Nothing

Set cdoMessage = CreateObject("CDO.Message")
Set cdoConfig = CreateObject("CDO.Configuration")

Set Flds = cdoConfig.Fields

With Flds
.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "intra.indo-palm.co.id"
.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
.Update
End With

With cdoMessage
Set .Configuration = cdoConfig
.To = email
.From = cfrom 
.Subject = subject
.TextBody = message
.Send
End With

Set cdoMessage = Nothing
Set cdoConfig = Nothing
Set Flds = Nothing

SendEmail = true


Else
                ' Windows XP / 2003 
        Set myMail=CreateObject("CDO.Message")
        myMail.Subject = subject 
        myMail.From=cfrom
        myMail.To=email
        myMail.TextBody= message 
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
        'Name or IP of remote SMTP server
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="intra.indo-palm.co.id"
        'Server port
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 
        ' SMTP username and passwords
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "12345678"
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "mail.coba@indo-palm.co.id"
        if csmtpuser<>"" then _
                myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        myMail.Configuration.Fields.Update
        myMail.Send
        Set myMail = Nothing 
End If


        if Err.Number<>0 then
                strMessage = "Error happened sending email to " & email & "<br>" & Err.Source & "<br>" & Err.Number & "<br>" & Err.Description

                Err.Clear
        end if


End Sub

%>
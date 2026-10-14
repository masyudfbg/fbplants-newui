<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Type Library" -->
<%
'--------------------------Start Send Email--------------------
Dim mailFrom, mailTo, mailSubject, mailMessage, mailServer, mailUsername, mailPassword, MyMail, mailCC, mailBCC
mailBCC		= "jaya.mega@indo-palm.co.id"
mailFrom 	= "mail.coba@indo-palm.co.id"
mailTo 		= "admin@indo-palm.co.id"
mailSubject = "Percobaan autosending email "&Month(Now())&"/"&Day(Now())&"/"&Year(Now())
mailMessage = "<p>Halo..!!,</p>"
mailMessage = mailMessage & "<p>Mohon dicheck percobaan pengiriman gambar GIS, untuk report: " & FormatDateTime(Now,1) & "</p>"
mailMessage = mailMessage & "<p>Best Regards,<br>"
mailMessage = mailMessage & "AMS-KMT Department</p>"
mailServer 	= "intra.indo-palm.co.id"
mailUsername = "mail.coba@indo-palm.co.id"
mailPassword = "12345678"


Sub sendmail(mailFrom, mailTo, mailSubject, mailMessage, mailServer, mailUsername, mailPassword, mailBCC)
Set myMail = CreateObject ("CDO.Message")
myMail.BCC	= mailBCC
myMail.From = mailFrom
myMail.To = mailTo
myMail.Subject = mailSubject
myMail.HTMLBody = mailMessage
'myMail.AddAttachment "D:\WebSites2\weberp\reports\exports\OutstandingPODetails_"& Year(Now())&Month(Now())&Day(Now()) &".pdf"
'myMail.AddAttachment "C:\Inetpub\wwwroot\email\upload\sunset.jpg"
'myMail.AddAttachment "OutstandingPODetails_"& Year(Now())&Month(Now())&Day(Now()) &".pdf"
myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 120
myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mailServer
myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = mailUsername
myMail.Configuration.Fields.Item ("http://schemas.microsofthttp://schemas.microsoft.com/cdo/configuration/sendusing.com/cdo/configuration/sendpassword") = mailPassword
myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
myMail.Configuration.Fields.Update
myMail.Send
Set myMail = nothing
End Sub

Call sendmail(mailFrom, mailTo, mailSubject, mailMessage, mailServer, mailUsername, mailPassword, mailBCC)

'--------------------------End Send Email----------------------

%>
<%
Set myMail=CreateObject("CDO.Message")
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="intra.indo-palm.co.id" ' On our server we have the IP in here
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 120
myMail.Subject="Sending email with CDO"
myMail.From="jaya.mega@indo-palm.co.id"
myMail.To="mail.coba@indo-palm.co.id"
myMail.TextBody="This is a message."
myMail.Send
set myMail=nothing
%>
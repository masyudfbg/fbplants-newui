<%
Dim ObjSendMail
Set ObjSendMail = CreateObject("CDO.Message")
     
'This section provides the configuration information for the remote SMTP server.
     
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="103.58.100.144"
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True 'Use SSL for the connection (True or False)
ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
     
' If your server requires outgoing authentication uncomment the lines bleow and use a valid email address and password.
'ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
'ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="YOU@DOMAINNAME.COM"
'ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="PASSWORD"
     
ObjSendMail.Configuration.Fields.Update
     
'End remote SMTP server configuration section==
     
ObjSendMail.To = "jaya.mega@indo-palm.co.id"
ObjSendMail.Subject = "ini percobaan dari win-2008"
ObjSendMail.From = "jaya.mega@indo-palm.co.id"
     
' we are sending a text email.. simply switch the comments around to send an html email instead
'ObjSendMail.HTMLBody = "this is the body"
ObjSendMail.TextBody = "this is the body"
     
ObjSendMail.Send
     
Set ObjSendMail = Nothing
 
Response.Write "Mail Send"
%> 
Option Explicit
On Error Resume Next

' Declare our vars
Dim objWinHttp, strURL

' Request URL from 1st Command Line Argument.  This is
' a nice option so you can use the same file to
' schedule any number of differnet scripts just by
' changing the command line parameter.
'strURL = WScript.Arguments(0)

' Could also hard code if you want:
'strURL = "http://harimanto/weberp/reports/ExportOutstandingPODetails.asp"
strURL = "http://172.30.61.221/email/SendMail.asp"

' For more WinHTTP v5.0 info, including where to get
' the component, see our HTTP sample:
' http://www.asp101.com/samples/winhttp5.asp
Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
objWinHttp.Open "GET", strURL
objWinHttp.Send

' Get the Status and compare it to the expected 200
' which is the code for a successful HTTP request:
' http://www.asp101.com/resources/httpcodes.asp
If objWinHttp.Status <> 200 Then
	' If it's not 200 we throw an error... we'll
	' check for it and others later.
	Err.Raise 1, "HttpRequester", "Invalid HTTP Response Code"
End If

' Since in this example I could really care less about
' what's returned, I never even check it, but in
' general checking for some expected text or some sort
' of status result from the ASP script would be a good
' idea.  Use objWinHttp.ResponseText

Set objWinHttp = Nothing

'If Err.Number <> 0 Then
'	' Something has gone wrong... do whatever is
'	' appropriate for your given situation... I'm
'	' emailing someone:
'	Dim objMessage
'	Set objMessage = Server.CreateObject("CDO.Message")
'	objMessage.To       = "Your Name <user@some domain.com>"
'	objMessage.From     = "Your Name <user@some domain.com>"
'	objMessage.Subject  = "An Error Has Occurred in a " _
'		& "Scheduled Task"
'	objMessage.TextBody = "Error #: " & Err.Number & vbCrLf _
'		& "From: " & Err.Source & vbCrLf _
'		& "Desc: " & Err.Description & vbCrLf _
'		& "Time: " & Now()
'							
'	objMessage.Send
'	Set objMessage = Nothing
'End If

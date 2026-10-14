<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salescontractno = Request("salescontractno")
fcba = Request("fcba")
result = ""

q = " SELECT sah.salescontractno, sad.vat ,sah.dp_vat" & _
" FROM SALESAGREEMENTHDR sah" & _
" INNER JOIN SALESAGREEMENTDET sad ON (sah.salescontractno = sad.salescontractno AND sah.fcba = sad.fcba)" & _
" WHERE sad.salescontractno = '"& salescontractno &"' and sah.fcba = '" & fcba & "'" 


' response.Write q
' Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
' 					  rsCallCombo.Open q, connstr
' 					  if not rsCallCombo.eof then
' 							html = "<select name='salesccontractno'>"
' 							Do While Not rsCallCombo.eof
' 							html = html & "<option value='" & rsCallCombo("salescontractno") & "'>" & rsCallCombo("salescontractno") & "</option>" & _
' 								rsCallCombo.movenext
' 							loop
' 							html = html & "</select>" 
' 							else
							
' 							html = "<select name='salesccontractno'></select>"
' 					  End if
' 					  rsCallCombo.Close
' 					  set rsCallCombo = Nothing  
					  
' 	response.expires=-1   
' 	response.write(html)

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	vat = rsrecurtrans("vat")
	dp_vat = rsrecurtrans("dp_vat")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """vat"" : """ & vat & ""","
response.write """dp_vat"" : """ & dp_vat & """"
response.write "}"

%>
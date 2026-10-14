<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transportercontractno = Request("transportercontractno")
vat = 0

q = " SELECT *" & _
	" FROM transportercontract" & _
	" WHERE transportercontractno = '"& transportercontractno &"' and fcba = '" & Session("fcba") & "'" 

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	vat = cdbl(to_nol(rsrecurtrans("vat")))
else
	vat = 0
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """vat"" : """ & vat & """"
response.write "}"

%>
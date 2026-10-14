<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
controljob = Request("controljob")
invoicedpno = Request("invoicedpno")
fcba = Request("fcba")
result = ""

q = " SELECT unitofmeasurement FROM job WHERE fcba = '"& session("fcba") &"' AND fccode = '"& controljob &"'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	uomjob = rsrecurtrans("unitofmeasurement")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """uomjob"" : """ & uomjob & """"
response.write "}"

%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transportercode = Request("transportercode")
rentby = Request("rentby")

result = ""

q = " select * from transporter where fccode ='"& transportercode &"' and fcba ='"& rentby &"'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	' suppliername = rsrecurtrans("fcname")
	transportername = rsrecurtrans("fcname")
	transporteraddress = rsrecurtrans("address")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """transportername"" : """ & transportername & ""","
response.write """transporteraddress"" : """ & transporteraddress & """"
response.write "}"

%>
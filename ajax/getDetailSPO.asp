<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
businessunit = Request("businessunit")

result = ""

q = " select * from servicepurchaseorder where pocode ='"& fccode &"' and fcba='"& businessunit &"'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	' suppliername = rsrecurtrans("fcname")
	suppliername = rsrecurtrans("suppliername")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
' response.write """suppliername"" : """ & suppliername & ""","
response.write """suppliername"" : """ & suppliername & """"
response.write "}"

%>
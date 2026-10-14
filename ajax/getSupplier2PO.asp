<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
suppliercode = Request("suppliercode")
purchaseby = Request("purchaseby")

result = ""

q = " select * from supplier where fccode ='"& suppliercode &"' and fcba ='"& purchaseby &"'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	suppliername = rsrecurtrans("fcname")
	supplieraddress = rsrecurtrans("address")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """suppliername"" : """ & suppliername & ""","
response.write """supplieraddress"" : """ & supplieraddress & """"
response.write "}"

%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salescontractno = Request("salescontractno")
fcba = Request("fcba")

arr_json_eof = true

q = " SELECT invoicedpno, totalamount FROM salesagreementhdr WHERE fcba = '"& fcba &"' AND salescontractno = '"& salescontractno &"'" 
 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	invoicedpno= rsrecurtrans("invoicedpno")
	totalamount= rsrecurtrans("totalamount")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """totalamount"" : """ & totalamount & ""","
response.write """invoicedpno"" : """ & invoicedpno & """"
response.write "}"

%>
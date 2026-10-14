<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salescontractno = Request("salescontractno")
fcba = Request("fcba")

arr_json_eof = true

q = " select sum(dp_aftervat) remainsdp from salesinvoicedp where salescontractno = '"& salescontractno &"' and fcba = '"& fcba &"'" 
 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	remainsdp= rsrecurtrans("remainsdp")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
' response.write """totalamount"" : """ & totalamount & ""","
response.write """remainsdp"" : """ & remainsdp & """"
response.write "}"

%>
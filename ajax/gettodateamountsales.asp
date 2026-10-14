<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salescontractno = Request("salescontractno")
fcba = Request("fcba")

arr_json_eof = true

q = " SELECT CASE WHEN NVL(SUM (dp_aftervat),0) = 0 THEN 0 ELSE NVL(SUM (dp_aftervat),0) END todate_dp_aftervat" & _
	" FROM salesinvoicedp" & _
	" WHERE fcba = '"& fcba &"' AND salescontractno = '"& salescontractno &"'" 
 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	todate_dp_aftervat= rsrecurtrans("todate_dp_aftervat")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """todate_dp_aftervat"" : """ & todate_dp_aftervat & """"
response.write "}"

%>
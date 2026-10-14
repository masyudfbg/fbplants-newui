<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salescontractno = Request("salescontractno")
fcba = Request("fcba")

arr_json_eof = true

q = " SELECT CASE WHEN NVL(MAX (percentage_dp),0) = 0 THEN 0 ELSE NVL(MAX (percentage_dp),0) END todate_percentage_dp" & _
	" FROM purchaseinvoicedp" & _
	" WHERE fcba = '"& fcba &"' AND purchasecontractno = '"& salescontractno &"'" 
 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	todate_percentage_dp= rsrecurtrans("todate_percentage_dp")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """todate_percentage_dp"" : """ & todate_percentage_dp & """"
response.write "}"

%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
wbt_chitno = Request("wbt_chitno")
result = ""

q = " SELECT dn.transportco, dn.postingdate, dn.quantity, wb.vehicleno " & _
" FROM DELIVERYNOTE dn" & _
" LEFT JOIN weighbridgeticket wb " & _
" ON (dn.fcba = wb.fcba and dn.wbt_chitno = wb.chitno)" & _
" WHERE dn.wbt_chitno= '"& wbt_chitno &"' and dn.fcba = '" & Session("fcba") & "'" 


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	postingdate = rsrecurtrans("postingdate")
	quantity = rsrecurtrans("quantity")
	transportco = rsrecurtrans("transportco")
	vehicleno = rsrecurtrans("vehicleno")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """postingdate"" : """ & postingdate & ""","
response.write """transportco"" : """ & transportco & ""","
response.write """quantity"" : """ & quantity & ""","
response.write """vehicleno"" : """ & vehicleno & """"
response.write "}"
%>
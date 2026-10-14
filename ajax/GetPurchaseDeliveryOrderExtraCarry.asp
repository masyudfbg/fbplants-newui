<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
deliveryorderno = Request("deliveryorderno")
fcba = Request("fcba")
doline = Request("doline")
result = ""

q =	  " select * from purchasedeliveryorder " & _
	  " WHERE fcba = '" & fcba & "'" & _
	  " AND deliveryorderno = '" & deliveryorderno & "' AND lineno = '"& doline &"'"  
	  
' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	quantity_do = 0
	salescontractno = rsrecurtrans("purchasecontractno")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing


response.write "{"
response.write """quantity_do"" : """ & quantity_do & ""","
response.write """salescontractno"" : """ & salescontractno & """"
response.write "}"
%>
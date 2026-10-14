<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
deliveryorderno = Request("deliveryorderno")
fcba = Request("fcba")
doline = Request("doline")
result = ""

q =	  " select quantity_do,a.salescontractno,productcode,customercode,transporter transportercode from deliveryorder a INNER JOIN salesagreementhdr b ON a.salescontractno=b.salescontractno AND a.fcba=b.fcba " & _
	  " WHERE a.fcba = '" & fcba & "'" & _
	  " AND deliveryorderno = '" & deliveryorderno & "' AND a.lineno = '"& doline &"'"  
	  
 'response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	quantity_do = rsrecurtrans("quantity_do")
	salescontractno = rsrecurtrans("salescontractno")
	productcode = rsrecurtrans("productcode")
	customercode = rsrecurtrans("customercode")
	transportercode = rsrecurtrans("transportercode")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing


response.write "{"
response.write """quantity_do"" : """ & quantity_do & ""","
response.write """salescontractno"" : """ & salescontractno & ""","
response.write """productcode"" : """ & productcode & ""","
response.write """customercode"" : """ & customercode & ""","
response.write """transportercode"" : """ & transportercode & """"


response.write "}"
%>
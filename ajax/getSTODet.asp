<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
stono = Request("stono")
stolineno = Request("stolineno")

result = ""

q = " select * from stocktransferorderdtl where stono ='"& stono &"' and stolineno ='"& stolineno &"'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	stockname = rsrecurtrans("stockname")
	itemcode = rsrecurtrans("itemcode")
	quantityorder = rsrecurtrans("quantityorder")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """stockname"" : """ & stockname & ""","
response.write """itemcode"" : """ & itemcode & ""","
response.write """quantityorder"" : """ & quantityorder & """"
response.write "}"

%>
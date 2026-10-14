<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
assettaxcode = Request("assettaxcode")
result = ""

q =   " SELECT fcname FROM fixedasset_taxgroup WHERE fccode = '"& assettaxcode &"'"


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	assettaxname = rsrecurtrans("fcname")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """assettaxname"" : """ & assettaxname & """"
response.write "}"
%>
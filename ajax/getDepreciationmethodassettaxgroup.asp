<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
assettaxcode = Request("assettaxcode")
fcba = Request("fcba")
result = ""

q = "select depreciationrange from fixedasset_taxgroup where fccode = '"& assettaxcode &"'" 
' response.Write q

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	depreciationrange = rsrecurtrans("depreciationrange")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """depreciationrange"" : """ & depreciationrange & """"
response.write "}"

%>
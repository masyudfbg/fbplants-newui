<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
assetgroupcode = Request("assetgroupcode")
fcba = Request("fcba")
result = ""

q = "select depreciationrange from assetgroup where fccode = '"& assetgroupcode &"'" 
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
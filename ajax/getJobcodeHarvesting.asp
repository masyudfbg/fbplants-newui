<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
prcode = Request("prcode")
jobcode = Request("jobcode")
result = ""

q = " SELECT COUNT (*) jobcodespbharvesting" & _
	"  FROM parameterdetail" & _
	" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
	"       AND fcba = '"& session("fcba") &"'" & _
	"       AND fccode = '"& jobcode &"'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """jobcodespbharvesting"" : """ & jobcodespbharvesting & """"
response.write "}"

%>
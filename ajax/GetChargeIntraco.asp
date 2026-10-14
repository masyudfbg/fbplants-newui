<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
prcode = Request("prcode")
jobcode = Request("jobcode")
result = ""

q = " SELECT COUNT (1) jobcodeintraco" & _
	"  FROM intraco_tab" & _
	" WHERE fccode = '"& jobcode &"' and status='TRUE'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	jobcodeintraco = rsrecurtrans("jobcodeintraco")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """jobcodeintraco"" : """ & jobcodeintraco & """"
response.write "}"

%>
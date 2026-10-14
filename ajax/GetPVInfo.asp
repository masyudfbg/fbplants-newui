<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
target = Request("target")
jobcode = Request("jobcode")

q = " SELECT DISTINCT TARGET FROM TRIGGER_CONTROL WHERE ACTIVITY = '"& jobcode &"' AND TARGET=JOBMASTER_API.get_job_category('" & jobcode & "')"

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	target= rsrecurtrans("target")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """target"" : """ & target & """"
response.write "}"

%>
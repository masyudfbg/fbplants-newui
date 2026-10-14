<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
lineno = Request("lineno")
fcba = Request("fcba")
transportercontractno = Request("transportercontractno")
result = ""

q =	"  SELECT tc.transporter," & _
	"       tc.rate," & _
	"       tc.losstolerance," & _
	"       tr.fcname" & _
	"  FROM transportercontract tc" & _
	"       INNER JOIN transporter tr" & _
	"          ON (tr.fccode = tc.transporter AND tr.fcba = tc.fcba)" & _
	" WHERE tc.lineno= '"& lineno &"' and tc.fcba = '" & fcba & "' and tc.transportercontractno = '"& transportercontractno &"'" 


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	transporter = rsrecurtrans("transporter")
	rate = rsrecurtrans("rate")
	losstolerance = rsrecurtrans("losstolerance")
	transportername = rsrecurtrans("fcname")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """transporter"" : """ & transporter & ""","
response.write """transportername"" : """ & transportername & ""","
response.write """rate"" : """ & rate & ""","
response.write """losstolerance"" : """ & losstolerance & """"
response.write "}"
%>
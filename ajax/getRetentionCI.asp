<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
agreementcode = Request("agreementcode")
fcba = Request("fcba")

arr_json_eof = true
q = " SELECT DISTINCT ca.retention, ca.contractorcode, cr.fcname " & _
	"  FROM contractagreement ca" & _
	"       INNER JOIN workinprogress wip" & _
	"          ON (wip.agreementcode = ca.agreementcode AND wip.fcba = ca.fcba)" & _
	"   INNER JOIN contractor cr ON (cr.fccode = ca.contractorcode AND cr.fcba = ca.fcba)" & _
	" WHERE wip.agreementcode = '"& agreementcode &"' AND wip.fcba = '"& fcba &"'"
' response.write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	contractorcode = rsrecurtrans("contractorcode")
	retention = rsrecurtrans("retention")
	contractorname = rsrecurtrans("fcname")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """retention"" : """ & retention & ""","
response.write """contractorname"" : """ & contractorname & ""","
response.write """contractorcode"" : """ & contractorcode & """"
response.write "}"
%>
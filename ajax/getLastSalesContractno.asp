<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
itemcode = Request("itemcode")
fcba = Request("fcba")
result = ""

q =   	"	SELECT DISTINCT salescontractno" & _
		"  FROM (  SELECT MAX (sah.salescontractno) salescontractno" & _
		"            FROM salesagreementhdr sah" & _
		"                 INNER JOIN salesagreementdet sad" & _
		"                    ON (    sah.salescontractno = sad.salescontractno" & _
		"                        AND sah.fcba = sad.fcba)" & _
		"           WHERE sah.fcba = '"& fcba &"'" & _
		"        GROUP BY  sah.fcba)" 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	lastnumber = rsrecurtrans("salescontractno")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
' response.write """assetname"" : """ & assetname & ""","
' response.write """fa_groupid"" : """ & fa_groupid & ""","
' response.write """acquisitionvalue"" : """ & acquisitionvalue & ""","
' response.write """costcenter"" : """ & costcenter & ""","
' response.write """location_id"" : """ & location_id & ""","
response.write """lastnumber"" : """ & lastnumber & """"
response.write "}"
%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
indikator = Request("indikator")
fcba = Request("fcba")
result = ""

q =   	"  SELECT ffa, moisture, dirt, dobi" & _
		"  FROM millquality WHERE fcba = '"& fcba &"' AND fccode = '"& indikator &"'" 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	ffa = cdbl(to_nol(rsrecurtrans("ffa")))
	moisture = cdbl(to_nol(rsrecurtrans("moisture")))
	dirt = cdbl(to_nol(rsrecurtrans("dirt")))
	dobi = cdbl(to_nol(rsrecurtrans("dobi")))
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
' response.write """assetname"" : """ & assetname & ""","
' response.write """fa_groupid"" : """ & fa_groupid & ""","
response.write """moisture"" : """ & moisture & ""","
response.write """dirt"" : """ & dirt & ""","
response.write """ffa"" : """ & ffa & ""","
response.write """dobi"" : """ & dobi & """"
response.write "}"
%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fixassetcode = Request("fixassetcode")
fcba = Request("fcba")
result = ""

q = " SELECT * FROM fixedasset_transportask WHERE fcba = '"& fcba &"'"
 
' response.Write q

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	fa_transportaskdate = rsrecurtrans("transdate")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """fa_transportaskdate"" : """ & fa_transportaskdate & """"
' response.write """total"" : """ & total & """"
' response.write """jobcode_depreciation"" : """ & jobcode_depreciation & ""","
' response.write """targettype_depreciation"" : """ & targettype_depreciation & ""","
' response.write """targetcode_depreciation"" : """ & targetcode_depreciation & """"
response.write "}"

%>
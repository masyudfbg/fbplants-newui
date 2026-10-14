<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
typex = cdbl(to_nol(Request("typex")))
hours = cdbl(to_nol(Request("hours")))

arr_json_eof = true
q = "select nvl(convertion,0) convertion from overtime_convertion where type = '"& typex &"' and hours = '"& hours &"' and fcba = '"& session("fcba") &"'" 

'response.write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	arr_json = rsrecurtrans.getrows()
	arr_json_eof = false
end if
rsrecurtrans.Close : set rsrecurtrans = Nothing

if not arr_json_eof then
	response.write "{"
	for xyyy=0 to ubound(arr_json,2)
		response.write """convertion"" : """ & arr_json(0,xyyy) & """"
	next
	response.write "}"
else 
response.write "{"
response.write """convertion"" : ""0"""
response.write "}"

end if

%>
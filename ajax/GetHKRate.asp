<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")

q = " select rate_1 from allowded where fcba ='" & fcba & "' and fccode ='A01'"
Dim arr_json 
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr
arr_json_eof = true

if Not rsrecurtrans.eof then
	arr_json = rsrecurtrans.getrows()
	arr_json_eof = false
end if
rsrecurtrans.Close : set rsrecurtrans = Nothing

if not arr_json_eof then
	response.write "{"
	for xyyy=0 to ubound(arr_json,2)
		response.write """ratemandays"" : """ & arr_json(0,xyyy) & """"
	next
	response.write "}"
end if
%>
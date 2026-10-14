<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
activitycode = Request("activitycode")

q = "select MANDAYSNORM from ACTIVITY_NORM Where JOB = '" & activitycode & "' and FCBA = '" & fcba & "'"

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	mandaysnorm= rsrecurtrans("MANDAYSNORM") 
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{" 
response.write """normamandays"" : """ & mandaysnorm & """"
response.write "}"

'set rsrecurtrans = server.CreateObject("Adodb.recordset")
'rsrecurtrans.Open q, connstr

'if Not rsrecurtrans.eof then
'	arr_json = rsrecurtrans.getrows()
'	arr_json_eof = false
'end if
'rsrecurtrans.Close : set rsrecurtrans = Nothing

'if not arr_json_eof then
'	response.write "{"
'	for xyyy=0 to ubound(arr_json,2)
'		response.write """normamandays"" : """ & arr_json(0,xyyy) & """"
'	next
'	response.write "}"
'end if
%>
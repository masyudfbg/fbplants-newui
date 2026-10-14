<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<%
		
arr_data_eof = true
documentno = Request("documentno")

q = " SELECT accountingyear, finishdate FROM recurring" & _
	" WHERE documentno = '" & documentno & "'" & _
	"       AND fcba = '" & Session("fcba") & "'"

'response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	arr_data = rsrecurtrans.getrows()
	arr_data_eof = false
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

if not arr_data_eof then
	post_json = "{"
	for xi= 0 to ubound(arr_data,2)
		post_json = post_json & """accountingyear"":"""&arr_data(0, xi)&""",""finishdate"":"""&arr_data(1, xi)&""""
	next
	post_json = post_json & "}"
end if
response.write post_json
%>
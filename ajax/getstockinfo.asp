<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")

arr_json_eof = true
q = " SELECT fcname, unitofmeasurement, stockgroupcode" & _
	" FROM bankstock" & _
	" WHERE fccode = '"&fccode&"'"
	

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
		response.write """fcname"" : """ & arr_json(0,xyyy) & ""","
		response.write """unitofmeasurement"" : """ & arr_json(1,xyyy) & ""","
		response.write """stockgroupcode"" : """ & arr_json(2,xyyy) & """"
	next
	response.write "}"
end if

%>
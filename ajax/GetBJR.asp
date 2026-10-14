<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fieldcode = Request("fieldcode")
receptiondate = Request("receptiondate")

arr_json_eof = true
q = "select field_api.GetValid_pressemester_abw('"& fieldcode &"', '"& session("fcba") &"', to_date('"& receptiondate &"', 'mm/dd/yyyy')) pressemester_abw from dual" 

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
		response.write """pressemester_abw"" : """ & arr_json(0,xyyy) & """"
	next
	response.write "}"
else 
response.write "{"
response.write """pressemester_abw"" : """
response.write "}"

end if

%>
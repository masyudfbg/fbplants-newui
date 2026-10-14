<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fieldcode = Request("fieldcode")
fcba = Request("fcba")
txtMonth = Request("txtMonth")
txtYear = Request("txtYear")

arr_json_eof = true
q = " SELECT field_api.getvalid_ha(fccode, fcba, TO_DATE('"& txtMonth &"/1/"& txtYear &"', 'mm/dd/yyyy')) hectarageplanted" & _
	" , field_api.getvalid_plantingpoint(fccode, fcba, TO_DATE('"& txtMonth &"/1/"& txtYear &"', 'mm/dd/yyyy')) plantingpoint" & _
	" , field_api.getvalid_plantingyear(fccode, fcba, TO_DATE('"& txtMonth &"/1/"& txtYear &"', 'mm/dd/yyyy')) plantingyear" & _
	" FROM field" & _
	" WHERE fccode = '"& fieldcode &"' AND fcba = '" & fcba & "'"

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
		response.write """hectarageplanted"" : """ & arr_json(0,xyyy) & ""","
		response.write """plantingpoint"" : """ & arr_json(1,xyyy) & ""","
		response.write """plantingyear"" : """ & arr_json(2,xyyy) & """"
	next
	response.write "}"
end if

%>
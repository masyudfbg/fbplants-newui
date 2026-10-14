<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transportercontractno = Request("transportercontractno")
fcba = Request("fcba")

arr_json_eof = true
q = " SELECT creationdate, transporter, startdate, enddate, shippingaddress, rate, unitofmeasure, losstolerance, vehicleno" & _
	" FROM transportercontract" & _
	" WHERE transportercontractno = '"&transportercontractno&"'"
if len(trim(fcba)) > 0 then q = q & "       AND fcba = '"&fcba&"'"

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
		response.write """creationdate"" : """ & arr_json(0,xyyy) & ""","
		response.write """transporter"" : """ & arr_json(1,xyyy) & ""","
		response.write """startdate"" : """ & arr_json(2,xyyy) & ""","
		response.write """enddate"" : """ & arr_json(3,xyyy) & ""","
		response.write """shippingaddress"" : """ & arr_json(4,xyyy) & ""","
		response.write """rate"" : """ & arr_json(5,xyyy) & ""","
		response.write """unitofmeasure"" : """ & arr_json(6,xyyy) & ""","
		response.write """losstolerance"" : """ & arr_json(7,xyyy) & ""","
		response.write """vehicleno"" : """ & arr_json(8,xyyy) & """"
	next
	response.write "}"
else
	response.write "{"
	response.write """creationdate"" : """","
	response.write """transporter"" : """","
	response.write """startdate"" : """","
	response.write """enddate"" : """","
	response.write """shippingaddress"" : """","
	response.write """rate"" : """","
	response.write """unitofmeasure"" : """","
	response.write """losstolerance"" : """","
	response.write """vehicleno"" : """""
	response.write "}"
end if

%>
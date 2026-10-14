<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
stono = Request("stono")
fcba = Request("fcba")

arr_json_eof = true
q =	"	SELECT sth.stono," & _
	"       sth.transportercode," & _
	"       tr.fcname transportername," & _
	"       NVL (sth.ppn, 0) pph," & _
	"       NVL (sth.dp, 0) dp" & _
	"  FROM stocktransferorderhdr sth" & _
	"       LEFT JOIN transporter tr" & _
	"          ON (tr.fccode = sth.transportercode AND tr.fcba = sth.fcba)" & _
	" WHERE 1 = 1"  

if len(trim(fcba)) > 0 then q = q & "       AND sth.fcba = '"&fcba&"'"
if len(trim(stono)) > 0 then q = q & "       AND sth.stono = '"&stono&"'"

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
		response.write """dp"" : """ & arr_json(3,xyyy) & ""","
		response.write """transportercode"" : """ & arr_json(1,xyyy) & ""","
		response.write """transportername"" : """ & arr_json(2,xyyy) & ""","
		response.write """pph"" : """ & arr_json(4,xyyy) & """"
	next
	response.write "}"
else
	response.write "{"
	response.write """dp"" : """","
	response.write """transportername"" : """","
	response.write """transportercode"" : """","
	response.write """pph"" : ""0"""
	response.write "}"
end if

%>
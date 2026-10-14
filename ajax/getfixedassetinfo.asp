<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fixedassetcode = Request("fixedassetcode")
vat = 0

arr_json_eof = true
q = " SELECT fa.fccode," & _
	"       fa.fcname," & _
	"       asg.controljob," & _
	"       asg.description" & _
	"  FROM fixedasset fa" & _
	"       LEFT JOIN assetgroup asg ON (asg.fccode = fa.assetgroupcode)" & _
	" WHERE fa.fccode = '"& fixedassetcode &"' AND fcba = '"& session("fcba") &"'" 

' response.write q
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
		response.write """targetcode"" : """ & arr_json(0,xyyy) & ""","
		response.write """fixedassetname"" : """ & arr_json(1,xyyy) & ""","
		response.write """jobcode"" : """ & arr_json(2,xyyy) & ""","
		response.write """description"" : """ & arr_json(3,xyyy) & """"
	next
	response.write "}"
end if

%>
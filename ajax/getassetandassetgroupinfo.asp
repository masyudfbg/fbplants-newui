<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
assetcode = Request("assetcode")
vat = 0

arr_json_eof = true
q = " SELECT ass.fccode assetcode," & _
	"       ass.fcname assetname," & _
	"       ass.assetgroupcode," & _
	"       asg.fcname assetgroupname," & _
	"       ass.assettaxcode," & _
	"       ast.fcname assettaxname," & _
	"       ass.unitofmeasurement," & _
	"       asg.assettypecode," & _
	"       asg.depreciationmethod," & _
	"       asg.depreciationrange," & _
	"       asg.controljob," & _
	"       asg.description" & _
	"  FROM asset ass" & _
	"       LEFT JOIN assetgroup asg ON ass.assetgroupcode = asg.fccode" & _
	"       LEFT JOIN fixedasset_taxgroup ast ON (ast.fccode = ass.assettaxcode)" & _
	" WHERE ass.fccode = '"& assetcode &"'" 


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
		response.write """assetcode"" : """ & arr_json(0,xyyy) & ""","
		response.write """assetname"" : """ & arr_json(1,xyyy) & ""","
		response.write """assetgroupcode"" : """ & arr_json(2,xyyy) & ""","
		response.write """assetgroupname"" : """ & arr_json(3,xyyy) & ""","
		response.write """assettaxcode"" : """ & arr_json(4,xyyy) & ""","
		response.write """assettaxname"" : """ & arr_json(5,xyyy) & ""","
		response.write """unitofmeasurement"" : """ & arr_json(6,xyyy) & ""","
		response.write """assettypecode"" : """ & arr_json(7,xyyy) & ""","
		response.write """depreciationmethod"" : """ & arr_json(8,xyyy) & ""","
		response.write """depreciationrange"" : """ & arr_json(9,xyyy) & ""","
		response.write """controljob"" : """ & arr_json(10,xyyy) & ""","
		response.write """description"" : """ & arr_json(11,xyyy) & """"
	next
	response.write "}"
end if

%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fixassetcode = Request("fixassetcode")
lineno = Request("lineno")
fcba = Request("fcba")
result = ""

q =	"	SELECT bu.fccompanycode," & _
	"       fat.fcba," & _
	"       fat.transdate," & _
	"       fat.fccode," & _
	"       fat.fcname," & _
	"       fat.assetcode," & _
	"       fat.assetname," & _
	"       fat.assetgroupcode," & _
	"       fat.assetgroupname," & _
	"       fat.assettaxcode," & _
	"       fat.assettaxname," & _
	"       fat.uom," & _
	"       fat.depreciationmethod_by," & _
	"       fat.depreciationmonth," & _
	"       fat.acquisitionvalue," & _
	"       fat.additionalvalue," & _
	"       fat.total," & _
	"       fat.description," & _
	"       fat.isactive," & _
	"       fatd.date_transaction," & _
	"       fatd.fctype," & _
	"       fatd.monthlydepreciation," & _
	"       fatd.yearlydepreciation," & _
	"       fatd.lastdepreciationmonth," & _
	"       fatd.lineno" & _
	"  FROM fixedasset fat" & _
	"       INNER JOIN fixedasset_depreciation_ledger fatd" & _
	"          ON (fat.fccode = fatd.fccode AND fat.fcba = fatd.fcba)" & _
	"       LEFT JOIN businessunit bu ON fat.fcba = bu.fccode" & _
	" WHERE  fat.fccode = '"& fixassetcode &"' AND lineno = '"& lineno &"'" 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	fixassetname = rsrecurtrans("fcname")
	fcba = rsrecurtrans("fcba")
	transdate = rsrecurtrans("transdate")
	assetcode = rsrecurtrans("assetcode")
	assetname = rsrecurtrans("assetname")
	assetgroupcode = rsrecurtrans("assetgroupcode")
	assetgroupname = rsrecurtrans("assetgroupname")
	assettaxcode = rsrecurtrans("assettaxcode")
	assettaxname = rsrecurtrans("assettaxname")
	uom = rsrecurtrans("uom")
	depreciationmonth = rsrecurtrans("depreciationmonth")
	totalmonth = cdbl(to_nol(rsrecurtrans("lastdepreciationmonth"))) - 1
	totalyearly = cdbl(to_nol(rsrecurtrans("total"))) - cdbl(to_nol(rsrecurtrans("yearlydepreciation")))
	total = rsrecurtrans("total")
	' totalmonth = sysdate - month(cdbl(to_nol(rsrecurtrans("transdate")))) 
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

' oSQLQuery = " select * " & _
'             " from periodcontrol" & _
' 	        " where fcba = '" & fcba & "'"
' 	  ' response.write oSQLQuery
' set oRecordSet = server.CreateObject("Adodb.recordset")
' oRecordSet.Open oSQLQuery, connstr
' If Not oRecordSet.Eof then
' 	currentdate = oRecordSet("currentdate")
' End if
' 'response.Write strNurseryCulling
' oRecordSet.Close
' Set oRecordSet = Nothing

 ' totalmonth = month(cdbl(to_nol(lastdepreciationmonth)) +1) 

response.write "{"
response.write """fixassetname"" : """ & fixassetname & ""","
response.write """assetcode"" : """ & assetcode & ""","
response.write """assetname"" : """ & assetname & ""","
response.write """assetgroupcode"" : """ & assetgroupcode & ""","
response.write """assetgroupname"" : """ & assetgroupname & ""","
response.write """assettaxcode"" : """ & assettaxcode & ""","
response.write """assettaxname"" : """ & assettaxname & ""","
response.write """uom"" : """ & uom & ""","
response.write """depreciationmonth"" : """ & depreciationmonth & ""","
response.write """totalmonth"" : """ & totalmonth & ""","
response.write """totalyearly"" : """ & totalyearly & ""","
response.write """total"" : """ & total & """"
response.write "}"

%>
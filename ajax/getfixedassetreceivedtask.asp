<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fixassetcode = Request("fixassetcode")
fcba = Request("fcba")
result = ""

q =	"	SELECT ftt.fccode," & _
	"       ftt.transdate," & _
	"       ftt.description," & _
	"       ftd.fixassetcode," & _
	"       ftd.fixassetname," & _
	"       ftd.assetcode," & _
	"       ftd.assetname," & _
	"       ftd.assetgroupcode," & _
	"       ftd.assetgroupname," & _
	"       ftd.assettaxcode," & _
	"       ftd.assettaxname," & _
	"       ftd.uom," & _
	"       ftd.depreciationmonth," & _
	"       ftd.fixassetamount," & _
	"       ftd.targettype_depreciation," & _
	"       ftd.targetcode_depreciation," & _
	"       ftd.jobcode_depreciation," & _
	"       ftd.accumulated_depreciation," & _
	"       ftd.targettype_intraco," & _
	"       ftd.targetcode_intraco," & _
	"       ftd.jobcode_intraco," & _
	"       ftd.intracoamount," & _
	"       ftd.amountreceived" & _
	"  FROM fixedasset_transportask ftt" & _
	"       INNER JOIN fixedasset_transportask_det ftd" & _
	"          ON (ftt.fccode = ftd.fccode AND ftt.fcba = ftd.fcba)" & _
	"       LEFT JOIN businessunit bu ON (ftt.fcba = bu.fccode)" & _
	" WHERE     1 = 1" & _
	"       AND ftd.fixassetcode = '"& fixassetcode &"' AND ftd.fcba = '"& fcba &"'" 
 
' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	targetcode_fixasset = rsrecurtrans("fixassetcode")
	fixassetname = rsrecurtrans("fixassetname")
	assetcode = rsrecurtrans("assetcode")
	assetname = rsrecurtrans("assetname")
	assetgroupcode = rsrecurtrans("assetgroupcode")
	assetgroupname = rsrecurtrans("assetgroupname")
	assettaxcode = rsrecurtrans("assettaxcode")
	assettaxname = rsrecurtrans("assettaxname")
	uom = rsrecurtrans("uom")
	depreciationmonth = rsrecurtrans("depreciationmonth")
	amountreceived = rsrecurtrans("amountreceived")
	fixassetamount = rsrecurtrans("fixassetamount")
	intracoamount = rsrecurtrans("intracoamount")
	accumulated_depreciation = rsrecurtrans("accumulated_depreciation")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """targetcode_fixasset"" : """ & targetcode_fixasset & ""","
response.write """fixassetname"" : """ & fixassetname & ""","
response.write """assetcode"" : """ & assetcode & ""","
response.write """assetname"" : """ & assetname & ""","
response.write """assetgroupcode"" : """ & assetgroupcode & ""","
response.write """assetgroupname"" : """ & assetgroupname & ""","
response.write """assettaxcode"" : """ & assettaxcode & ""","
response.write """assettaxname"" : """ & assettaxname & ""","
response.write """uom"" : """ & uom & ""","
response.write """depreciationmonth"" : """ & depreciationmonth & ""","
response.write """fixassetamount"" : """ & fixassetamount & ""","
response.write """intracoamount"" : """ & intracoamount & ""","
response.write """accumulated_depreciation"" : """ & accumulated_depreciation & ""","
response.write """amountreceived"" : """ & amountreceived & """"
response.write "}"

%>
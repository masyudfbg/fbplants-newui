<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fixassetcode = Request("fixassetcode")
fcba = Request("fcba")
lineno = Request("lineno")
result = ""

set rsrecurtrans = server.CreateObject("Adodb.recordset")
q = " SELECT fatd.lineno," & _
	"        fat.total," & _
	"        CASE" & _
	"           WHEN monthtobe = depreciationmonth_ledger" & _
	"           THEN" & _
	"              yearlydepreciation" & _
	"           WHEN monthtobe + 1 = depreciationmonth_ledger" & _
	"           THEN" & _
	"              (amountdepreciation - yearlydepreciation) + yearlydepreciation" & _
	"           ELSE" & _
	"              monthlydepreciation + yearlydepreciation" & _
	"        END" & _
	"           yearlydepreciation," & _
	"        fat.fcname" & _
	"   FROM fixedasset fat" & _
	"        INNER JOIN fixedasset_depreciation_ledger fatd" & _
	"           ON (fat.fccode = fatd.fccode AND fat.fcba = fatd.fcba)" & _
	"  WHERE fat.fccode = '"& fixassetcode &"' AND fat.fcba = '"& session("fcba") &"' AND fatd.lineno = '"& lineno &"'" 

	' "  SELECT fatd.lineno," & _ 
	' "         fat.total," & _ 
	' "         fatd.yearlydepreciation," & _ 
	' "         fat.fcname" & _ 
	' "    FROM fixedasset fat" & _ 
	' "         INNER JOIN fixedasset_depreciation_ledger fatd" & _ 
	' "            ON (fat.fccode = fatd.fccode AND fat.fcba = fatd.fcba)" & _ 
	' "   WHERE fat.fccode = '"& fixedasset_depreciation_ledger &"' AND fat.fcba = '"& session("fcba") &"' AND fatd.lineno = '"& lineno &"'" & _ 
	' "   ORDER BY lineno" 

' response.Write q

rsrecurtrans.Open q, connstr
if Not rsrecurtrans.Eof then
	while Not rsrecurtrans.Eof 
	fixassetname = rsrecurtrans("fcname")
	total = rsrecurtrans("total")
	yearlydepreciation = rsrecurtrans("yearlydepreciation")
	totalyearly = cdbl(to_nol(rsrecurtrans("yearlydepreciation")))
	' dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercharge")))
	rsrecurtrans.movenext
wend 
	
End if

rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """fixassetname"" : """ & fixassetname & ""","
response.write """total"" : """ & total & ""","
response.write """totalyearly"" : """ & totalyearly & """"
' response.write """targettype_depreciation"" : """ & targettype_depreciation & ""","
' response.write """targetcode_depreciation"" : """ & targetcode_depreciation & """"
response.write "}"

%>
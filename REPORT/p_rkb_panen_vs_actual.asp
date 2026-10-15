<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtMonth, v_txtYear
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName, v_txtsectionname, v_txtGangCode, v_txtStatus
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no

v_txtYear		= Request("txtYear")
v_txtMonth		= Request("txtMonth")
v_txtGroupPlantation  = Request("txtGroupPlantation")
v_txtPlantation  	= Request("txtPlantation")
v_txtGroupRegion  	= Request("txtGroupRegion")
v_txtRegion  		= Request("txtRegion")
v_txtcompany  		= Request("txtCompany")
v_txtfcba 			= Request("txtfcba")
v_txtDivision		= Request("txtDivision")
v_txtBlok			= Request("txtBlok")
v_txtPlantingDate	= Request("txtPlantingDate")
v_txtJobcode		= Request("txtJobcode")
txtownership 		= Request("txtownership")

If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
If "" & txttdate = "" Then txttdate = textDate(date)

txttdate = txtMonth & "" & txtYear

ScriptName		= Request.ServerVariables("SCRIPT_NAME")
nSpanTotal = 0
Dim v_paymenttype : v_paymenttype = Request("paymenttype")

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManagerName = rsBaInfo("FCMANAGER")
	varKTUName = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
	varManagerName = ""
	varKTUName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

txtComp         = Request("txtCompany")
nSpanTotal = 8

'response.write "Year = " & v_txtyear & " Month = " & v_txtmonth 


'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate

Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
strFieldCons = 	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
				" ,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation " & _
			   	" ,USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion " & _
			   	" ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
			   	" ,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut," 

strFieldCons2 = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba)" & _
				" ,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba)  " & _
			   	" ,USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) " & _
			   	" ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) " & _
			   	" ,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba), fcba, BUSINESSUNIT_API.get_OrderData(fcba)," 

			   
wherecons = " where company is not null "
if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
 
end if
if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
    if v_txtPlantation <> "--ALL--" then
	   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
	end if   
	selectCons = selectCons & ",Plantation"
	orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
else
    nSpanTotal = nSpanTotal - 1 
end if
if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
    if v_txtGroupRegion <> "--ALL--" then
	   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
	end if
	selectCons = selectCons & " ,GroupRegion"
	orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
else
    nSpanTotal = nSpanTotal - 1 
end if
if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
    if v_txtRegion <> "--ALL--" then
	   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
	end if
	selectCons = selectCons & " ,Region"
	orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
else
    nSpanTotal = nSpanTotal - 1 
end if
if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
    if v_txtCompany  <> "--ALL--" then
	   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
	end if
	selectCons = selectCons & " ,company"
    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   CompanyName=rsCat("FCCOMPANYNAME")			   
    end if
else
    nSpanTotal = nSpanTotal - 1 	
end if
if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
    if v_txtfcba  <> "--ALL--" then
       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
	end if
    selectCons = selectCons & ",fcba"
    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba) "
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   fcbaname = rsCat("fcName")
	   CompanyName=rsCat("FCCOMPANYNAME")
	   MgrName = rsCat("FCMANAGER")
	   KTUName = rsCat("FCKTU")
	   CropName=rsCat("FCCROP")
	   FcbaType=rsCat("fcType")
    end if
else
    nSpanTotal = nSpanTotal - 1 
end if
 

if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
	if v_txtDivision <>  "--ALL--"  then
		whereCons = whereCons & " AND unitkerja = '" & v_txtDivision & "'"
	end if
	selectCons = selectCons &  " , unitkerja "
	orderCons  = orderCons & ",unitkerja "
else
	nSpanTotal = nSpanTotal - 1 	  
end if   



if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
	if v_txtPlantingDate <>  "--ALL--" then
		whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
	end if
	selectCons = selectCons &  " "
	orderCons  = orderCons & " "
	nSpanTotal = nSpanTotal - 1    
else
	nSpanTotal = nSpanTotal - 1    
end if


if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
   if v_txtBlok <>  "--ALL--" then
	  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
   end if
   selectCons = selectCons & " , FIELDCODE "
   orderCons  = orderCons & ",FIELDCODE "
   nSpanTotal = nSpanTotal - 1    
else
nSpanTotal = nSpanTotal - 1    
end if

if v_txtJobcode <> "--SUMMARY--" and v_txtJobcode <> "" then
   if v_txtJobcode <>  "--ALL--" then
	  whereCons = whereCons & " AND activitycode = '" & v_txtJobcode & "'"
   end if
   selectCons = selectCons & " , activitycode "
   orderCons  = orderCons & ",activitycode "
   nSpanTotal = nSpanTotal - 1    
else
nSpanTotal = nSpanTotal - 1    
end if

		strMain	=	"  SELECT"& strFieldCons &"" & _
					"                    rkbno," & _
					"                    unitkerja," & _
					"                    period_year," & _
					"                    period_month," & _
					"                    activitycode," & _
					"                    activityname," & _
					"                    plantingyear," & _
					"                    fieldcode," & _
					"                    SUM (luas_ha) luas_ha," & _
					"                    SUM (plantingpoint) plantingpoint," & _
					"                    SUM (totalstand) totalstand," & _
					"                    satuan," & _
					"                    MAX (harvestingbased_abw_budget) harvestingbased_abw_budget," & _
					"                    rotasi," & _
					"                    MAX (bunch) bunch," & _
					"                    SUM (kg_budget) kg_budget," & _
					"                    SUM (jjg_actual) bunch_actual," & _
					"                    SUM (kg_actual) kg_actual," & _
					"                    MAX (bjr_actual) bjr_actual," & _
					"                    CASE" & _
					"                       WHEN SUM (kg_actual) = 0 THEN 0" & _
					"                       ELSE (SUM (jjg_actual) / SUM (kg_actual) * 100)" & _
					"                    END" & _
					"                       produksi_percent," & _
					"                    norma_hk," & _
					"                    SUM (totalhk) totalhk," & _
					"                    SUM (rp_hk) rp_hk," & _
					"                    SUM (mandays) hk_actual," & _
					"                    (SUM (mandays) / SUM (totalhk) * 100) mandays_percent," & _
					"                    SUM (basis_jjg) basis_jjg," & _
					"                    percent_sb," & _
					"                    SUM (premi_basis) premi_basis," & _
					"                    percent_lb," & _
					"                    SUM (premi_lebih_basis) premi_lebih_basis," & _
					"                    (SUM (premi_basis) + SUM (premi_lebih_basis)) total_premi," & _
					"                    (SUM (rp_hk) + (SUM (premi_basis) + SUM (premi_lebih_basis)))" & _
					"                       total," & _
					"                    (  (  SUM (rp_hk) + (SUM (premi_basis) + SUM (premi_lebih_basis) ) ) / SUM (kg_budget))" & _
					"                       total_rp_kg" & _
					"               FROM (SELECT mwp.fcba," & _
					"                            mwp.rkbno," & _
					"                            mwp.unitkerja," & _
					"                            mwp.period_year," & _
					"                            mwp.period_month," & _
					"                            mwp.activitycode," & _
					"                            mwp.activityname," & _
					"                            mwp.plantingyear," & _
					"                            mwp.fieldcode," & _
					"                            mwp.luas_ha," & _
					"                            mwp.plantingpoint," & _
					"                            mwp.totalstand," & _
					"                            mwp.satuan," & _
					"                            mwp.harvestingbased_abw_budget," & _
					"                            NVL (anh.rotasi, 0) rotasi," & _
					"                            mwp.bunch," & _
					"                            NVL (hv.bunch, 0) jjg_actual," & _
					"                            CASE" & _
					"                               WHEN NVL (hv.bunch, 0) = 0" & _
					"                               THEN" & _
					" 0" & _
					"                               ELSE" & _
					"                                  (  NVL (hv.bunch_estateweight, 0)" & _
					"                                   / NVL (hv.bunch, 0))" & _
					"                            END" & _
					"                               bjr_actual," & _
					"                            mwp.kg_budget," & _
					"                            NVL (hv.bunch_estateweight, 0) kg_actual," & _
					"                            NVL (anh.norma_hk, 0) norma_hk," & _
					"                            ROUND (" & _
					"                               (  NVL (mwp.luas_ha, 0)" & _
					"                                * NVL (anh.rotasi, 0)" & _
					"                                * NVL (anh.norma_hk, 0))," & _
					"                               2)" & _
					"                               totalhk," & _
					"                            ROUND (" & _
					"                               (  (mwp.luas_ha * anh.rotasi * anh.norma_hk)" & _
					"                                * anh.rate1)," & _
					"                               0)" & _
					"                               rp_hk," & _
					"                            NVL (ag.mandays, 0) mandays," & _
					"                            anh.basis_jjg," & _
					"                            mwp.percent_sb," & _
					"                            ROUND (" & _
					"                               (  anh.rate_sb" & _
					"                                * (mwp.percent_sb / 100)" & _
					"                                * (  NVL (mwp.luas_ha, 0)" & _
					"                                   * NVL (anh.rotasi, 0)" & _
					"                                   * NVL (anh.norma_hk, 0)))," & _
					"                               2)" & _
					"                               premi_basis," & _
					"                            mwp.percent_lb," & _
					"                            (  NVL (mwp.bunch, 0)" & _
					"                             * (mwp.percent_lb / 100)" & _
					"                             * NVL (anh.lb_rate, 0))" & _
					"                               premi_lebih_basis" & _
					"                       FROM (SELECT mwp.fcba," & _
					"                                    mwp.rkbno," & _
					"                                    mwd.unitkerja," & _
					"                                    mwp.period_year," & _
					"                                    mwp.period_month," & _
					"                                    mwp.activitycode," & _
					"                                    job_api.get_jobcode_name (mwp.activitycode," & _
					"                                                              mwp.fcba)" & _
					"                                       activityname," & _
					"                                    field_api.getvalid_plantingyear (" & _
					"                                       mwd.fieldcode," & _
					"                                       mwd.fcba," & _
					"                                       TO_DATE ('"& v_txtMonth &"/01/"& v_txtYear &"', 'mm/dd/yyyy'))" & _
					"                                       plantingyear," & _
					"                                    mwd.fieldcode," & _
					"                                    field_api.getvalid_ha (" & _
					"                                       mwd.fieldcode," & _
					"                                       mwd.fcba," & _
					"                                       TO_DATE ('"& v_txtMonth &"/01/"& v_txtYear &"', 'mm/dd/yyyy'))" & _
					"                                       luas_ha," & _
					"                                    field_api.getvalid_plantingpoint (" & _
					"                                       mwd.fieldcode," & _
					"                                       mwd.fcba," & _
					"                                       TO_DATE ('"& v_txtMonth &"/01/"& v_txtYear &"', 'mm/dd/yyyy'))" & _
					"                                       plantingpoint," & _
					"                                    field_api.getvalid_pokok (" & _
					"                                       mwd.fieldcode," & _
					"                                       mwd.fcba," & _
					"                                       TO_DATE ('"& v_txtMonth &"/01/"& v_txtYear &"', 'mm/dd/yyyy'))" & _
					"                                       totalstand," & _
					"                                    'KG' satuan," & _
					"                                    mwd.harvestingbased_abw_budget," & _
					"                                    mwd.kg_budget," & _
					"                                    ROUND (" & _
					"                                       (  NVL (mwd.kg_budget, 0)" & _
					"                                        / NVL (mwd.harvestingbased_abw_budget, 0))," & _
					"                                       0)" & _
					"                                       bunch," & _
					"                                    NVL (mwd.percent_lb, 0) percent_lb," & _
					"                                    NVL (mwd.percent_sb, 0) percent_sb" & _
					"                               FROM monthlyworkplan_panen mwp" & _
					"                                    INNER JOIN monthlyworkplan_panendet mwd" & _
					"                                       ON (    mwp.fcba = mwd.fcba" & _
					"                                           AND mwp.rkbno = mwd.rkbno)) mwp" & _
					"                            LEFT JOIN activity_norm_harv anh" & _
					"                               ON (    mwp.fcba = anh.fcba" & _
					"                                   AND mwp.plantingyear = anh.plantingyear)" & _
					"                            LEFT JOIN harvestingspb hv" & _
					"                               ON (    mwp.fcba = hv.fcba" & _
					"                                   AND mwp.fieldcode = hv.fieldcode" & _
					"                                   AND TO_NUMBER (" & _
					"                                          TO_CHAR (hv.receptiondate, 'mm')) =" & _
					"                                          mwp.period_month" & _
					"                                   AND TO_NUMBER (" & _
					"                                          TO_CHAR (hv.receptiondate, 'yyyy')) =" & _
					"                                          mwp.period_year)" & _
					"                            LEFT JOIN attendance_gad ag" & _
					"                               ON (    mwp.fcba = ag.fcba" & _
					"                                   AND mwp.fieldcode = ag.locationcode" & _
					"                                   AND TO_NUMBER (TO_CHAR (ag.fddate, 'mm')) =" & _
					"                                          mwp.period_month" & _
					"                                   AND TO_NUMBER (TO_CHAR (ag.fddate, 'yyyy')) =" & _
					"                                          mwp.period_year)" & _
					"                      WHERE mwp.period_month = '"& v_txtMonth &"' AND mwp.period_year = '"& v_txtYear &"')" & _
					"           GROUP BY "& strFieldCons2 &"" & _ 
					"                    rkbno," & _
					"                    unitkerja," & _
					"                    period_year," & _
					"                    period_month," & _
					"                    activitycode," & _
					"                    activityname," & _
					"                    plantingyear," & _
					"                    fieldcode," & _
					"                    satuan," & _
					"                    rotasi," & _
					"                    norma_hk," & _
					"                    percent_sb," & _
					"                    percent_lb" 

		strEPS =	"		  ,rkbno," & _
					"          unitkerja," & _
					"          period_year," & _
					"          period_month," & _
					"          activitycode," & _
					"          activityname," & _
					"          plantingyear," & _
					"          fieldcode," & _
					"          luas_ha," & _
					"          plantingpoint," & _
					"          totalstand," & _
					"          satuan," & _
					"          harvestingbased_abw_budget," & _
					"          rotasi," & _
					"          bunch," & _
					"          kg_budget," & _
					"          bunch_actual," & _
					"          kg_actual," & _
					"          bjr_actual," & _
					"          produksi_percent," & _
					"          norma_hk," & _
					"          totalhk," & _
					"          rp_hk," & _
					"          hk_actual," & _
					"          mandays_percent," & _
					"          basis_jjg," & _
					"          percent_sb," & _
					"          premi_basis," & _
					"          percent_lb," & _
					"          premi_lebih_basis," & _
					"          total_premi," & _
					"          total," & _
					"          total_rp_kg"
		
	strPayroll = "select GroupPlantation " & selectCons & strEPS &"  From (" & strMain & " )" & whereCons & " "

	strPayroll = strPayroll & " order by plantingyear " 
	' response.write strPayroll
	Set rsPayroll = Server.CreateObject("ADODB.Recordset")
	rsPayroll.Open strPayroll, connstr, adOpenStatic
			
%>
<html>
<head>
	<title>ePlaS Reporting - Rencana Kerja Bulanan Panen </title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/jquery-ui.css" />
	<script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);
	</script>
	<script language="JavaScript" src="/js/date-picker.js"></script> 
	<script language="JavaScript">
	var newwin;
	function launchwin(winurl,winname,winfeatures)
	{
		//This launches a new window and then
		//focuses it if window.focus() is supported.
		newwin = window.open(winurl,winname,winfeatures);

	}
	</script>
	
<style type="text/css">
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
			<div class="printhide">
			<input type="button" value="&#xf02f;" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmAR" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include file="FormSelectNewsum.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<th>TAHUN</th>
						<th>BULAN</th> 
						<th>PEKERJAAN</th>
						<th>UNIT KERJA</th>
						<th>TAHUN TANAM</th>
						<th>BLOK</th>
					</tr>
					<tr>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM

							Yr = v_txtYear
							MinYr = 2015
							MaxYr = Yr  + 2
						%>
								<td width="4">
									<select name="txtYear">
						<%
							For i = MinYr to MaxYr
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
						<%
							Next
						%>
									</select>
								</td> 
								<td width="2">
									<select name="txtMonth">
						<% 
							For i = 1 to 12
								
								if i <= 9 then 
									iBulan = 0&i
								else
									iBulan = i
								end if
								i	= iBulan
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
						<%
							Next
						%>
									</select>
								</td>
						<td align="center">
							<select name="txtJobcode">	
								<option value="--SUMMARY--">--SUMMARY--</option> 
								<% 
									Set rsDivision = Server.CreateObject("ADODB.Recordset")
									rsDivision.CursorLocation = adUseClient
									

							
									if len(trim(v_txtJobcode)) > 0 then 
										strsql =  	" SELECT '--ALL--' as fccode,'' fcName FROM DUAL " & _
													" UNION ALL SELECT fccode, fcname FROM job WHERE fcba = '"& v_txtFcba &"'" & _
													" AND job_forfieldstatus NOT IN ('BALANCE SHEET','BIBITAN')" & _
													" AND jobgroupcode in (select fccode from jobgroup where fcba = '"& v_txtFcba &"'" & _
													" AND activity_heading = 'PNN')" & _
													" ORDER BY fccode"  
									else
										strsql =  	" SELECT '--ALL--' as fccode,'' fcName FROM DUAL " & _
													" UNION ALL SELECT fccode, fcname FROM job WHERE fcba = '"& Session("fcba") &"'" & _
													" AND job_forfieldstatus NOT IN ('BALANCE SHEET','BIBITAN')" & _
													" AND jobgroupcode in (select fccode from jobgroup where fcba = '"& Session("fcba") &"'" & _
													" AND activity_heading = 'PNN')" & _
													" ORDER BY fccode"  
									end if
									
									
									rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsDivision.eof then
										DO while not rsDivision.eof
										if v_txtJobcode = rsDivision("fccode") then
								%>			<option value="<%=rsDivision("fccode")%>" Selected><%=rsDivision("fccode")%> : <%=rsDivision("fcname")%></option>
										<%else%>
											<option value="<%=rsDivision("fccode")%>"><%=rsDivision("fccode")%> : <%=rsDivision("fcname")%></option>
										<%end if%>
								<%			rsDivision.movenext
										loop
									End if
									rsDivision.Close
									set rsDivision = nothing
								%>
							</select>		
						</td>
						<td align="center">
							<select name="txtDivision">	
								<option value="--SUMMARY--">--SUMMARY--</option> 
								<%
									Dim rsDivision
									Set rsDivision = Server.CreateObject("ADODB.Recordset")
									rsDivision.CursorLocation = adUseClient
									

							
									if len(trim(v_txtDivision)) > 0 then 
										strsql =  	" SELECT '--ALL--' as sectionname, '' fcName from DUAL UNION ALL SELECT DISTINCT emp.sectionname, sc.fcname" & _
													"    FROM employee emp" & _
													"         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
													"   WHERE fcba = '"& v_txtFcba &"' order by sectionname" 
									else
										strsql =  	" SELECT '--ALL--' as sectionname, '' fcName FROM DUAL UNION ALL SELECT DISTINCT emp.sectionname, sc.fcname" & _
													"    FROM employee emp" & _
													"         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
													"   WHERE fcba = '"& Session("fcba") &"' order by sectionname"
									end if
									
									
									rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsDivision.eof then
										DO while not rsDivision.eof
										if v_txtDivision = rsDivision("sectionname") then
								%>			<option value="<%=rsDivision("sectionname")%>" Selected><%=rsDivision("sectionname")%> : <%=rsDivision("fcname")%></option>
										<%else%>
											<option value="<%=rsDivision("sectionname")%>"><%=rsDivision("sectionname")%> : <%=rsDivision("fcname")%></option>
										<%end if%>
								<%			rsDivision.movenext
										loop
									End if
									rsDivision.Close
									set rsDivision = nothing
								%>
							</select>		
						</td>
						<td align="center">
							<select name="txtPlantingDate">	
									<option value="--SUMMARY--">--SUMMARY--</option>
									<%
										Dim rsPlantingDate
										Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
										rsPlantingDate.CursorLocation = adUseClient
								
										if len(trim(v_txtPlantingDate)) > 0 then 
											strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
										else
											strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
										end if
										
										
										rsPlantingDate.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
										if not rsPlantingDate.eof then
											DO while not rsPlantingDate.eof
											if v_txtPlantingDate = rsPlantingDate("PlantingDate") then
									%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
											<%else%>
												<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
											<%end if%>
									<%			rsPlantingDate.movenext
											loop
										End if
										rsPlantingDate.Close
										set rsPlantingDate = nothing
									%>
							</select>		
						</td>
						<td align="center">
							<select name="txtBlok">	
								<option value="--SUMMARY--">--SUMMARY--</option>
								<%
									Dim rsBlok
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.CursorLocation = adUseClient

							
									if len(trim(v_txtBlok)) > 0 then 
										strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
									else
										strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
									end if
									
									rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsBlok.eof then
										DO while not rsBlok.eof
										if v_txtBlok = rsBlok("Blok") then
								%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%></option>
										<%else%>
											<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%></option>
										<%end if%>
								<%			rsBlok.movenext
										loop
									End if
									rsBlok.Close
									set rsBlok = nothing
								%>
							</select>		
						</td>
						<td align="center">
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>
		
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="31" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<h4>PT. <%=varCompanyName%></h4>
						<h5><%=varFcbaName%></h5>
					</div>
					<div id="subheader" class="centered">
						<h2>LAPORAN RENCANA KERJA BULANAN VS AKTUAL PANEN </h2>
						<br/>
						<h4>
							PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%>
						</h4> 

					</div>
					</div>
				</th>
			</tr>
		<tr> 
			<th rowspan="3" style="color:#fff;">NO</th>
			<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
				<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
			<% 			
		    end if %>
			
			<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
			<% 			
		    end if %>
			
			<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
			<%
		    end if %>
			
			<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
			<% 		
		    end if %>
			
			<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
			<% 		
		    end if %>
			
			<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;">Unit Kerja</th>
			<% 		
		    end if %>
			
			<th rowspan="3" style="color:#fff;">Tahun Tanam</th>
			<th rowspan="3" style="color:#fff;">Nomor Dokumen</th> 
			<th rowspan="3" style="color:#fff;">Kode Pekerjaan</th>
			<th rowspan="3" style="color:#fff;">Nama Pekerjaan</th>
			<th colspan="3" style="color:#fff;">Blok</th>
			<th rowspan="3" style="color:#fff;">Satuan</th>
			<th rowspan="3" style="color:#fff;">BJR</th> 
			<th rowspan="3" style="color:#fff;">Rotasi</th>
			<th colspan="6" style="color:#fff;">Pencapaian Hasil</th>
			<th colspan="5" style="color:#fff;">Tenaga Kerja</th>
			<th colspan="6" style="color:#fff;">Panen</th>
			<th colspan="2" style="color:#fff;">Total Biaya Panen</th>
		</tr>
		<tr>	
			<th rowspan="2" style="color:#fff;">Blok</th>
			<th rowspan="2" style="color:#fff;">Luas(Ha)</th>
			<th rowspan="2" style="color:#fff;">Jumlah Pokok</th> 
			<th colspan="2" style="color:#fff;">RKB</th> 
			<th colspan="3" style="color:#fff;">AKT</th> 
			<th rowspan="2" style="color:#fff;">Percent (%)</th> 
			<th colspan="3" style="color:#fff;">RKB</th> 
			<th colspan="2" style="color:#fff;">AKT</th> 
			<th rowspan="2" style="color:#fff;">JJG Basis</th> 
			<th rowspan="2" style="color:#fff;">(%) SB</th> 
			<th rowspan="2" style="color:#fff;">Premi Basis (Rp)</th> 
			<th rowspan="2" style="color:#fff;">(%) LB</th> 
			<th rowspan="2" style="color:#fff;">Premi Lebih Basis (Rp)</th> 
			<th rowspan="2" style="color:#fff;">Total Premi</th> 
			<th rowspan="2" style="color:#fff;">Total</th> 
			<th rowspan="2" style="color:#fff;">RP/KG</th> 
		</tr>
		<tr>
			<th style="color:#fff;">KG</th> 
			<th style="color:#fff;">JJG</th> 
			<th style="color:#fff;">JJG Kirim</th> 
			<th style="color:#fff;">BJR</th> 
			<th style="color:#fff;">KG Aktual</th> 
			<th style="color:#fff;">Norma HK</th> 
			<th style="color:#fff;">Total HK</th> 
			<th style="color:#fff;">Total HK(RP)</th> 
			<th style="color:#fff;">Total HK Aktual</th> 
			<th style="color:#fff;">Percent (%)</th> 
		</tr>
	</thead>
	<tbody>
<%if rsPayroll.BOF and rsPayroll.EOF then%>
		<tr>
			<td colspan="25" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 0
		rows = 1
		while not rsPayroll.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
		
		' totalbiaya = cdbl(to_nol(rsPayroll("totalamount"))) + cdbl(to_nol(rsPayroll("boronganpremi"))) + cdbl(to_nol(rsPayroll("totalamountstock"))) 
		' If cdbl(to_nol(rsPayroll("volume"))) = 0 Then costsatuan = 0 Else costsatuan = cdbl(to_nol(totalbiaya)) / cdbl(to_nol(rsPayroll("volume")))
%>
		<tr class="<%=cClass%>">
			<% if rows = 1 then
					sqlrow = "  select count(plantingyear) rowsd  from ( " & strPayroll & " ) where plantingyear = '" & rsPayroll("plantingyear") & "'"
					' response.write sqlrow
					Set rsrow = Server.CreateObject("ADODB.Recordset")
					rsrow.Open sqlrow, connstr, adOpenStatic
					if not rsrow.EOF then
						temprows = CDbl(rsrow("rowsd"))
					else
						temprows = 0
					end if
				end if 
			no = NO+1
			%>
			<td align="center"><%=No%></td>
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% tempPlantation = rsPayroll("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsPayroll("Plantation") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsPayroll("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsPayroll("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsPayroll("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsPayroll("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsPayroll("GroupRegion") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsPayroll("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsPayroll("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsPayroll("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsPayroll("Region") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsPayroll("region")%>
							</a>
						</td>
						<% SprevRegion = rsPayroll("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPayroll("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsPayroll("Company")
					   tempfcba = "--ALL--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsPayroll("Company") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsPayroll("company")%>
							</a>
						</td>
						<% SprevCompany = rsPayroll("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPayroll("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsPayroll("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsPayroll("fcba")
					   tempownership = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsPayroll("fcba") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsPayroll("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsPayroll("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				 
				<% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> "" then %>
				   <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPayroll("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsPayroll("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsPayroll("fcba")
					   else
					       tempfcba = ""
					   end if
					   if txtownership <> "--SUMMARY--"  and txtownership <> "" then
					      tempownership = rsPayroll("ownership")
					   else
					      tempownership = ""
                       end if
					   if v_txtDivision <> "--SUMMARY--"  and v_txtDivision <> "" then
					      tempDivision = rsPayroll("unitkerja")
					   else
					      tempDivision = ""
                       end if					   
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevDivision <> rsPayroll("unitkerja") then %>
					    <td align="center"> 
							<%=rsPayroll("unitkerja")%> 
						</td>
						<% sprevDivision = rsPayroll("unitkerja") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %> 
			<%if rows = 1 then%>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsPayroll("plantingyear")%></td> 
			<%end if%>
			<td align="center"><%=rsPayroll("rkbno")%></td>
			<td align="center"><%=rsPayroll("activitycode")%></td>
			<td align="left"><%=rsPayroll("activityname")%></td> 
			<td align="center"><%=rsPayroll("fieldcode")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("luas_ha")), 2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("totalstand")),2)%></td>
			<td align="center"><%=rsPayroll("satuan")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("harvestingbased_abw_budget")),2)%></td> 
			<td align="center"><%=rsPayroll("rotasi")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("kg_budget")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("bunch")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("bunch_actual")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("bjr_actual")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("kg_actual")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("produksi_percent")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("norma_hk")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("totalhk"))),2)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("rp_hk"))),2)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("hk_actual"))),2)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("mandays_percent"))),2)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("basis_jjg"))),2)%></td>
			<td align="left"><%=rsPayroll("percent_sb")%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("premi_basis"))),2)%></td>
			<td align="center"><%=rsPayroll("percent_lb")%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("premi_lebih_basis"))),2)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("total_premi"))),2)%></td>
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("total"))),2)%></td> 
			<td align="right"><%=formatnumber( cdbl(to_nol(rsPayroll("total_rp_kg"))),2)%></td> 
		</tr>
<%	if rows = temprows then
			noline			= noline		+ 1
			nBunch = nBunch + cdbl(rsPayroll("Bunch"))
			nkg_budget 		= nkg_budget 	+ cdbl(rsPayroll("kg_budget"))
			nbunch_actual 		= nbunch_actual 	+ cdbl(rsPayroll("bunch_actual"))
			nkg_actual 		= nkg_actual 	+ cdbl(rsPayroll("kg_actual"))
			nproduksi_percent 		= cdbl(to_nol(nkg_actual)) / cdbl(to_nol(nkg_budget)) * 100 
			nnorma_hk 	= nnorma_hk	+ cdbl(rsPayroll("norma_hk"))
			ntotalhk 	= ntotalhk + cdbl(rsPayroll("totalhk"))
			nrp_hk 	= nrp_hk + cdbl(to_nol(rsPayroll("rp_hk")))
			nhk_actual 	= nhk_actual + cdbl(to_nol(rsPayroll("hk_actual")))
			nmandays_percent 	= cdbl(to_nol(nhk_actual)) / cdbl(to_nol(ntotalhk)) * 100
			nbasis_jjg 	= nbasis_jjg + cdbl(to_nol(rsPayroll("basis_jjg")))
			npremi_basis = npremi_basis + cdbl(to_nol(rsPayroll("premi_basis")))
			npremi_lebih_basis = npremi_lebih_basis	+ cdbl(to_nol(rsPayroll("premi_lebih_basis"))) 
			ntotal_premi = ntotal_premi + cdbl(to_nol(rsPayroll("total_premi"))) 
			ntotal = ntotal	+ cdbl(to_nol(rsPayroll("Total"))) 
			ntotal_rp_kg = (  (  cdbl(to_nol(nrp_hk)) + (cdbl(to_nol(npremi_basis)) + cdbl(to_nol(npremi_lebih_basis)) ) ) / cdbl(to_nol(nkg_budget)))
			rows = 1
			
	%>
		<tr style="background-color:#F4F4A0;">	 
			<td align="center" colspan="<%= 12+ nSpanTotal %>" ><strong>Sub Total</strong></td> 
			<td align="right"><strong><%=formatnumber(nkg_budget,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nBunch,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nbunch_actual,2)%></strong></td> 
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(nkg_actual,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nproduksi_percent,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nnorma_hk,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(ntotalhk,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nrp_hk,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nhk_actual,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nmandays_percent,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(nbasis_jjg,2)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(npremi_basis,2)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(npremi_lebih_basis,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(ntotal_premi,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(ntotal,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(ntotal_rp_kg,2)%></strong></td>
		</tr>
		<tr>
			<td align="center" colspan="<%= 24+ nSpanTotal %>" ></td>
		</tr>
		<%	
			noline = 1
			nBunch = cdbl("0")
			nkg_budget = cdbl("0")
			nkg_actual = cdbl("0")
			nbunch_actual = cdbl("0")
			nnorma_hk = cdbl("0")
			ntotalhk = cdbl("0")
			nrp_hk = cdbl("0")
			nhk_actual = cdbl("0")
			nbasis_jjg = cdbl("0")
			nTotalamountstock = cdbl("0")
			npremi_basis = cdbl("0")
			npremi_lebih_basis = cdbl("0")
			ntotal_premi = cdbl("0")
			ntotal = cdbl("0")
			ntotal_rp_kg = cdbl("0")
			nmandays_percent = cdbl("0")
			nproduksi_percent = cdbl("0")
			 
		else
		' cari sub total
			noline			= noline		+ 1
			nBunch = nBunch + cdbl(rsPayroll("Bunch"))
			nkg_budget 		= nkg_budget 	+ cdbl(rsPayroll("kg_budget"))
			nbunch_actual 		= nbunch_actual 	+ cdbl(rsPayroll("bunch_actual"))
			nkg_actual 		= nkg_actual 	+ cdbl(rsPayroll("kg_actual"))
			nproduksi_percent 		= cdbl(to_nol(nkg_actual)) / cdbl(to_nol(nkg_budget)) * 100 
			nnorma_hk 	= nnorma_hk	+ cdbl(rsPayroll("norma_hk"))
			ntotalhk 	= ntotalhk + cdbl(rsPayroll("totalhk"))
			nrp_hk 	= nrp_hk + cdbl(to_nol(rsPayroll("rp_hk")))
			nhk_actual 	= nhk_actual + cdbl(to_nol(rsPayroll("hk_actual")))
			nmandays_percent 	= cdbl(to_nol(nhk_actual)) / cdbl(to_nol(ntotalhk)) * 100
			nbasis_jjg 	= nbasis_jjg + cdbl(to_nol(rsPayroll("basis_jjg")))
			npremi_basis = npremi_basis + cdbl(to_nol(rsPayroll("premi_basis")))
			npremi_lebih_basis = npremi_lebih_basis	+ cdbl(to_nol(rsPayroll("premi_lebih_basis"))) 
			ntotal_premi = ntotal_premi + cdbl(to_nol(rsPayroll("total_premi"))) 
			ntotal = ntotal	+ cdbl(to_nol(rsPayroll("Total"))) 
			ntotal_rp_kg = (  (  cdbl(to_nol(nrp_hk)) + (cdbl(to_nol(npremi_basis)) + cdbl(to_nol(npremi_lebih_basis)) ) ) / cdbl(to_nol(nkg_budget)))
			rows = rows + 1
		End if
							  
			' cari grand total 
			gtBunch = gtBunch + cdbl(rsPayroll("Bunch"))
			gtkg_budget = gtkg_budget + cdbl(rsPayroll("kg_budget"))
			gtkg_actual = gtkg_actual + cdbl(rsPayroll("kg_actual"))
			gtbunch_actual = gtbunch_actual + cdbl(rsPayroll("bunch_actual"))
			gtproduksi_percent = cdbl(to_nol(gtkg_actual)) / cdbl(to_nol(gtkg_budget)) * 100 
			gtnorma_hk 	= gtnorma_hk + cdbl(rsPayroll("norma_hk"))
			gttotalhk 	= gttotalhk	+ cdbl(to_nol(rsPayroll("totalhk")))
			gtrp_hk 	= gtrp_hk + cdbl(to_nol(rsPayroll("rp_hk")))
			gthk_actual 	= gthk_actual + cdbl(to_nol(rsPayroll("hk_actual")))
			gtmandays_percent 	= cdbl(to_nol(gthk_actual)) / cdbl(to_nol(gttotalhk)) * 100
			gtbasis_jjg 	= gtbasis_jjg + cdbl(to_nol(rsPayroll("basis_jjg")))
			gtpremi_basis = gtpremi_basis + cdbl(to_nol(rsPayroll("premi_basis")))
			gtpremi_lebih_basis = gtpremi_lebih_basis + cdbl(to_nol(rsPayroll("premi_lebih_basis")))
			gttotal_premi = gttotal_premi + cdbl(to_nol(rsPayroll("total_premi")))
			gttotal = gttotal + cdbl(to_nol(rsPayroll("total"))) 
			gttotal_rp_kg = (  (  cdbl(to_nol(gtrp_hk)) + (cdbl(to_nol(gtpremi_basis)) + cdbl(to_nol(gtpremi_lebih_basis)) ) ) / cdbl(to_nol(gtkg_budget)))
		 
			rsPayroll.MoveNext			
			Wend
		rsPayroll.Close
		set rsPayroll = Nothing
	%> 
		<tr style="background-color:#27AE60;">	 
			<td align="center" colspan="<%= 12+ nSpanTotal %>" ><strong>Grand Total</strong></td> 
			<td align="right"><strong><%=formatnumber(gtkg_budget,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(gtBunch,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gtbunch_actual,2)%></strong></td> 
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(gtkg_actual,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(gtproduksi_percent,2)%></strong></td> 
			<td align="right"><strong><%=formatnumber(gtnorma_hk,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gttotalhk,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gtrp_hk,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gthk_actual,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gtmandays_percent,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gtbasis_jjg,2)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(gtpremi_basis,2)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(gtpremi_lebih_basis,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gttotal_premi,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gttotal,2)%></strong></td>
			<td align="right"><strong><%=formatnumber(gttotal_rp_kg,2)%></strong></td>
		</tr> 
	</tbody>
</table> 
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
   </body>
   </html>
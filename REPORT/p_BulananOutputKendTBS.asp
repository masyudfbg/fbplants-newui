<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtVehicleGroup
Dim ScriptName
Dim rsBOKPT, strBOKPT, strBOKPT2, strMain
Dim cClass, nday, LoopNo

Dim var_colspan : var_colspan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtYear 				= Request("txtYear")
v_txtMonth 				= Request("txtMonth")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where company is not null "


If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba, VEHICLECODE"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut, VEHICLECODE"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear	=	Year(Session("v_Costbook"))
End If

If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

v_txtVehicleGroup	=	Request("txtVehicleGroup")

If Len(Trim(v_txtVehicleGroup)) = 0 Then
	v_txtVehicleGroup	=	"--ALL--"
End If

If Len(Trim(v_txtVehicleGroup)) > 0 Then
	if v_txtVehicleGroup = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtVehicleGroup = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And  VEHICLEGROUPCODE = '" & v_txtVehicleGroup & "'"
		orderCons	=	orderCons	& " "
	End If
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT " & strFieldCons & "," & _
			"         VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"         , VEHICLECODE" & _
			"         , SUM(DECODE(NMONTH, 1, JUMLAHUNIT, 0)) JUMLAHUNIT_1" & _
			"         , SUM(DECODE(NMONTH, 1, UNITKERJA, 0)) UNITKERJA_1" & _
			"         , SUM(DECODE(NMONTH, 1, KMHM, 0)) KMHM_1" & _
			"         , SUM(DECODE(NMONTH, 1, TONASE, 0)) TONASE_1" & _
			"         , SUM(DECODE(NMONTH, 2, JUMLAHUNIT, 0)) JUMLAHUNIT_2" & _
			"         , SUM(DECODE(NMONTH, 2, UNITKERJA, 0)) UNITKERJA_2" & _
			"         , SUM(DECODE(NMONTH, 2, KMHM, 0)) KMHM_2" & _
			"         , SUM(DECODE(NMONTH, 2, TONASE, 0)) TONASE_2" & _
			"         , SUM(DECODE(NMONTH, 3, JUMLAHUNIT, 0)) JUMLAHUNIT_3" & _
			"         , SUM(DECODE(NMONTH, 3, UNITKERJA, 0)) UNITKERJA_3" & _
			"         , SUM(DECODE(NMONTH, 3, KMHM, 0)) KMHM_3" & _
			"         , SUM(DECODE(NMONTH, 3, TONASE, 0)) TONASE_3" & _
			"         , SUM(DECODE(NMONTH, 4, JUMLAHUNIT, 0)) JUMLAHUNIT_4" & _
			"         , SUM(DECODE(NMONTH, 4, UNITKERJA, 0)) UNITKERJA_4" & _
			"         , SUM(DECODE(NMONTH, 4, KMHM, 0)) KMHM_4" & _
			"         , SUM(DECODE(NMONTH, 4, TONASE, 0)) TONASE_4" & _
			"         , SUM(DECODE(NMONTH, 5, JUMLAHUNIT, 0)) JUMLAHUNIT_5" & _
			"         , SUM(DECODE(NMONTH, 5, UNITKERJA, 0)) UNITKERJA_5" & _
			"         , SUM(DECODE(NMONTH, 5, KMHM, 0)) KMHM_5" & _
			"         , SUM(DECODE(NMONTH, 5, TONASE, 0)) TONASE_5" & _
			"         , SUM(DECODE(NMONTH, 6, JUMLAHUNIT, 0)) JUMLAHUNIT_6" & _
			"         , SUM(DECODE(NMONTH, 6, UNITKERJA, 0)) UNITKERJA_6" & _
			"         , SUM(DECODE(NMONTH, 6, KMHM, 0)) KMHM_6" & _
			"         , SUM(DECODE(NMONTH, 6, TONASE, 0)) TONASE_6" & _
			"         , SUM(DECODE(NMONTH, 7, JUMLAHUNIT, 0)) JUMLAHUNIT_7" & _
			"         , SUM(DECODE(NMONTH, 7, UNITKERJA, 0)) UNITKERJA_7" & _
			"         , SUM(DECODE(NMONTH, 7, KMHM, 0)) KMHM_7" & _
			"         , SUM(DECODE(NMONTH, 7, TONASE, 0)) TONASE_7" & _
			"         , SUM(DECODE(NMONTH, 8, JUMLAHUNIT, 0)) JUMLAHUNIT_8" & _
			"         , SUM(DECODE(NMONTH, 8, UNITKERJA, 0)) UNITKERJA_8" & _
			"         , SUM(DECODE(NMONTH, 8, KMHM, 0)) KMHM_8" & _
			"         , SUM(DECODE(NMONTH, 8, TONASE, 0)) TONASE_8" & _
			"         , SUM(DECODE(NMONTH, 9, JUMLAHUNIT, 0)) JUMLAHUNIT_9" & _
			"         , SUM(DECODE(NMONTH, 9, UNITKERJA, 0)) UNITKERJA_9" & _
			"         , SUM(DECODE(NMONTH, 9, KMHM, 0)) KMHM_9" & _
			"         , SUM(DECODE(NMONTH, 9, TONASE, 0)) TONASE_9" & _
			"         , SUM(DECODE(NMONTH, 10, JUMLAHUNIT, 0)) JUMLAHUNIT_10" & _
			"         , SUM(DECODE(NMONTH, 10, UNITKERJA, 0)) UNITKERJA_10" & _
			"         , SUM(DECODE(NMONTH, 10, KMHM, 0)) KMHM_10" & _
			"         , SUM(DECODE(NMONTH, 10, TONASE, 0)) TONASE_10" & _
			"         , SUM(DECODE(NMONTH, 11, JUMLAHUNIT, 0)) JUMLAHUNIT_11" & _
			"         , SUM(DECODE(NMONTH, 11, UNITKERJA, 0)) UNITKERJA_11" & _
			"         , SUM(DECODE(NMONTH, 11, KMHM, 0)) KMHM_11" & _
			"         , SUM(DECODE(NMONTH, 11, TONASE, 0)) TONASE_11" & _
			"         , SUM(DECODE(NMONTH, 12, JUMLAHUNIT, 0)) JUMLAHUNIT_12" & _
			"         , SUM(DECODE(NMONTH, 12, UNITKERJA, 0)) UNITKERJA_12" & _
			"         , SUM(DECODE(NMONTH, 12, KMHM, 0)) KMHM_12" & _
			"         , SUM(DECODE(NMONTH, 12, TONASE, 0)) TONASE_12" & _
			"         , SUM(DECODE(NYEAR, " & v_txtYear & ", JUMLAHUNIT, 0)) JUMLAHUNIT_TOTAL" & _
			"         , SUM(DECODE(NYEAR, " & v_txtYear & ", UNITKERJA, 0)) UNITKERJA_TOTAL" & _
			"         , SUM(DECODE(NYEAR, " & v_txtYear & ", KMHM, 0)) KMHM_TOTAL" & _
			"         , SUM(DECODE(NYEAR, " & v_txtYear & ", TONASE, 0)) TONASE_TOTAL" & _
			"     FROM (" & _
			"         SELECT FCBA" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"         , VEHICLECODE" & _
			"         , TO_NUMBER(TO_CHAR(VEHICLEDATE, 'YYYY')) NYEAR" & _
			"         , TO_NUMBER(TO_CHAR(VEHICLEDATE, 'MM')) NMONTH" & _
			"         , 1 JUMLAHUNIT" & _
			"         , SUM(UNITKERJA) UNITKERJA" & _
			"         , SUM(KMHM) KMHM" & _
			"         , SUM(TONASE) TONASE" & _
			"     FROM (" & _
			"         SELECT" & _
			"             FCBA" & _
			"             , VEHICLEDATE " & _
			"             , VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA) VEHICLEGROUPCODE" & _
			"             , VEHICLEGROUP_API.GET_VEHICLEGROUP_NAME(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) VEHICLEGROUPNAME" & _
			"             , VEHICLEGROUP_API.GET_ORDERREPORT(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) URUTVEHICLEGROUP" & _
			"             , 1 UNITKERJA" & _
			"             , VEHICLECODE" & _
			"             , NVL(SUM(NVL(UNIT, 0)),0) KMHM" & _
			"             , NVL(SUM(NVL(OUTPUT, 0)),0) TONASE" & _
			"         FROM" & _
			"             VEHICLEACTIVITY" & _
			"         WHERE" & _
			"             JOBCODE IN (select distinct fccode from parameterdetail where parcode = 'JOBCODE_PENGANGKUTAN')" & _
			"         GROUP BY" & _
			"             FCBA" & _
			"             , VEHICLEDATE" & _
			"             , VEHICLECODE" & _
			"         )" & _
			"     WHERE VEHICLEGROUPCODE IN ('DT', 'TK', 'TR')" & _
			"     AND TO_NUMBER(TO_CHAR(VEHICLEDATE, 'YYYY')) = " & v_txtYear & "" & _
			"     AND TO_NUMBER(TO_CHAR(VEHICLEDATE, 'MM')) <= " & v_txtMonth & "" & _
			"     AND USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E'" & _
			"     GROUP BY" & _
			"         FCBA" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"         , VEHICLECODE" & _
			"         , TO_NUMBER(TO_CHAR(VEHICLEDATE, 'YYYY'))" & _
			"         , TO_NUMBER(TO_CHAR(VEHICLEDATE, 'MM'))" & _
			"     )" & _
			"     GROUP BY" & _
			"         FCBA" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"         , VEHICLECODE"
			
strBOKPT2	=	"         , VEHICLEGROUPCODE" & _
				"         , VEHICLEGROUPNAME" & _
				"         , URUTVEHICLEGROUP" & _
				"         , SUM(JUMLAHUNIT_1) JUMLAHUNIT_1" & _
				"         , SUM(UNITKERJA_1) UNITKERJA_1" & _
				"         , SUM(KMHM_1) KMHM_1" & _
				"         , SUM(TONASE_1) TONASE_1" & _
				"         , SUM(JUMLAHUNIT_2) JUMLAHUNIT_2" & _
				"         , SUM(UNITKERJA_2) UNITKERJA_2" & _
				"         , SUM(KMHM_2) KMHM_2" & _
				"         , SUM(TONASE_2) TONASE_2" & _
				"         , SUM(JUMLAHUNIT_3) JUMLAHUNIT_3" & _
				"         , SUM(UNITKERJA_3) UNITKERJA_3" & _
				"         , SUM(KMHM_3) KMHM_3" & _
				"         , SUM(TONASE_3) TONASE_3" & _
				"         , SUM(JUMLAHUNIT_4) JUMLAHUNIT_4" & _
				"         , SUM(UNITKERJA_4) UNITKERJA_4" & _
				"         , SUM(KMHM_4) KMHM_4" & _
				"         , SUM(TONASE_4) TONASE_4" & _
				"         , SUM(JUMLAHUNIT_5) JUMLAHUNIT_5" & _
				"         , SUM(UNITKERJA_5) UNITKERJA_5" & _
				"         , SUM(KMHM_5) KMHM_5" & _
				"         , SUM(TONASE_5) TONASE_5" & _
				"         , SUM(JUMLAHUNIT_6) JUMLAHUNIT_6" & _
				"         , SUM(UNITKERJA_6) UNITKERJA_6" & _
				"         , SUM(KMHM_6) KMHM_6" & _
				"         , SUM(TONASE_6) TONASE_6" & _
				"         , SUM(JUMLAHUNIT_7) JUMLAHUNIT_7" & _
				"         , SUM(UNITKERJA_7) UNITKERJA_7" & _
				"         , SUM(KMHM_7) KMHM_7" & _
				"         , SUM(TONASE_7) TONASE_7" & _
				"         , SUM(JUMLAHUNIT_8) JUMLAHUNIT_8" & _
				"         , SUM(UNITKERJA_8) UNITKERJA_8" & _
				"         , SUM(KMHM_8) KMHM_8" & _
				"         , SUM(TONASE_8) TONASE_8" & _
				"         , SUM(JUMLAHUNIT_9) JUMLAHUNIT_9" & _
				"         , SUM(UNITKERJA_9) UNITKERJA_9" & _
				"         , SUM(KMHM_9) KMHM_9" & _
				"         , SUM(TONASE_9) TONASE_9" & _
				"         , SUM(JUMLAHUNIT_10) JUMLAHUNIT_10" & _
				"         , SUM(UNITKERJA_10) UNITKERJA_10" & _
				"         , SUM(KMHM_10) KMHM_10" & _
				"         , SUM(TONASE_10) TONASE_10" & _
				"         , SUM(JUMLAHUNIT_11) JUMLAHUNIT_11" & _
				"         , SUM(UNITKERJA_11) UNITKERJA_11" & _
				"         , SUM(KMHM_11) KMHM_11" & _
				"         , SUM(TONASE_11) TONASE_11" & _
				"         , SUM(JUMLAHUNIT_12) JUMLAHUNIT_12" & _
				"         , SUM(UNITKERJA_12) UNITKERJA_12" & _
				"         , SUM(KMHM_12) KMHM_12" & _
				"         , SUM(TONASE_12) TONASE_12" & _
				"         , SUM(JUMLAHUNIT_TOTAL) JUMLAHUNIT_TOTAL" & _
				"         , SUM(UNITKERJA_TOTAL) UNITKERJA_TOTAL" & _
				"         , SUM(KMHM_TOTAL) KMHM_TOTAL" & _
				"         , SUM(TONASE_TOTAL) TONASE_TOTAL"

			
strBOKPT	=	" select GroupPlantation " & selectCons & strBOKPT2 & "  From (" & strMain & " )" & _
			whereCons & _
			" group by GroupPlantation" & selectCons & " , VEHICLEGROUPCODE, VEHICLEGROUPNAME, URUTVEHICLEGROUP" & _
			" Order by GroupPlantation" & orderCons &" , URUTVEHICLEGROUP"
						
' response.write strBOKPT
%>
<html>

<head>
<title>ePlaS Reporting - Output Kendaraan Pengangkutan TBS</title>
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
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSumE.html"-->
				<table id="form-table" class="centered-margin">	
					<tr>
							<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM
						mon = Month(Session("v_Costbook"))
						If len(v_txtMonth)=0 then v_txtMonth=mon

						Yr = Year(Session("v_Costbook"))
						MinYr = 2015
						MaxYr = Yr  + 2
						If len(v_txtYear)= 0 then v_txtYear = Yr
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
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>
						<td><font face="verdana" size="1">GROUP KENDARAAN</font></td>
						<td>
							<select name="txtVehicleGroup" >
								<option value="--ALL--"<%If v_txtVehicleGroup = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsVehicleGroup, strVehicleGroup
								strVehicleGroup	=	" SELECT DISTINCT VEHICLEGROUPCODE, DESCRIPTION FROM VEHICLEGROUP" & _
													" WHERE VEHICLEGROUPCODE IN ('DT','TK','TR') ORDER BY VEHICLEGROUPCODE"
								
								Set rsVehicleGroup = Server.CreateObject("ADODB.Recordset")
								rsVehicleGroup.Open strVehicleGroup, connstr, adOpenStatic 
								
								if not rsVehicleGroup.eof then
								Do While Not rsVehicleGroup.eof
								%>

								<option value="<%=rsVehicleGroup("VEHICLEGROUPCODE")%>" <% If v_txtVehicleGroup = rsVehicleGroup("VEHICLEGROUPCODE") Then %> Selected <%End if%>><%=rsVehicleGroup("VEHICLEGROUPCODE")%> : <%=rsVehicleGroup("DESCRIPTION")%></option> 
								<%
									rsVehicleGroup.movenext
								loop
								end if
								rsVehicleGroup.Close
								set rsVehicleGroup = Nothing
								%>
							</select>
						</td>
						<td><input type="submit" value="View" /></td>
					</tr>				
				</table>
			</div>
			</form>
   <table border="0" width="100%" id="product-table" class="xthf-site">
        <thead>
        <tr>
        <th colspan="<%=var_colspan+54%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">    
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			</div>
			<div id="subheader" class="centered">
				<h2>OUTPUT KENDARAAN PENGANGKUTAN TBS</h2>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		</div>
        </th>
    </tr>
	<tr>		
		<%If Len(Trim(v_txtVehicleGroup)) > 0 And v_txtVehicleGroup <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP KENDARAAN</th>
		<%End If%>
		
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
			<% if v_txtfcba <>  "--ALL--" then %>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KODE KENDARAAN</th>
			<% end if %>
		<%End If%>
		
		<th colspan="4" style="background-color:#00A300;color:#fff;">JANUARI</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">FEBRUARI</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">MARET</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">APRIL</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">MEI</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">JUNI</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">JULI</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">AGUSTUS</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">SEPTEMBER</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">OKTOBER</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">NOVEMBER</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">DESEMBER</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">REALISASI</th>
	</tr>
	<tr>
		<%For LoopNo = 1 to 13%>
		<th style="background-color:#00A300;color:#fff;">JUMLAH UNIT</th>
		<th style="background-color:#00A300;color:#fff;">KM-HM</th>
		<th style="background-color:#00A300;color:#fff;">TONASE</th>
		<th style="background-color:#00A300;color:#fff;">OUTPUT</th>
		<%Next%>	
	</tr>
        </thead>
 <tbody>
<%
Set rsBOKPT = Server.CreateObject("ADODB.Recordset")
rsBOKPT.Open strBOKPT, connstr, adOpenStatic

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempVehicleGroupCode : tempVehicleGroupCode = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim SprevUraian : SprevUraian = ""

if (rsBOKPT.bof and rsBOKPT.eof) then
%>
	<tr>
		<td colspan="<%=var_colspan+53%>">NO DATA AVAILABLE</td>
	</tr>
<%else
		while not rsBOKPT.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<% if v_txtVehicleGroup <>  "--SUMMARY--" and  v_txtVehicleGroup <>  "" then %>
					<%	tempVehicleGroupCode = rsBOKPT("VEHICLEGROUPCODE")
					If Cint(visibleP) > 1 Then
						tempGroupPlantation = "--ALL--"
					Else
						If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then
							tempGroupPlantation = rsBOKPT("GroupPlantation")
						Else
							tempGroupPlantation = ""
						End If
					End If
					
					If Cint(visibleGR) > 1 And Cint(visibleP) > 1 Then
						tempPlantation = "--SUMMARY--"
					Else
						If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then
							tempPlantation = rsBOKPT("Plantation")
						Else
							tempPlantation = ""
						End If
					End If
					
					If Cint(visibleR) > 1 And Cint(visibleGR) > 1 Then
						tempGroupRegion = "--SUMMARY--"
					Else
						If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then
							tempGroupRegion = rsBOKPT("GroupRegion")
						Else
							tempGroupRegion = ""
						End If
					End If
					
					If Cint(visibleC) > 1 And Cint(visibleR) > 1 Then
						tempRegion = "--SUMMARY--"
					Else
						If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then
							tempRegion = rsBOKPT("Region")
						Else
							tempRegion = ""
						End If
					End If
					
					If Cint(visibleS) > 1 And Cint(visibleC) > 1 Then
						tempCompany = "--SUMMARY--"
					Else
						If Len(Trim(v_txtcompany)) > 0 And v_txtcompany <>  "--SUMMARY--" Then
							tempCompany = rsBOKPT("Company")
						Else
							tempCompany = ""
						End If
					End If
					
					If Cint(VisibleS) > 1 Then
						tempfcba = "--SUMMARY--"
					Else
						If Len(Trim(v_txtFcba)) > 0 And v_txtFcba <>  "--SUMMARY--" Then
							tempfcba = rsBOKPT("Fcba")
						Else
							tempfcba = ""
						End If
					End If
					%>
						<% if tempVehicleGroupCode <> SprevUraian then %>
							<td align="left">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
								<%=rsBOKPT("VEHICLEGROUPNAME")%>
								</a>
							</td>
							<% SprevUraian = tempVehicleGroupCode %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>

				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsBOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   tempGroupPlantation = rsBOKPT("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsBOKPT("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsBOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsBOKPT("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsBOKPT("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsBOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsBOKPT("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsBOKPT("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsBOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsBOKPT("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsBOKPT("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsBOKPT("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsBOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsBOKPT("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsBOKPT("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsBOKPT("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsBOKPT("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsBOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsBOKPT("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsBOKPT("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsBOKPT("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsBOKPT("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
						    <%=rsBOKPT("fcba")%>
							</a>
						</td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsBOKPT("VEHICLECODE")%></td>
						<% End if%>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsBOKPT("VEHICLECODE")%></td>
						<% End if%>
					<% end if %>
				<% end if %>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_1"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_1"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_1"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_1")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_1"))/Cdbl(rsBOKPT("UNITKERJA_1")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_2"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_2"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_2"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_2")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_2"))/Cdbl(rsBOKPT("UNITKERJA_2")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_3"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_3"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_3"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_3")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_3"))/Cdbl(rsBOKPT("UNITKERJA_3")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_4"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_4"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_4"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_4")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_4"))/Cdbl(rsBOKPT("UNITKERJA_4")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_5"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_5"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_5"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_5")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_5"))/Cdbl(rsBOKPT("UNITKERJA_5")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_6"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_6"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_6"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_6")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_6"))/Cdbl(rsBOKPT("UNITKERJA_6")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_7"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_7"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_7"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_7")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_7"))/Cdbl(rsBOKPT("UNITKERJA_7")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_8"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_8"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_8"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_8")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_8"))/Cdbl(rsBOKPT("UNITKERJA_8")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_9"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_9"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_9"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_9")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_9"))/Cdbl(rsBOKPT("UNITKERJA_9")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_10"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_10"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_10"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_10")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_10"))/Cdbl(rsBOKPT("UNITKERJA_10")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_11"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_11"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_11"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_11")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_11"))/Cdbl(rsBOKPT("UNITKERJA_11")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_12"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_12"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_12"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_12")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_12"))/Cdbl(rsBOKPT("UNITKERJA_12")), 2), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FbpNum2(rsBOKPT("JUMLAHUNIT_TOTAL"))%></td>				
				<td align="right"><%=FormatNumber(rsBOKPT("KMHM_TOTAL"), 2)%></td>
				<td align="right"><%=FormatNumber(rsBOKPT("TONASE_TOTAL"), 2)%></td>
				<%If Cdbl(rsBOKPT("UNITKERJA_TOTAL")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsBOKPT("TONASE_TOTAL"))/Cdbl(rsBOKPT("UNITKERJA_TOTAL")), 2), 2)%></td>
				<%End If%>
				
			</tr>
<%
			rsBOKPT.MoveNext
		Wend
	rsBOKPT.Close
	set rsBOKPT = Nothing
%>

<%End if%>
      <tbody>
</table>
		</div>
	</body>
</html>
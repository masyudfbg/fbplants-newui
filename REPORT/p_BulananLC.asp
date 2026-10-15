<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtJobcode, v_txtAfdeling, v_txtTahunTanam, v_txtBlok
Dim ScriptName
Dim rsBulananLC, strBulananLC, strMain, strBulananLC2
Dim NO, cClass, nday, xNum, nSpanTotal

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

nSpanTotal = 9

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
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

v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

Dim varLastDayofMonth : varLastDayofMonth = v_txtMonth & "/" & GetLastDay(v_txtMonth,v_txtYear) & "/" & v_txtYear

v_txtJobcode	=	Request("txtJobcode")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtTahunTanam	=	Request("txtTahunTanam")
v_txtBlok		=	Request("txtBlok")

If Len(Trim(v_txtJobcode)) = 0 Then
	v_txtJobcode = "--ALL--"
End If

If Len(Trim(v_txtJobcode)) > 0 Then
	if v_txtJobcode = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtJobcode = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And JOBCODE = '" & v_txtJobcode & "'"
		orderCons	=	orderCons	& " "
	end if
End If

If Len(Trim(v_txtAfdeling)) > 0 Then
	If v_txtAfdeling = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtAfdeling = "--ALL--" then
		selectCons	=	selectCons	& " , AFDELING"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , AFDELING"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , AFDELING"
		whereCons	=	wherecons	& " And AFDELING = '" & v_txtAfdeling & "'"
		orderCons	=	orderCons	& " , AFDELING"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtTahunTanam)) > 0 Then
	If v_txtTahunTanam = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtTahunTanam = "--ALL--" then
		selectCons	=	selectCons	& " , TAHUNTANAM"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , TAHUNTANAM"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , TAHUNTANAM"
		whereCons	=	wherecons	& " And TAHUNTANAM = '" & v_txtTahunTanam & "'"
		orderCons	=	orderCons	& " , TAHUNTANAM"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtBlok)) > 0 Then
	If v_txtBlok = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtBlok = "--ALL--" then
		selectCons	=	selectCons	& " , Blok"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , Blok"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Blok"
		whereCons	=	wherecons	& " And Blok = '" & v_txtBlok & "'"
		orderCons	=	orderCons	& " , Blok"
		var_colspan	=	var_colspan	+ 1
	End if
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _
				"          , JOBCODE," & _
				"          JOBNAME," & _
				"          AFDELING," & _
				"          TAHUNTANAM," & _
				"          BLOK," & _
				"          SUM (DECODE (BULAN, 1, HA_REALISASI_BI, 0)) HA_MONTH_1," & _
				"          SUM (DECODE (BULAN, 2, HA_REALISASI_BI, 0)) HA_MONTH_2," & _
				"          SUM (DECODE (BULAN, 3, HA_REALISASI_BI, 0)) HA_MONTH_3," & _
				"          SUM (DECODE (BULAN, 4, HA_REALISASI_BI, 0)) HA_MONTH_4," & _
				"          SUM (DECODE (BULAN, 5, HA_REALISASI_BI, 0)) HA_MONTH_5," & _
				"          SUM (DECODE (BULAN, 6, HA_REALISASI_BI, 0)) HA_MONTH_6," & _
				"          SUM (DECODE (BULAN, 7, HA_REALISASI_BI, 0)) HA_MONTH_7," & _
				"          SUM (DECODE (BULAN, 8, HA_REALISASI_BI, 0)) HA_MONTH_8," & _
				"          SUM (DECODE (BULAN, 9, HA_REALISASI_BI, 0)) HA_MONTH_9," & _
				"          SUM (DECODE (BULAN, 10, HA_REALISASI_BI, 0)) HA_MONTH_10," & _
				"          SUM (DECODE (BULAN, 11, HA_REALISASI_BI, 0)) HA_MONTH_11," & _
				"          SUM (DECODE (BULAN, 12, HA_REALISASI_BI, 0)) HA_MONTH_12," & _
				"          SUM (DECODE (BULAN, 1, RP_REALISASI_BI, 0)) RP_MONTH_1," & _
				"          SUM (DECODE (BULAN, 2, RP_REALISASI_BI, 0)) RP_MONTH_2," & _
				"          SUM (DECODE (BULAN, 3, RP_REALISASI_BI, 0)) RP_MONTH_3," & _
				"          SUM (DECODE (BULAN, 4, RP_REALISASI_BI, 0)) RP_MONTH_4," & _
				"          SUM (DECODE (BULAN, 5, RP_REALISASI_BI, 0)) RP_MONTH_5," & _
				"          SUM (DECODE (BULAN, 6, RP_REALISASI_BI, 0)) RP_MONTH_6," & _
				"          SUM (DECODE (BULAN, 7, RP_REALISASI_BI, 0)) RP_MONTH_7," & _
				"          SUM (DECODE (BULAN, 8, RP_REALISASI_BI, 0)) RP_MONTH_8," & _
				"          SUM (DECODE (BULAN, 9, RP_REALISASI_BI, 0)) RP_MONTH_9," & _
				"          SUM (DECODE (BULAN, 10, RP_REALISASI_BI, 0)) RP_MONTH_10," & _
				"          SUM (DECODE (BULAN, 11, RP_REALISASI_BI, 0)) RP_MONTH_11," & _
				"          SUM (DECODE (BULAN, 12, RP_REALISASI_BI, 0)) RP_MONTH_12," & _
				"          SUM (" & _
				"             DECODE (TAHUN," & _
				"                     " & v_txtYear & ", HA_REALISASI_BI, 0))" & _
				"             HA_REAL," & _
				"          SUM (" & _
				"             DECODE (TAHUN," & _
				"                     " & v_txtYear & ", RP_REALISASI_BI, 0))" & _
				"             RP_REAL," & _
				"          SUM (HA_BUDGET_BI)" & _
				"             HA_BUDGET," & _
				"          SUM (RP_BUDGET_BI)" & _
				"             RP_BUDGET" & _
				"     FROM (  SELECT FCBA," & _
				"                    TAHUN," & _
				"                    BULAN," & _
				"                    JOBCODE," & _
				"                    JOBNAME," & _
				"                    AFDELING," & _
				"                    TAHUNTANAM," & _
				"                    BLOK," & _
				"                    HA_REALISASI_BI," & _
				"                    HA_BUDGET_BI," & _
				"                    RP_REALISASI_BI," & _
				"                    RP_BUDGET_BI" & _
				"               FROM REPORT_LANDCLEARING_MONTH_TAB" & _
				"              WHERE TAHUN = " & v_txtYear & "" & _
				"                    AND BULAN <= " & v_txtMonth & "" & _
				"           )" & _
				" GROUP BY FCBA," & _
				"          JOBCODE," & _
				"          JOBNAME," & _
				"          AFDELING," & _
				"          TAHUNTANAM," & _
				"          BLOK"
								
strBulananLC2	=	"     , JOBCODE" & _
					"     , JOBNAME" & _
					"     , SUM(HA_MONTH_1) HA_MONTH_1" & _
					"     , SUM(HA_MONTH_2) HA_MONTH_2" & _
					"     , SUM(HA_MONTH_3) HA_MONTH_3" & _
					"     , SUM(HA_MONTH_4) HA_MONTH_4" & _
					"     , SUM(HA_MONTH_5) HA_MONTH_5" & _
					"     , SUM(HA_MONTH_6) HA_MONTH_6" & _
					"     , SUM(HA_MONTH_7) HA_MONTH_7" & _
					"     , SUM(HA_MONTH_8) HA_MONTH_8" & _
					"     , SUM(HA_MONTH_9) HA_MONTH_9" & _
					"     , SUM(HA_MONTH_10) HA_MONTH_10" & _
					"     , SUM(HA_MONTH_11) HA_MONTH_11" & _
					"     , SUM(HA_MONTH_12) HA_MONTH_12" & _
					"     , SUM(RP_MONTH_1) RP_MONTH_1" & _
					"     , SUM(RP_MONTH_2) RP_MONTH_2" & _
					"     , SUM(RP_MONTH_3) RP_MONTH_3" & _
					"     , SUM(RP_MONTH_4) RP_MONTH_4" & _
					"     , SUM(RP_MONTH_5) RP_MONTH_5" & _
					"     , SUM(RP_MONTH_6) RP_MONTH_6" & _
					"     , SUM(RP_MONTH_7) RP_MONTH_7" & _
					"     , SUM(RP_MONTH_8) RP_MONTH_8" & _
					"     , SUM(RP_MONTH_9) RP_MONTH_9" & _
					"     , SUM(RP_MONTH_10) RP_MONTH_10" & _
					"     , SUM(RP_MONTH_11) RP_MONTH_11" & _
					"     , SUM(RP_MONTH_12) RP_MONTH_12" & _
					"     , SUM(HA_REAL) HA_REAL" & _
					"     , SUM(HA_BUDGET) HA_BUDGET" & _
					"     , SUM(RP_REAL) RP_REAL" & _
					"     , SUM(RP_BUDGET) RP_BUDGET"

strBulananLC	=	" select GroupPlantation " & selectCons & strBulananLC2 &"  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation" & selectCons & " , JOBCODE, JOBNAME" & _
					" having ((SUM(HA_REAL) + SUM(HA_BUDGET) > 0) Or (SUM(RP_REAL) + SUM(RP_BUDGET) > 0))" & _
					" Order by GroupPlantation" & orderCons & ", JOBCODE"

'Response.write strBulananLC
%>
<html>

<head>
<title>ePlaS Reporting - Bulanan Pembukaan Lahan</title>
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
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
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
							<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td>						
						<td>
							<font face="verdana" size="1">JOB CODE</font>
						</td>
						<td>
							<select name="txtJobcode">
								<option value="--ALL--"<%If v_txtJobcode = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsJobCode, strJobCode
								strJobCode	=	" SELECT DISTINCT FCCODE, FCNAME FROM JOB WHERE FCCODE IN (" & _
												" '1320103', '1320104', '1320105', '1320106'," & _
												" '1320401', '1320501', '1320502', '1320508') ORDER BY FCCODE"
								Set rsJobCode = Server.CreateObject("ADODB.Recordset")
								rsJobCode.Open strJobCode, connstr, adOpenStatic 
								
								if not rsJobCode.eof then
								Do While Not rsJobCode.eof
								if rsJobCode("FCCODE") = v_txtJobcode then%>
								<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> : <%=rsJobCode("FCNAME")%></option> 
								<%Else%>
								<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> : <%=rsJobCode("FCNAME")%></option>  
								<%End if
									rsJobCode.movenext
								loop
								end if
								rsJobCode.Close
								set rsJobCode = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">AFDELING</font>
						</td>
						<td>
							<select name="txtAfdeling" >
								<option value="--SUMMARY--"<%If v_txtAfdeling = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtAfdeling = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsAfdeling, strAfdeling
								strAfdeling	=	" SELECT DISTINCT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' ORDER BY FCCODE"
								
								Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
								rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
								
								if not rsAfdeling.eof then
								Do While Not rsAfdeling.eof
								if rsAfdeling("FCCODE") = v_txtAfdeling then%>
								<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
								<%Else%>
								<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
								<%End if
									rsAfdeling.movenext
								loop
								end if
								rsAfdeling.Close
								set rsAfdeling = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">TT</font>
						</td>
						<td>
							<select name="txtTahunTanam" >
								<option value="--SUMMARY--"<%If v_txtTahunTanam = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtTahunTanam = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsTahunTanam, strTahunTanam
								strTahunTanam	=	" SELECT DISTINCT TO_CHAR(PLANTINGDATE, 'YYYY') TT FROM FIELD WHERE ACTIVATION = 'Y'" & _
													" ORDER BY TO_CHAR(PLANTINGDATE, 'YYYY')"
								
								Set rsTahunTanam = Server.CreateObject("ADODB.Recordset")
								rsTahunTanam.Open strTahunTanam, connstr, adOpenStatic 
								
								if not rsTahunTanam.eof then
								Do While Not rsTahunTanam.eof
								if rsTahunTanam("TT") = v_txtTahunTanam then%>
								<option value="<%=rsTahunTanam("TT")%>" Selected><%=rsTahunTanam("TT")%></option> 
								<%Else%>
								<option value="<%=rsTahunTanam("TT")%>"><%=rsTahunTanam("TT")%></option>  
								<%End if
									rsTahunTanam.movenext
								loop
								end if
								rsTahunTanam.Close
								set rsTahunTanam = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">BLOK</font>
						</td>
						<td>
							<select name="txtBlok" >
								<option value="--SUMMARY--"<%If v_txtBlok = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtBlok = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsBlok, strBlok
								strBlok	=	" SELECT DISTINCT FCCODE FROM FIELD WHERE ACTIVATION = 'Y' ORDER BY FCCODE"
								
								Set rsBlok = Server.CreateObject("ADODB.Recordset")
								rsBlok.Open strBlok, connstr, adOpenStatic 
								
								if not rsBlok.eof then
								Do While Not rsBlok.eof
								if rsBlok("FCCODE") = v_txtBlok then%>
								<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%></option> 
								<%Else%>
								<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%></option>  
								<%End if
									rsBlok.movenext
								loop
								end if
								rsBlok.Close
								set rsBlok = Nothing
								%>
							</select>
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>

<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
		<tr>
			<th colspan="<%=var_colspan+36%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header-wrapper">
		<div id="head_judul">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>LAPORAN BULANAN PEMBUKAAN LAHAN</h3>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
              </div
			</th>
		</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">NAMA PEKEJAAN</th>
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
		<%End If%>
		
		<%If Len(Trim(v_txtAfdeling)) > 0 And v_txtAfdeling <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<%End if %>
		
		<%If Len(Trim(v_txtTahunTanam)) > 0 And v_txtTahunTanam <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">TAHUN TANAM</th>
		<%End if %>
		
		<%If Len(Trim(v_txtBlok)) > 0 And v_txtBlok <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<%End if %>
		
		<th colspan="3" style="background-color:#00A300;color:#fff;">JANUARI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">FEBRUARI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">MARET</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">APRIL</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">MEI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JUNI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JULI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">AGUSTUS</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SEPTEMBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">OKTOBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">NOVEMBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">DESEMBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">VARIAN</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
		<th style="background-color:#00A300;color:#fff;">HA</th>
		<th style="background-color:#00A300;color:#fff;">RP</th>
		<th style="background-color:#00A300;color:#fff;">RP/ HA</th>
	</tr>
    </thead>
	<tbody>
<%
Set rsBulananLC = Server.CreateObject("ADODB.Recordset")
'response.write strBulananLC
rsBulananLC.Open strBulananLC, connstr, adOpenStatic

Dim Rp_Per_Ha_RowReal : Rp_Per_Ha_RowReal = 0
Dim Rp_Per_Ha_RowBudget : Rp_Per_Ha_RowBudget = 0
Dim RowPercentage : RowPercentage = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempJobCode : tempJobCode = ""
Dim tempAfdeling : tempAfdeling = ""
Dim tempTahunTanam : tempTahunTanam = ""
Dim tempBlok : tempBlok = ""


Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sprevJobCode : sprevJobCode = ""
Dim sprevAfdeling : sprevAfdeling = ""
Dim sprevTahunTanam : sprevTahunTanam = ""
Dim sprevBlok : sprevBlok = ""

if rsBulananLC.bof and rsBulananLC.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+36%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsBulananLC.eof
%>
				<tr>
				<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <>  "" then %>
					<%
					   tempJobCode = rsBulananLC("JOBCODE")
					   tempGroupPlantation = "--ALL--"
					   tempPlantation = "--SUMMARY--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempJobCode <> SprevJobCode then %>
						<td align="left">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsBulananLC("JOBNAME")%>
							</a>
						</td>
						<% SprevJobCode = tempJobCode %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   tempGroupPlantation = rsBulananLC("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsBulananLC("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBulananLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsBulananLC("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsBulananLC("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBulananLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBulananLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsBulananLC("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsBulananLC("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBulananLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBulananLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsBulananLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsBulananLC("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsBulananLC("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBulananLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBulananLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsBulananLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsBulananLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsBulananLC("Company")
					   tempfcba = "--ALL--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsBulananLC("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBulananLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBulananLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsBulananLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsBulananLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsBulananLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsBulananLC("fcba")
					   tempAfdeling = "--ALL--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsBulananLC("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtAfdeling <>  "--SUMMARY--" and v_txtAfdeling <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBulananLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBulananLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsBulananLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsBulananLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsBulananLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsBulananLC("fcba")
					   else
					       tempfcba =""
					   end if

					   tempAfdeling = rsBulananLC("AFDELING")
					   tempTahunTanam = "--ALL--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempAfdeling <> sprevAfdeling then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsBulananLC("AFDELING")%>
							</a>
						</td>
						<% sprevAfdeling = tempAfdeling %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>

				<% if v_txtTahunTanam <>  "--SUMMARY--" and v_txtTahunTanam <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsBulananLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsBulananLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsBulananLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsBulananLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsBulananLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsBulananLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsBulananLC("fcba")
					   else
					       tempfcba =""
					   end if
					   
					   if  v_txtAfdeling  <>  "--SUMMARY--" and v_txtAfdeling <> "" then
					       tempAfdeling = rsBulananLC("AFDELING")
					   else
					       tempAfdeling =""
					   end if

					   tempTahunTanam = Cstr(rsBulananLC("TAHUNTANAM"))
					   tempBlok = "--ALL--"
					 %>
					
					<% if tempTahunTanam <> sprevTahunTanam then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsBulananLC("TAHUNTANAM")%>
							</a>
						</td>
						<% sprevTahunTanam = tempTahunTanam %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <>  "" then %>
				    <% if sprevBlok <> rsBulananLC("Blok") then %>				
					    <td align="center"><%=rsBulananLC("Blok")%></td>				
						<% sprevBlok = rsBulananLC("Blok")%>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_1"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_1"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_1")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_1"))/Cdbl(rsBulananLC("HA_MONTH_1")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_2"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_2"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_2")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_2"))/Cdbl(rsBulananLC("HA_MONTH_2")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_3"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_3"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_3")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_3"))/Cdbl(rsBulananLC("HA_MONTH_3")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_4"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_4"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_4")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_4"))/Cdbl(rsBulananLC("HA_MONTH_4")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_5"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_5"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_5")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_5"))/Cdbl(rsBulananLC("HA_MONTH_5")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_6"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_6"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_6")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_6"))/Cdbl(rsBulananLC("HA_MONTH_6")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_7"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_7"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_7")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_7"))/Cdbl(rsBulananLC("HA_MONTH_7")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_8"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_8"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_8")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_8"))/Cdbl(rsBulananLC("HA_MONTH_8")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_9"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_9"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_9")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_9"))/Cdbl(rsBulananLC("HA_MONTH_9")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_10"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_10"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_10")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_10"))/Cdbl(rsBulananLC("HA_MONTH_10")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_11"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_11"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_11")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_11"))/Cdbl(rsBulananLC("HA_MONTH_11")), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_MONTH_12"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_MONTH_12"), 2)%></td>
					<%If Cdbl(rsBulananLC("HA_MONTH_12")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsBulananLC("RP_MONTH_12"))/Cdbl(rsBulananLC("HA_MONTH_12")), 2)%></td>
					<%End If%>

					<%
					If Cdbl(rsBulananLC("HA_REAL")) = 0 Then
						Rp_Per_Ha_RowReal = 0
					Else
						Rp_Per_Ha_RowReal = (Cdbl(rsBulananLC("RP_REAL"))/Cdbl(rsBulananLC("HA_REAL")))
					End If
					%>
					<td align="right"><%=FormatNumber(rsBulananLC("HA_REAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_REAL"), 2)%></td>
					<td align="right"><%=FormatNumber(Rp_Per_Ha_RowReal, 2)%></td>

					<%
					If Cdbl(rsBulananLC("HA_BUDGET")) = 0 Then
						Rp_Per_Ha_RowBudget = 0
					Else
						Rp_Per_Ha_RowBudget = (Cdbl(rsBulananLC("RP_BUDGET"))/Cdbl(rsBulananLC("HA_BUDGET")))
					End If
					%>					
					<td align="right"><%=FormatNumber(rsBulananLC("HA_BUDGET"), 2)%></td>
					<td align="right"><%=FormatNumber(rsBulananLC("RP_BUDGET"), 2)%></td>
					<td align="right"><%=FormatNumber(Rp_Per_Ha_RowBudget, 2)%></td>
					
					<%If Cdbl(rsBulananLC("HA_BUDGET")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber((Cdbl(rsBulananLC("HA_BUDGET")) - Cdbl(rsBulananLC("HA_REAL")))/Cdbl(rsBulananLC("HA_BUDGET"))*100, 2)%></td>
					<%End If%>
					<%If Cdbl(rsBulananLC("RP_BUDGET")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber((Cdbl(rsBulananLC("RP_BUDGET")) - Cdbl(rsBulananLC("RP_REAL")))/Cdbl(rsBulananLC("RP_BUDGET"))*100, 2)%></td>
					<%End If%>
					<%If Rp_Per_Ha_RowBudget = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber((Rp_Per_Ha_RowBudget - Rp_Per_Ha_RowReal)/Rp_Per_Ha_RowBudget*100, 2)%></td>
					<%End If%>
				</tr>
<%
			rsBulananLC.MoveNext
			Rp_Per_Ha_RowReal = 0
			Rp_Per_Ha_RowBudget = 0
			Wend	
rsBulananLC.Close
set rsBulananLC = Nothing
%>
<%End if%>
	</tbody>
</table>

   </body>
   </html>
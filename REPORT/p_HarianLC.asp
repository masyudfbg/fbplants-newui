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
Dim rsHarianLC, strHarianLC, strMain, strHarianLC2
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
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 1, REAL_HI, 0)) DAY_1," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 2, REAL_HI, 0)) DAY_2," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 3, REAL_HI, 0)) DAY_3," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 4, REAL_HI, 0)) DAY_4," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 5, REAL_HI, 0)) DAY_5," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 6, REAL_HI, 0)) DAY_6," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 7, REAL_HI, 0)) DAY_7," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 8, REAL_HI, 0)) DAY_8," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 9, REAL_HI, 0)) DAY_9," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 10, REAL_HI, 0)) DAY_10," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 11, REAL_HI, 0)) DAY_11," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 12, REAL_HI, 0)) DAY_12," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 13, REAL_HI, 0)) DAY_13," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 14, REAL_HI, 0)) DAY_14," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 15, REAL_HI, 0)) DAY_15," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 16, REAL_HI, 0)) DAY_16," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 17, REAL_HI, 0)) DAY_17," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 18, REAL_HI, 0)) DAY_18," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 19, REAL_HI, 0)) DAY_19," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 20, REAL_HI, 0)) DAY_20," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 21, REAL_HI, 0)) DAY_21," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 22, REAL_HI, 0)) DAY_22," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 23, REAL_HI, 0)) DAY_23," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 24, REAL_HI, 0)) DAY_24," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 25, REAL_HI, 0)) DAY_25," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 26, REAL_HI, 0)) DAY_26," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 27, REAL_HI, 0)) DAY_27," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 28, REAL_HI, 0)) DAY_28," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 29, REAL_HI, 0)) DAY_29," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 30, REAL_HI, 0)) DAY_30," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 31, REAL_HI, 0)) DAY_31," & _
				"          SUM (" & _
				"             DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY'))," & _
				"                     " & v_txtYear & ", REAL_HI," & _
				"                     TO_NUMBER (TO_CHAR (TANGGAL, 'MM')), " & v_txtMonth & "," & _
				"                     REAL_HI, 0))" & _
				"             REAL," & _
				"          SUM (" & _
				"             DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY'))," & _
				"                     " & v_txtYear & ", BUDGET_HI," & _
				"                     TO_NUMBER (TO_CHAR (TANGGAL, 'MM')), " & v_txtMonth & "," & _
				"                     BUDGET_HI, 0))" & _
				"             BUDGET" & _
				"     FROM (  SELECT FCBA," & _
				"                    TANGGAL," & _
				"                    JOBCODE," & _
				"                    JOBNAME," & _
				"                    AFDELING," & _
				"                    TAHUNTANAM," & _
				"                    BLOK," & _
				"                    REAL_HI," & _
				"                    BUDGET_HI" & _
				"               FROM REPORT_LANDCLEARING_DAILY_TAB" & _
				"              WHERE TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY')) = " & v_txtYear & "" & _
				"                    AND TO_NUMBER (TO_CHAR (TANGGAL, 'MM')) = " & v_txtMonth & "" & _
				"                    AND REAL_HI + BUDGET_HI > 0" & _
				"          )" & _
				"           GROUP BY FCBA," & _
				"                    TANGGAL," & _
				"                    JOBCODE," & _
				"                    JOBNAME," & _
				"                    AFDELING," & _
				"                    TAHUNTANAM," & _
				"                    BLOK"
								
strHarianLC2	=	"     , JOBCODE" & _
					"     , JOBNAME" & _
					"     , SUM(DAY_1) DAY_1" & _
					"     , SUM(DAY_2) DAY_2" & _
					"     , SUM(DAY_3) DAY_3" & _
					"     , SUM(DAY_4) DAY_4" & _
					"     , SUM(DAY_5) DAY_5" & _
					"     , SUM(DAY_6) DAY_6" & _
					"     , SUM(DAY_7) DAY_7" & _
					"     , SUM(DAY_8) DAY_8" & _
					"     , SUM(DAY_9) DAY_9" & _
					"     , SUM(DAY_10) DAY_10" & _
					"     , SUM(DAY_11) DAY_11" & _
					"     , SUM(DAY_12) DAY_12" & _
					"     , SUM(DAY_13) DAY_13" & _
					"     , SUM(DAY_14) DAY_14" & _
					"     , SUM(DAY_15) DAY_15" & _
					"     , SUM(DAY_16) DAY_16" & _
					"     , SUM(DAY_17) DAY_17" & _
					"     , SUM(DAY_18) DAY_18" & _
					"     , SUM(DAY_19) DAY_19" & _
					"     , SUM(DAY_20) DAY_20" & _
					"     , SUM(DAY_21) DAY_21" & _
					"     , SUM(DAY_22) DAY_22" & _
					"     , SUM(DAY_23) DAY_23" & _
					"     , SUM(DAY_24) DAY_24" & _
					"     , SUM(DAY_25) DAY_25" & _
					"     , SUM(DAY_26) DAY_26" & _
					"     , SUM(DAY_27) DAY_27" & _
					"     , SUM(DAY_28) DAY_28" & _
					"     , SUM(DAY_29) DAY_29" & _
					"     , SUM(DAY_30) DAY_30" & _
					"     , SUM(DAY_31) DAY_31" & _
					"     , SUM(REAL) REAL" & _
					"     , SUM(BUDGET) BUDGET"

strHarianLC	=	" select GroupPlantation " & selectCons & strHarianLC2 &"  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation" & selectCons & " , JOBCODE, JOBNAME" & _
					" having (SUM(REAL) + SUM(BUDGET) > 0)" & _
					" Order by GroupPlantation" & orderCons & ", JOBCODE"

'response.write strHarianLC
%>
<html>

<head>
<title>ePlaS Reporting - Harian Pembukaan Lahan</title>
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
												" select distinct fccode from parameter where parcode = 'RPT_HARIAN_LC') ORDER BY FCCODE"
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
								strTahunTanam	=	" SELECT DISTINCT TO_CHAR(PLANTINGDATE, 'YYYY') TT FROM Budget_Field WHERE ACTIVATION = 'Y'" & _
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
								strBlok	=	" SELECT DISTINCT FCCODE FROM fielddetail WHERE ACTIVATION = 'Y' ORDER BY FCCODE"
								
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
				<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>LAPORAN HARIAN PEMBUKAAN LAHAN</h3>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
			</th>
		</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">NAMA PEKEJAAN</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">UOM</th>
		
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
		
		<th colspan="31" style="background-color:#00A300;color:#fff;">TANGGAL</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SHI</th>
	</tr>
	<tr>
		<%for nday = 1 to 31%>
		<th style="background-color:#00A300;color:#fff;"><%=nday%></th>
		<%next%>
		<th style="background-color:#00A300;color:#fff;">REAL</th>
		<th style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
    </thead>
	<tbody>
<%
Set rsHarianLC = Server.CreateObject("ADODB.Recordset")
rsHarianLC.Open strHarianLC, connstr, adOpenStatic

Dim RowReal : RowReal = 0
Dim RowBudget : RowBudget = 0
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

if rsHarianLC.bof and rsHarianLC.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+36%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsHarianLC.eof
%>
				<tr>
				<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <>  "" then %>
					<%
					   tempJobCode = rsHarianLC("JOBCODE")
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
							<%=rsHarianLC("JOBNAME")%>
							</a>
						</td>
						<% SprevJobCode = tempJobCode %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
					<td align="center">HA</td>
				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   tempGroupPlantation = rsHarianLC("GroupPlantation")
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
							<%=rsHarianLC("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsHarianLC("Plantation")
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
							<%=rsHarianLC("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsHarianLC("GroupRegion")
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
							<%=rsHarianLC("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsHarianLC("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsHarianLC("Company")
					   tempfcba = "--ALL--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsHarianLC("fcba")
					   tempAfdeling = "--ALL--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsHarianLC("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtAfdeling <>  "--SUMMARY--" and v_txtAfdeling <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsHarianLC("fcba")
					   else
					       tempfcba =""
					   end if

					   tempAfdeling = rsHarianLC("AFDELING")
					   tempTahunTanam = "--ALL--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempAfdeling <> sprevAfdeling then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsHarianLC("AFDELING")%>
							</a>
						</td>
						<% sprevAfdeling = tempAfdeling %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>

				<% if v_txtTahunTanam <>  "--SUMMARY--" and v_txtTahunTanam <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsHarianLC("fcba")
					   else
					       tempfcba =""
					   end if
					   
					   if  v_txtAfdeling  <>  "--SUMMARY--" and v_txtAfdeling <> "" then
					       tempAfdeling = rsHarianLC("AFDELING")
					   else
					       tempAfdeling =""
					   end if

				
					        tempTahunTanam = rsHarianLC("TAHUNTANAM")
					 
					   
					
					   tempBlok = "--ALL--"
					   
					   
					 %>
					
					<% 
                      if   sprevTahunTanam ="" then  
                           sprevTahunTanam ="1976"   					  
					  end if
					  if cint(tempTahunTanam) <> cint(sprevTahunTanam) then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsHarianLC("TAHUNTANAM")%>
							</a>
						</td>
						<% sprevTahunTanam = tempTahunTanam %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <>  "" then %>
				    <% if sprevBlok <> rsHarianLC("Blok") then %>				
					    <td align="center"><%=rsHarianLC("Blok")%></td>				
						<% sprevBlok = rsHarianLC("Blok")%>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_1"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_2"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_3"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_4"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_5"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_6"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_7"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_8"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_9"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_10"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_11"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_12"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_13"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_14"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_15"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_16"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_17"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_18"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_19"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_20"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_21"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_22"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_23"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_24"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_25"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_26"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_27"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_28"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_29"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_30"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_31"), 2)%></td>
<%
			RowReal	=	RowReal + Cdbl(rsHarianLC("REAL"))
			RowBudget	=	RowBudget + Cdbl(rsHarianLC("BUDGET"))
			If RowBudget = 0 Then
				RowPercentage = 0
			Else
				RowPercentage = RowReal/RowBudget*100
			End If
%>
					<td align="right"><%=FormatNumber(RowReal, 2)%></td>
					<td align="right"><%=FormatNumber(RowBudget, 2)%></td>
					<td align="right"><%=FormatNumber(RowPercentage, 2)%></td>
				</tr>
<%
			rsHarianLC.MoveNext
			
			RowReal = 0
			RowBudget = 0
			RowPercentage = 0
			Wend	
rsHarianLC.Close
set rsHarianLC = Nothing
%>
<%End if%>
	</tbody>
</table>

   </body>
   </html>
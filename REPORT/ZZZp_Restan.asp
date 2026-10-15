<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtAfdeling, v_txtBlok
Dim ScriptName
Dim rsRestan, strRestan, strMain, strRestan2
Dim NO, cClass, nday, xNum, nSpanTotal

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

nSpanTotal = 7

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
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	+ 1
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

v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")

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

If Len(Trim(v_txtBlok)) > 0 Then
	If v_txtBlok = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtBlok = "--ALL--" then
		selectCons	=	selectCons	& " , FIELDCODE"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , FIELDCODE"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , FIELDCODE"
		whereCons	=	wherecons	& " And FIELDCODE = '" & v_txtBlok & "'"
		orderCons	=	orderCons	& " , FIELDCODE"
		var_colspan	=	var_colspan	+ 1
	End if
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _
				"     , AFDELING" & _
				"     , FIELDCODE" & _
				"     , FIELDNAME" & _
				"     , TAHUN" & _
				"     , BULAN" & _
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
				"     FROM (" & _
				"         SELECT " & _
				"             HR.FCBA" & _
				"             , FF.DIVISION AFDELING" & _
				"             , HR.FIELDCODE" & _
				"             , FF.FCNAME FIELDNAME" & _
				"             , TO_NUMBER(TO_CHAR(ONDATE,'YYYY')) TAHUN" & _
				"             , TO_NUMBER(TO_CHAR(ONDATE,'MM')) BULAN" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 1, ROUND(RESTANQTY, 0), 0)) DAY_1" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 2, ROUND(RESTANQTY, 0), 0)) DAY_2" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 3, ROUND(RESTANQTY, 0), 0)) DAY_3" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 4, ROUND(RESTANQTY, 0), 0)) DAY_4" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 5, ROUND(RESTANQTY, 0), 0)) DAY_5" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 6, ROUND(RESTANQTY, 0), 0)) DAY_6" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 7, ROUND(RESTANQTY, 0), 0)) DAY_7" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 8, ROUND(RESTANQTY, 0), 0)) DAY_8" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 9, ROUND(RESTANQTY, 0), 0)) DAY_9" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 10, ROUND(RESTANQTY, 0), 0)) DAY_10" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 11, ROUND(RESTANQTY, 0), 0)) DAY_11" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 12, ROUND(RESTANQTY, 0), 0)) DAY_12" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 13, ROUND(RESTANQTY, 0), 0)) DAY_13" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 14, ROUND(RESTANQTY, 0), 0)) DAY_14" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 15, ROUND(RESTANQTY, 0), 0)) DAY_15" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 16, ROUND(RESTANQTY, 0), 0)) DAY_16" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 17, ROUND(RESTANQTY, 0), 0)) DAY_17" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 18, ROUND(RESTANQTY, 0), 0)) DAY_18" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 19, ROUND(RESTANQTY, 0), 0)) DAY_19" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 20, ROUND(RESTANQTY, 0), 0)) DAY_20" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 21, ROUND(RESTANQTY, 0), 0)) DAY_21" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 22, ROUND(RESTANQTY, 0), 0)) DAY_22" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 23, ROUND(RESTANQTY, 0), 0)) DAY_23" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 24, ROUND(RESTANQTY, 0), 0)) DAY_24" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 25, ROUND(RESTANQTY, 0), 0)) DAY_25" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 26, ROUND(RESTANQTY, 0), 0)) DAY_26" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 27, ROUND(RESTANQTY, 0), 0)) DAY_27" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 28, ROUND(RESTANQTY, 0), 0)) DAY_28" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 29, ROUND(RESTANQTY, 0), 0)) DAY_29" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 30, ROUND(RESTANQTY, 0), 0)) DAY_30" & _
				"             , SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE,'DD')), 31, ROUND(RESTANQTY, 0), 0)) DAY_31" & _
				"         FROM " & _
				"             HARVESTING_RESTAN_TAB HR" & _
				"         INNER JOIN FIELD FF ON (HR.FCBA = FF.FCBA AND HR.FIELDCODE = FF.FCCODE)" & _
				"         GROUP BY " & _
				"             HR.FCBA" & _
				"             , FF.DIVISION" & _
				"             , HR.FIELDCODE" & _
				"             , FF.FCNAME" & _
				"             , TO_NUMBER(TO_CHAR(ONDATE,'YYYY'))" & _
				"             , TO_NUMBER(TO_CHAR(ONDATE,'MM'))" & _
				"     ) WHERE TAHUN = " & v_txtYear & " AND BULAN = " & v_txtMonth

strMain	=	strMain	&	" GROUP BY" & _
								"     BUSINESSUNIT_API.get_OrderData(fcba)" & _
								"     , FCBA" & _
								"     , AFDELING" & _
								"     , FIELDCODE" & _
								"     , FIELDNAME" & _
								"     , TAHUN" & _
								"     , BULAN"

strMain	=	strMain	&	" ORDER BY" & _
								"   BUSINESSUNIT_API.get_OrderData(fcba)" & _
								" 	, TAHUN" & _
								" 	, BULAN" & _
								" 	, AFDELING" & _
								" 	, FIELDCODE"
								
								
strRestan2	=	"     , SUM(DAY_1) DAY_1" & _
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
				"     , SUM(DAY_31) DAY_31"

strRestan	=	" select GroupPlantation " & selectCons & strRestan2 &"  From (" & strMain & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by GroupPlantation" & selectCons & "  Order by GroupPlantation" & orderCons

'response.write strRestan
%>
<html>

<head>
<title>ePlaS Reporting - Restan</title>
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
							MinYr = Yr - 2
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
							<font face="verdana" size="1">AFDELING</font>
						</td>
						<td>
							<select name="txtAfdeling" >
								<option value="--SUMMARY--"<%If v_txtAfdeling = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtAfdeling = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsAfdeling, strAfdeling
								strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE"
								
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
							<font face="verdana" size="1">BLOK</font>
						</td>
						<td>
							<select name="txtBlok" >
								<option value="--SUMMARY--"<%If v_txtBlok = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtBlok = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsBlok, strBlok
								strBlok	=	" SELECT FCCODE, FCNAME FROM FIELD WHERE FCBA = '" & Session("FCBA") & "'" & _
											" AND STATUS = 'TM' ORDER BY FCCODE"
								
								Set rsBlok = Server.CreateObject("ADODB.Recordset")
								rsBlok.Open strBlok, connstr, adOpenStatic 
								
								if not rsBlok.eof then
								Do While Not rsBlok.eof
								if rsBlok("FCCODE") = v_txtBlok then%>
								<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option> 
								<%Else%>
								<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option>  
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
			<th colspan="<%=var_colspan+31%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>LAPORAN RESTAN</h3>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
			</th>
		</tr>
	<tr>
		
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
		
		<%If Len(Trim(v_txtBlok)) > 0 And v_txtBlok <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<%End if %>
		
		<%If nSpantotal = 0 Then%><th rowspan="2" style="background-color:#00A300;color:#fff;">&nbsp;</th><%End If%>
		
		<th colspan="31" style="background-color:#00A300;color:#fff;">TANGGAL</th>
	</tr>
	<tr>
		<%for nday = 1 to 31%>
		<th style="background-color:#00A300;color:#fff;"><%=nday%></th>
		<%next%>
	</tr>
    </thead>
	<tbody>
<%
Set rsRestan = Server.CreateObject("ADODB.Recordset")
rsRestan.Open strRestan, connstr, adOpenStatic

Dim Subnday_1 : Subnday_1 = 0
Dim Subnday_2 : Subnday_2 = 0
Dim Subnday_3 : Subnday_3 = 0
Dim Subnday_4 : Subnday_4 = 0
Dim Subnday_5 : Subnday_5 = 0
Dim Subnday_6 : Subnday_6 = 0
Dim Subnday_7 : Subnday_7 = 0
Dim Subnday_8 : Subnday_8 = 0
Dim Subnday_9 : Subnday_9 = 0
Dim Subnday_10 : Subnday_10 = 0
Dim Subnday_11 : Subnday_11 = 0
Dim Subnday_12 : Subnday_12 = 0
Dim Subnday_13 : Subnday_13 = 0
Dim Subnday_14 : Subnday_14 = 0
Dim Subnday_15 : Subnday_15 = 0
Dim Subnday_16 : Subnday_16 = 0
Dim Subnday_17 : Subnday_17 = 0
Dim Subnday_18 : Subnday_18 = 0
Dim Subnday_19 : Subnday_19 = 0
Dim Subnday_20 : Subnday_20 = 0
Dim Subnday_21 : Subnday_21 = 0
Dim Subnday_22 : Subnday_22 = 0
Dim Subnday_23 : Subnday_23 = 0
Dim Subnday_24 : Subnday_24 = 0
Dim Subnday_25 : Subnday_25 = 0
Dim Subnday_26 : Subnday_26 = 0
Dim Subnday_27 : Subnday_27 = 0
Dim Subnday_28 : Subnday_28 = 0
Dim Subnday_29 : Subnday_29 = 0
Dim Subnday_30 : Subnday_30 = 0
Dim Subnday_31 : Subnday_31 = 0

Dim Grandnday_1 : Grandnday_1 = 0
Dim Grandnday_2 : Grandnday_2 = 0
Dim Grandnday_3 : Grandnday_3 = 0
Dim Grandnday_4 : Grandnday_4 = 0
Dim Grandnday_5 : Grandnday_5 = 0
Dim Grandnday_6 : Grandnday_6 = 0
Dim Grandnday_7 : Grandnday_7 = 0
Dim Grandnday_8 : Grandnday_8 = 0
Dim Grandnday_9 : Grandnday_9 = 0
Dim Grandnday_10 : Grandnday_10 = 0
Dim Grandnday_11 : Grandnday_11 = 0
Dim Grandnday_12 : Grandnday_12 = 0
Dim Grandnday_13 : Grandnday_13 = 0
Dim Grandnday_14 : Grandnday_14 = 0
Dim Grandnday_15 : Grandnday_15 = 0
Dim Grandnday_16 : Grandnday_16 = 0
Dim Grandnday_17 : Grandnday_17 = 0
Dim Grandnday_18 : Grandnday_18 = 0
Dim Grandnday_19 : Grandnday_19 = 0
Dim Grandnday_20 : Grandnday_20 = 0
Dim Grandnday_21 : Grandnday_21 = 0
Dim Grandnday_22 : Grandnday_22 = 0
Dim Grandnday_23 : Grandnday_23 = 0
Dim Grandnday_24 : Grandnday_24 = 0
Dim Grandnday_25 : Grandnday_25 = 0
Dim Grandnday_26 : Grandnday_26 = 0
Dim Grandnday_27 : Grandnday_27 = 0
Dim Grandnday_28 : Grandnday_28 = 0
Dim Grandnday_29 : Grandnday_29 = 0
Dim Grandnday_30 : Grandnday_30 = 0
Dim Grandnday_31 : Grandnday_31 = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempDivision : tempDivision = ""
Dim tempBlok : tempBlok = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sprevDivision : sprevDivision = ""
Dim sprevFieldCode : sprevFieldCode = ""

if rsRestan.bof and rsRestan.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+31%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsRestan.eof
%>
				<tr>
				
				<%If nSpantotal = 0 Then%><td align="center">&nbsp;</td><%End If%>
						
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsRestan("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsRestan("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRestan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsRestan("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsRestan("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRestan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsRestan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsRestan("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsRestan("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRestan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsRestan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsRestan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsRestan("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsRestan("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRestan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsRestan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsRestan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsRestan("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsRestan("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsRestan("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRestan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsRestan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsRestan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsRestan("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsRestan("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsRestan("fcba")
					   tempDivision = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
						    <%=rsRestan("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtAfdeling <> "--SUMMARY--" And v_txtAfdeling <> "" then %>
				   <%  if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRestan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsRestan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsRestan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsRestan("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsRestan("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtfcba <>  "--SUMMARY--" and v_txtfcba <> "" then
					       tempfcba = rsRestan("fcba")
					   else
					       tempfcba = ""
					   end if
					   
					   tempDivision = rsRestan("AFDELING")
					   tempBlok = "--ALL--"					   
					 %>
				
				
				    <% if tempDivision <> sprevDivision then %>
					    <td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsRestan("AFDELING")%>
							</a>
						</td>
						<% sprevDivision = tempDivision %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <>  "" then %>
				    <% if sprevFieldCode <> rsRestan("FIELDCODE") then %>				
					    <td align="center"><%=rsRestan("FIELDCODE")%></td>				
						<% sprevFieldCode = rsRestan("FIELDCODE")%>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
					<td align="right"><%=FormatNumber(rsRestan("DAY_1"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_2"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_3"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_4"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_5"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_6"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_7"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_8"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_9"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_10"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_11"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_12"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_13"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_14"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_15"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_16"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_17"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_18"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_19"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_20"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_21"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_22"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_23"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_24"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_25"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_26"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_27"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_28"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_29"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_30"), 2)%></td>
					<td align="right"><%=FormatNumber(rsRestan("DAY_31"), 2)%></td>
				</tr>
<%
			Grandnday_1 =  Grandnday_1 + Cdbl(rsRestan("DAY_1"))
			Grandnday_2 =  Grandnday_2 + Cdbl(rsRestan("DAY_2"))
			Grandnday_3 =  Grandnday_3 + Cdbl(rsRestan("DAY_3"))
			Grandnday_4 =  Grandnday_4 + Cdbl(rsRestan("DAY_4"))
			Grandnday_5 =  Grandnday_5 + Cdbl(rsRestan("DAY_5"))
			Grandnday_6 =  Grandnday_6 + Cdbl(rsRestan("DAY_6"))
			Grandnday_7 =  Grandnday_7 + Cdbl(rsRestan("DAY_7"))
			Grandnday_8 =  Grandnday_8 + Cdbl(rsRestan("DAY_8"))
			Grandnday_9 =  Grandnday_9 + Cdbl(rsRestan("DAY_9"))
			Grandnday_10 =  Grandnday_10 + Cdbl(rsRestan("DAY_10"))
			Grandnday_11 =  Grandnday_11 + Cdbl(rsRestan("DAY_11"))
			Grandnday_12 =  Grandnday_12 + Cdbl(rsRestan("DAY_12"))
			Grandnday_13 =  Grandnday_13 + Cdbl(rsRestan("DAY_13"))
			Grandnday_14 =  Grandnday_14 + Cdbl(rsRestan("DAY_14"))
			Grandnday_15 =  Grandnday_15 + Cdbl(rsRestan("DAY_15"))
			Grandnday_16 =  Grandnday_16 + Cdbl(rsRestan("DAY_16"))
			Grandnday_17 =  Grandnday_17 + Cdbl(rsRestan("DAY_17"))
			Grandnday_18 =  Grandnday_18 + Cdbl(rsRestan("DAY_18"))
			Grandnday_19 =  Grandnday_19 + Cdbl(rsRestan("DAY_19"))
			Grandnday_20 =  Grandnday_20 + Cdbl(rsRestan("DAY_20"))
			Grandnday_21 =  Grandnday_21 + Cdbl(rsRestan("DAY_21"))
			Grandnday_22 =  Grandnday_22 + Cdbl(rsRestan("DAY_22"))
			Grandnday_23 =  Grandnday_23 + Cdbl(rsRestan("DAY_23"))
			Grandnday_24 =  Grandnday_24 + Cdbl(rsRestan("DAY_24"))
			Grandnday_25 =  Grandnday_25 + Cdbl(rsRestan("DAY_25"))
			Grandnday_26 =  Grandnday_26 + Cdbl(rsRestan("DAY_26"))
			Grandnday_27 =  Grandnday_27 + Cdbl(rsRestan("DAY_27"))
			Grandnday_28 =  Grandnday_28 + Cdbl(rsRestan("DAY_28"))
			Grandnday_29 =  Grandnday_29 + Cdbl(rsRestan("DAY_29"))
			Grandnday_30 =  Grandnday_30 + Cdbl(rsRestan("DAY_30"))
			Grandnday_31 =  Grandnday_31 + Cdbl(rsRestan("DAY_31"))
			
			rsRestan.MoveNext
			
			Wend	
rsRestan.Close
set rsRestan = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td colspan="<%=var_colspan%>" align="center"><strong>GRAND TOTAL</strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_1)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_2)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_3)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_4)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_5)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_6)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_7)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_8)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_9)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_10)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_11)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_12)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_13)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_14)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_15)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_16)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_17)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_18)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_19)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_20)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_21)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_22)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_23)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_24)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_25)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_26)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_27)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_28)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_29)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_30)%></strong></td>
		<td align="right"><strong><%=FbpNum2(Grandnday_31)%></strong></td>
	</tr>
<%End if%>
	</tbody>
</table>

   </body>
   </html>

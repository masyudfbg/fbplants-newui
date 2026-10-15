<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtStockCode, v_txtAfdeling, v_txtTahunTanam, v_txtBlok
Dim ScriptName
Dim rsHarianPupuk, strHarianPupuk, strMain, strHarianPupuk2
Dim NO, cClass, nday, xNum

Dim var_colspan : var_colspan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

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

v_txtStockCode	=	Request("txtStockCode")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtTahunTanam	=	Request("txtTahunTanam")
v_txtBlok		=	Request("txtBlok")

If v_txtStockCode = "" Then
	v_txtStockCode	=	"--ALL--"
End If

If Len(Trim(v_txtStockCode)) > 0 Then
	if v_txtStockCode = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtStockCode = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And STOCKCODE = '" & v_txtStockCode & "'"
		orderCons	=	orderCons	& " "
	end if
End If

If Len(Trim(v_txtAfdeling)) > 0 Then
	If v_txtAfdeling = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
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
	Elseif v_txtBlok = "--ALL--" then
		selectCons	=	selectCons	& " , BLOK"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , BLOK"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , BLOK"
		whereCons	=	wherecons	& " And BLOK = '" & v_txtBlok & "'"
		orderCons	=	orderCons	& " , BLOK"
		var_colspan	=	var_colspan	+ 1
	End if
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT" & strFieldCons & _
			"         , STOCKCODE" & _
			"         , STOCKNAME" & _
			"         , AFDELING" & _
			"         , TAHUNTANAM" & _
			"         , BLOK" & _
			"         , ACTIVEMATERIAL" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 1, QTY_HI, 0)) / 1000 DAY_1" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 2, QTY_HI, 0)) / 1000 DAY_2" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 3, QTY_HI, 0)) / 1000 DAY_3" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 4, QTY_HI, 0)) / 1000 DAY_4" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 5, QTY_HI, 0)) / 1000 DAY_5" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 6, QTY_HI, 0)) / 1000 DAY_6" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 7, QTY_HI, 0)) / 1000 DAY_7" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 8, QTY_HI, 0)) / 1000 DAY_8" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 9, QTY_HI, 0)) / 1000 DAY_9" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 10, QTY_HI, 0)) / 1000 DAY_10" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 11, QTY_HI, 0)) / 1000 DAY_11" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 12, QTY_HI, 0)) / 1000 DAY_12" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 13, QTY_HI, 0)) / 1000 DAY_13" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 14, QTY_HI, 0)) / 1000 DAY_14" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 15, QTY_HI, 0)) / 1000 DAY_15" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 16, QTY_HI, 0)) / 1000 DAY_16" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 17, QTY_HI, 0)) / 1000 DAY_17" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 18, QTY_HI, 0)) / 1000 DAY_18" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 19, QTY_HI, 0)) / 1000 DAY_19" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 20, QTY_HI, 0)) / 1000 DAY_20" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 21, QTY_HI, 0)) / 1000 DAY_21" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 22, QTY_HI, 0)) / 1000 DAY_22" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 23, QTY_HI, 0)) / 1000 DAY_23" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 24, QTY_HI, 0)) / 1000 DAY_24" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 25, QTY_HI, 0)) / 1000 DAY_25" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 26, QTY_HI, 0)) / 1000 DAY_26" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 27, QTY_HI, 0)) / 1000 DAY_27" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 28, QTY_HI, 0)) / 1000 DAY_28" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 29, QTY_HI, 0)) / 1000 DAY_29" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 30, QTY_HI, 0)) / 1000 DAY_30" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'DD')), 31, QTY_HI, 0)) / 1000 DAY_31" & _
			"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL,'MM')), " & v_txtMonth & ", QTY_HI, TO_NUMBER(TO_CHAR(TANGGAL,'YYYY')), " & v_txtYear & ", QTY_HI, 0)) / 1000 TOTAL" & _
			"         , SUM(BUDGET_HI) BUDGET" & _
			"     FROM (    " & _
			"         SELECT " & _
			"         FCBA" & _
			"         , TANGGAL" & _
			"         , STOCKCODE" & _
			"         , STOCKNAME" & _
			"         , AFDELING" & _
			"         , TAHUNTANAM" & _
			"         , BLOK" & _
			"         , ACTIVEMATERIAL" & _
			"         , QTY_HI" & _
			"         , BUDGET_HI" & _
			"         FROM" & _
			"             REPORT_PUPUK_DAILY_TAB" & _
			"         WHERE" & _
			"             TO_NUMBER(TO_CHAR(TANGGAL,'YYYY')) = " & v_txtYear & "" & _
			"             AND TO_NUMBER(TO_CHAR(TANGGAL,'MM')) = " & v_txtMonth & "" & _
			"         )" & _
			"     GROUP BY" & _
			"         FCBA" & _
			"         , STOCKCODE" & _
			"         , STOCKNAME" & _
			"         , AFDELING" & _
			"         , TAHUNTANAM" & _
			"         , BLOK" & _
			"         , ACTIVEMATERIAL"
								
strHarianPupuk2	=	"     , STOCKCODE" & _
					"     , STOCKNAME" & _
					"     , ACTIVEMATERIAL" & _
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
					"     , SUM(TOTAL) TOTAL" & _
					"     , SUM(BUDGET) BUDGET"

strHarianPupuk	=	" select GroupPlantation " & selectCons & strHarianPupuk2 &"  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation" & selectCons & ", STOCKCODE, STOCKNAME, ACTIVEMATERIAL" & _
					" Order by GroupPlantation " & orderCons & ", ACTIVEMATERIAL"

'response.write strHarianPupuk
%>
<html>

<head>
<title>ePlaS Reporting - Harian Pemupukan</title>
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
							<font face="verdana" size="1">STOCK CODE</font>
						</td>
						<td>
							<select name="txtStockCode">
								<option value="--ALL--"<%If v_txtStockCode = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsStockCode, strStockCode
								strStockCode	=	" SELECT DISTINCT FCCODE, FCNAME FROM STOCK WHERE STOCKGROUPCODE = '302'" & _
													" ORDER BY FCCODE"
								
								Set rsStockCode = Server.CreateObject("ADODB.Recordset")
								rsStockCode.Open strStockCode, connstr, adOpenStatic 
								
								if not rsStockCode.eof then
								Do While Not rsStockCode.eof
								if rsStockCode("FCCODE") = v_txtStockCode then%>
								<option value="<%=rsStockCode("FCCODE")%>" Selected><%=rsStockCode("FCCODE")%> : <%=rsStockCode("FCNAME")%></option> 
								<%Else%>
								<option value="<%=rsStockCode("FCCODE")%>"><%=rsStockCode("FCCODE")%> : <%=rsStockCode("FCNAME")%></option>  
								<%End if
									rsStockCode.movenext
								loop
								end if
								rsStockCode.Close
								set rsStockCode = Nothing
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
								<option value="BIBITAN"<%If v_txtAfdeling = "BIBITAN" then%> Selected <%End If%>>BIBITAN</option>
								<option value="LAIN-LAIN"<%If v_txtAfdeling = "LAIN-LAIN" then%> Selected <%End If%>>LAIN-LAIN</option>
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
								<option value="BIBITAN"<%If v_txtAfdeling = "BIBITAN" then%> Selected <%End If%>>BIBITAN</option>
								<option value="LAIN-LAIN"<%If v_txtAfdeling = "LAIN-LAIN" then%> Selected <%End If%>>LAIN-LAIN</option>
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
								<option value="BIBITAN"<%If v_txtAfdeling = "BIBITAN" then%> Selected <%End If%>>BIBITAN</option>
								<option value="LAIN-LAIN"<%If v_txtAfdeling = "LAIN-LAIN" then%> Selected <%End If%>>LAIN-LAIN</option>
								<%
								Dim rsBlok, strBlok
								strBlok	=	" SELECT DISTINCT FCCODE, FCNAME FROM FIELD WHERE FCBA = '" & Session("Fcba") & "' AND ACTIVATION = 'Y' AND DX_BLOCK_STATUS = 'TM' ORDER BY FCCODE"
								
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
         <th colspan="<%=var_colspan+35%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN HARIAN PEMUPUKAN</h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		 </th>
	</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">JENIS PUPUK</th>
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
		<th rowspan="2" style="background-color:#00A300;color:#fff;">REAL (TON)</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">REKOM (TON)</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">%</th>
	</tr>
	<tr>
		<%for nday = 1 to 31%>
		<th style="background-color:#00A300;color:#fff;"><%=nday%></th>
		<%next%>
	</tr>
    </thead>
	<tbody>
<%
Set rsHarianPupuk = Server.CreateObject("ADODB.Recordset")
rsHarianPupuk.Open strHarianPupuk, connstr, adOpenStatic

Dim RowRealTotal : RowRealTotal = 0
Dim RowRekomTotal : RowRekomTotal = 0
Dim RowVarTotal : RowVarTotal = 0

Dim GrandTotalday_1 : GrandTotalday_1 = 0
Dim GrandTotalday_2 : GrandTotalday_2 = 0
Dim GrandTotalday_3 : GrandTotalday_3 = 0
Dim GrandTotalday_4 : GrandTotalday_4 = 0
Dim GrandTotalday_5 : GrandTotalday_5 = 0
Dim GrandTotalday_6 : GrandTotalday_6 = 0
Dim GrandTotalday_7 : GrandTotalday_7 = 0
Dim GrandTotalday_8 : GrandTotalday_8 = 0
Dim GrandTotalday_9 : GrandTotalday_9 = 0
Dim GrandTotalday_10 : GrandTotalday_10 = 0
Dim GrandTotalday_11 : GrandTotalday_11 = 0
Dim GrandTotalday_12 : GrandTotalday_12 = 0
Dim GrandTotalday_13 : GrandTotalday_13 = 0
Dim GrandTotalday_14 : GrandTotalday_14 = 0
Dim GrandTotalday_15 : GrandTotalday_15 = 0
Dim GrandTotalday_16 : GrandTotalday_16 = 0
Dim GrandTotalday_17 : GrandTotalday_17 = 0
Dim GrandTotalday_18 : GrandTotalday_18 = 0
Dim GrandTotalday_19 : GrandTotalday_19 = 0
Dim GrandTotalday_20 : GrandTotalday_20 = 0
Dim GrandTotalday_21 : GrandTotalday_21 = 0
Dim GrandTotalday_22 : GrandTotalday_22 = 0
Dim GrandTotalday_23 : GrandTotalday_23 = 0
Dim GrandTotalday_24 : GrandTotalday_24 = 0
Dim GrandTotalday_25 : GrandTotalday_25 = 0
Dim GrandTotalday_26 : GrandTotalday_26 = 0
Dim GrandTotalday_27 : GrandTotalday_27 = 0
Dim GrandTotalday_28 : GrandTotalday_28 = 0
Dim GrandTotalday_29 : GrandTotalday_29 = 0
Dim GrandTotalday_30 : GrandTotalday_30 = 0
Dim GrandTotalday_31 : GrandTotalday_31 = 0
Dim GrandRowRealTotal : GrandRowRealTotal = 0
Dim GrandRowRekomTotal : GrandRowRekomTotal = 0
Dim GrandRowVarTotal : GrandRowVarTotal = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempStockCode : tempStockCode = ""
Dim tempAfdeling : tempAfdeling = ""
Dim tempTahunTanam : tempTahunTanam = ""
Dim tempLocationCode : tempLocationCode = ""


Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sprevStockCode : sprevStockCode = ""
Dim sprevAfdeling : sprevAfdeling = ""
Dim sprevTahunTanam : sprevTahunTanam = ""
Dim sprevLocationCode : sprevLocationCode = ""

if rsHarianPupuk.bof and rsHarianPupuk.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+35%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsHarianPupuk.eof
%>
				<tr>
				<% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <>  "" then %>
					<%
					   tempStockCode = rsHarianPupuk("STOCKCODE")
					   tempGroupPlantation = "--ALL--"
					   tempPlantation = "--SUMMARY--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					<% if tempStockCode <> SprevStockCode then %>
						<td align="left">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
							<%=rsHarianPupuk("STOCKNAME")%>
							</a>
						</td>
						<% SprevStockCode = tempStockCode %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
							<%=rsHarianPupuk("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsHarianPupuk("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
							<%=rsHarianPupuk("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsHarianPupuk("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
							<%=rsHarianPupuk("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsHarianPupuk("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsHarianPupuk("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
							<%=rsHarianPupuk("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianPupuk("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianPupuk("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsHarianPupuk("Company")
					   tempfcba = "--ALL--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
							<%=rsHarianPupuk("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianPupuk("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianPupuk("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianPupuk("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsHarianPupuk("fcba")
					   tempAfdeling = "--ALL--"
					   tempTahunTanam = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
						    <%=rsHarianPupuk("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtAfdeling <>  "--SUMMARY--" and v_txtAfdeling <>  "" then %>
				    <% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianPupuk("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianPupuk("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianPupuk("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsHarianPupuk("fcba")
					   else
					       tempfcba =""
					   end if

					   tempAfdeling = rsHarianPupuk("AFDELING")
					   tempTahunTanam = "--ALL--"
					   tempLocationCode = "--SUMMARY--"
					 %>
					
					<% if tempAfdeling <> sprevAfdeling then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
						    <%=rsHarianPupuk("AFDELING")%>
							</a>
						</td>
						<% sprevAfdeling = tempAfdeling %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>

				<% if v_txtTahunTanam <>  "--SUMMARY--" and v_txtTahunTanam <>  "" then %>
				    <% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("STOCKCODE")
					   else
					      tempStockCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianPupuk("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianPupuk("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianPupuk("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsHarianPupuk("fcba")
					   else
					       tempfcba =""
					   end if
					   
					   if  v_txtAfdeling  <>  "--SUMMARY--" and v_txtAfdeling <> "" then
					       tempAfdeling = rsHarianPupuk("AFDELING")
					   else
					       tempAfdeling =""
					   end if

					   tempTahunTanam = rsHarianPupuk("TAHUNTANAM")
					   tempLocationCode = "--ALL--"
					 %>
					
					<% if tempTahunTanam <> sprevTahunTanam then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtStockCode=<%=tempStockCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempLocationCode%>">
						    <%=rsHarianPupuk("TAHUNTANAM")%>
							</a>
						</td>
						<% sprevTahunTanam = tempTahunTanam %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <>  "" then %>
				    <% if sprevLocationCode <> rsHarianPupuk("BLOK") then %>				
					    <td align="center"><%=rsHarianPupuk("BLOK")%></td>				
						<% sprevLocationCode = rsHarianPupuk("BLOK")%>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_1"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_2"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_3"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_4"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_5"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_6"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_7"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_8"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_9"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_10"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_11"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_12"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_13"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_14"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_15"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_16"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_17"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_18"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_19"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_20"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_21"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_22"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_23"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_24"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_25"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_26"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_27"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_28"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_29"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_30"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("DAY_31"), 2)%></td>
<%
			RowRealTotal	=	RowRealTotal + Cdbl(rsHarianPupuk("TOTAL"))
			RowRekomTotal	=	RowRekomTotal + Cdbl(rsHarianPupuk("BUDGET"))
			If RowRekomTotal = 0 Then
				RowVarTotal	=	0	
			Else
				RowVarTotal	=	((RowRealTotal / RowRekomTotal)* 100)
			End If
%>
					<td align="right"><%=FormatNumber(RowRealTotal, 2)%></td>
					<td align="right"><%=FormatNumber(RowRekomTotal, 2)%></td>
					<td align="right"><%=FormatNumber(RowVarTotal, 2)%></td>
				</tr>
<%

			GrandTotalday_1 =  GrandTotalday_1 + Cdbl(rsHarianPupuk("DAY_1"))
			GrandTotalday_2 =  GrandTotalday_2 + Cdbl(rsHarianPupuk("DAY_2"))
			GrandTotalday_3 =  GrandTotalday_3 + Cdbl(rsHarianPupuk("DAY_3"))
			GrandTotalday_4 =  GrandTotalday_4 + Cdbl(rsHarianPupuk("DAY_4"))
			GrandTotalday_5 =  GrandTotalday_5 + Cdbl(rsHarianPupuk("DAY_5"))
			GrandTotalday_6 =  GrandTotalday_6 + Cdbl(rsHarianPupuk("DAY_6"))
			GrandTotalday_7 =  GrandTotalday_7 + Cdbl(rsHarianPupuk("DAY_7"))
			GrandTotalday_8 =  GrandTotalday_8 + Cdbl(rsHarianPupuk("DAY_8"))
			GrandTotalday_9 =  GrandTotalday_9 + Cdbl(rsHarianPupuk("DAY_9"))
			GrandTotalday_10 =  GrandTotalday_10 + Cdbl(rsHarianPupuk("DAY_10"))
			GrandTotalday_11 =  GrandTotalday_11 + Cdbl(rsHarianPupuk("DAY_11"))
			GrandTotalday_12 =  GrandTotalday_12 + Cdbl(rsHarianPupuk("DAY_12"))
			GrandTotalday_13 =  GrandTotalday_13 + Cdbl(rsHarianPupuk("DAY_13"))
			GrandTotalday_14 =  GrandTotalday_14 + Cdbl(rsHarianPupuk("DAY_14"))
			GrandTotalday_15 =  GrandTotalday_15 + Cdbl(rsHarianPupuk("DAY_15"))
			GrandTotalday_16 =  GrandTotalday_16 + Cdbl(rsHarianPupuk("DAY_16"))
			GrandTotalday_17 =  GrandTotalday_17 + Cdbl(rsHarianPupuk("DAY_17"))
			GrandTotalday_18 =  GrandTotalday_18 + Cdbl(rsHarianPupuk("DAY_18"))
			GrandTotalday_19 =  GrandTotalday_19 + Cdbl(rsHarianPupuk("DAY_19"))
			GrandTotalday_20 =  GrandTotalday_20 + Cdbl(rsHarianPupuk("DAY_20"))
			GrandTotalday_21 =  GrandTotalday_21 + Cdbl(rsHarianPupuk("DAY_21"))
			GrandTotalday_22 =  GrandTotalday_22 + Cdbl(rsHarianPupuk("DAY_22"))
			GrandTotalday_23 =  GrandTotalday_23 + Cdbl(rsHarianPupuk("DAY_23"))
			GrandTotalday_24 =  GrandTotalday_24 + Cdbl(rsHarianPupuk("DAY_24"))
			GrandTotalday_25 =  GrandTotalday_25 + Cdbl(rsHarianPupuk("DAY_25"))
			GrandTotalday_26 =  GrandTotalday_26 + Cdbl(rsHarianPupuk("DAY_26"))
			GrandTotalday_27 =  GrandTotalday_27 + Cdbl(rsHarianPupuk("DAY_27"))
			GrandTotalday_28 =  GrandTotalday_28 + Cdbl(rsHarianPupuk("DAY_28"))
			GrandTotalday_29 =  GrandTotalday_29 + Cdbl(rsHarianPupuk("DAY_29"))
			GrandTotalday_30 =  GrandTotalday_30 + Cdbl(rsHarianPupuk("DAY_30"))
			GrandTotalday_31 =  GrandTotalday_31 + Cdbl(rsHarianPupuk("DAY_31"))
			GrandRowRealTotal = GrandRowRealTotal + RowRealTotal
			GrandRowRekomTotal = GrandRowRekomTotal + RowRekomTotal
			GrandRowVarTotal = GrandRowVarTotal + RowVarTotal
			
			
			rsHarianPupuk.MoveNext
			
			RowRealTotal	=	0
			RowRekomTotal	=	0
			RowVarTotal		=	0
			Wend	
rsHarianPupuk.Close
Set rsHarianPupuk = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td colspan="<%=var_colspan+1%>" align="center"><strong>TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_1, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_2, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_3, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_4, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_5, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_6, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_7, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_8, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_9, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_10, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_11, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_12, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_13, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_14, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_15, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_16, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_17, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_18, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_19, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_20, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_21, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_22, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_23, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_24, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_25, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_26, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_27, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_28, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_29, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_30, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalday_31, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandRowRealTotal, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandRowRekomTotal, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandRowVarTotal, 2)%></strong></td>
	</tr>
<%End if%>
	</tbody>
</table>

   </body>
   </html>
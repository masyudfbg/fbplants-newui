<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtStockCode, v_txtAfdeling, v_txtTahunTanam, v_txtBlok,v_txtStatus
Dim ScriptName
Dim rsHarianPupuk, strHarianPupuk, strMain, strHarianPupuk2
Dim cClass, nday, xNum
Dim strRowsCount,rsSqlRow

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtStatus				= Request("txtStatus")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons, whereConsTtl, whereConsRow

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons		=	" where company is not null "
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
v_txtOwnership		=	Request("txtOwnership")

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

If Len(Trim(v_txtOwnership)) > 0 Then
	If v_txtOwnership = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtOwnership = "--ALL--" then
		selectCons	=	selectCons	& " , ownership"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , ownership"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , ownership"
		whereCons	=	wherecons	& " And ownership = '" & v_txtOwnership & "'"
		orderCons	=	orderCons	& " , ownership"
		var_colspan	=	var_colspan	+ 1
	End if
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim whereConsMain, TitleTmTbm

' If v_txtOwnership = "--ALL--" Then
' 	whereConsMain 	= ""
' 	TitleTmTbm	  	= ""
' Else	
' 	whereConsMain = "AND FIELD_API.GET_OWENERSHIP (rp.blok, rp.fcba) = '"& v_txtOwnership &"'"
' End if 

If v_txtStatus = "--ALL--" then
	whereConsMain 	= ""
	TitleTmTbm	  	= ""
Else
	if v_txtStatus = "TMTBM" then
		whereConsMain = " AND rp.DX_BLOCK_STATUS IN ('TM', 'TBM')"
		TitleTmTbm = "TM dan TBM"
	else
		whereConsMain = " AND rp.DX_BLOCK_STATUS = '"& v_txtStatus &"'"
		TitleTmTbm = v_txtStatus
	end if
End if



strMain	=	" SELECT" & strFieldCons & _
			"         , ownership" & _
			"         , stockname" & _
			"         , afdeling" & _
			"         , tahuntanam" & _
			"         , blok" & _
			"         , SUM(DECODE(bulan, 1, qty_bi, 0)) month_1" & _
			"         , SUM(DECODE(bulan, 2, qty_bi, 0))  month_2" & _
			"         , SUM(DECODE(bulan, 3, qty_bi, 0)) month_3" & _
			"         , SUM(DECODE(bulan, 4, qty_bi, 0))  month_4" & _
			"         , SUM(DECODE(bulan, 5, qty_bi, 0)) month_5" & _
			"         , SUM(DECODE(bulan, 6, qty_bi, 0))  month_6" & _
			"         , SUM(DECODE(bulan, 7, qty_bi, 0))  month_7" & _
			"         , SUM(DECODE(bulan, 8, qty_bi, 0))  month_8" & _
			"         , SUM(DECODE(bulan, 9, qty_bi, 0)) month_9" & _
			"         , SUM(DECODE(bulan, 10, qty_bi, 0))  month_10" & _
			"         , SUM(DECODE(bulan, 11, qty_bi, 0))  month_11" & _
			"         , SUM(DECODE(bulan, 12, qty_bi, 0))  month_12" & _
			"         , SUM(DECODE(tahun, " & v_txtyear & ", qty_bi, 0)) total" & _
			"         , SUM(DECODE(tahun, " & v_txtYear & ", budget_bi, 0))  budget" & _
			"     FROM (    " & _
			"         SELECT rp.fcba" & _
			"                , rp.tahun" & _
			"                , rp.bulan" & _
			"                , rp.stockcode" & _
			"                , pr.fcname stockname" & _
			"                , rp.afdeling" & _
			"				 , FIELD_API.GET_OWENERSHIP (rp.blok, rp.fcba)ownership" & _
			"                , rp.tahuntanam" & _
			"                , rp.blok" & _
			"                , rp.qty_bi" & _
			"                , rp.budget_bi" & _
			"                , rp.rekom_bi" & _
			"         FROM report_pupuk_month_tab rp" & _
			"         INNER JOIN (SELECT fccode, fcname, activematerial" & _
            "                  FROM stock " & _
            "                 WHERE (STOCKGROUPCODE IN ('305')) AND fcba = 'KNS') pr" & _
            "         ON pr.fccode = rp.stockcode" & _
			"         WHERE rp.tahun = " & v_txtYear & "" & _
			"               AND rp.bulan <= " & v_txtMonth & "" & _
			"          " & whereConsMain & ")" & _
			"     GROUP BY fcba" & _
			"              , stockname" & _
			"              , afdeling" & _
			"              , ownership" & _
			"              , tahuntanam" & _
			"              , blok"
								
strHarianPupuk2	=	"     , stockname" & _
					"     , SUM(MONTH_1) MONTH_1" & _
					"     , SUM(MONTH_2) MONTH_2" & _
					"     , SUM(MONTH_3) MONTH_3" & _
					"     , SUM(MONTH_4) MONTH_4" & _
					"     , SUM(MONTH_5) MONTH_5" & _
					"     , SUM(MONTH_6) MONTH_6" & _
					"     , SUM(MONTH_7) MONTH_7" & _
					"     , SUM(MONTH_8) MONTH_8" & _
					"     , SUM(MONTH_9) MONTH_9" & _
					"     , SUM(MONTH_10) MONTH_10" & _
					"     , SUM(MONTH_11) MONTH_11" & _
					"     , SUM(MONTH_12) MONTH_12" & _
					"     , SUM(TOTAL) TOTAL" & _
					"     , SUM(BUDGET) BUDGET"

strHarianPupuk	=	"select GroupPlantation " & selectCons & strHarianPupuk2 &"  From (" & strMain & " )" & _
					whereCons & _
					" group by GroupPlantation" & selectCons & ", stockname" & _
					" having (SUM(TOTAL) + SUM(BUDGET) > 0)" & _
					" Order by GroupPlantation " & orderCons & ", stockname"

' response.write strHarianPupuk

Dim strMainHK, strHK, strHKEx,rsHKEx
strMainHK	=	" SELECT" & strFieldCons & _
				"         , ownership" & _
				"         , JOBCODE" & _
				"         , JOBNAME" & _
				"         , AFDELING" & _
				"         , TAHUNTANAM" & _
				"         , BLOK" & _
				"         , ACTIVEMATERIAL" & _
				"         , SUM(DECODE(BULAN, 1, QTY_BI, 0)) MONTH_1" & _
				"         , SUM(DECODE(BULAN, 2, QTY_BI, 0)) MONTH_2" & _
				"         , SUM(DECODE(BULAN, 3, QTY_BI, 0)) MONTH_3" & _
				"         , SUM(DECODE(BULAN, 4, QTY_BI, 0)) MONTH_4" & _
				"         , SUM(DECODE(BULAN, 5, QTY_BI, 0)) MONTH_5" & _
				"         , SUM(DECODE(BULAN, 6, QTY_BI, 0)) MONTH_6" & _
				"         , SUM(DECODE(BULAN, 7, QTY_BI, 0)) MONTH_7" & _
				"         , SUM(DECODE(BULAN, 8, QTY_BI, 0)) MONTH_8" & _
				"         , SUM(DECODE(BULAN, 9, QTY_BI, 0)) MONTH_9" & _
				"         , SUM(DECODE(BULAN, 10, QTY_BI, 0)) MONTH_10" & _
				"         , SUM(DECODE(BULAN, 11, QTY_BI, 0)) MONTH_11" & _
				"         , SUM(DECODE(BULAN, 12, QTY_BI, 0)) MONTH_12" & _
				"         , SUM(DECODE(TAHUN, " & v_txtYear & ", QTY_BI, 0)) TOTAL" & _
				"     FROM (    " & _
				"         SELECT fcba" & _
				"                , tahun" & _
				"                , bulan" & _
				"                , jobcode" & _
				"                , jobname" & _
				"                , afdeling" & _
				"				 , FIELD_API.GET_OWENERSHIP (blok, fcba)ownership" & _
				"                , tahuntanam" & _
				"                , blok" & _
				"                , activematerial" & _
				"                , qty_bi" & _
				"                , budget_bi" & _
				"                , rekom_bi" & _
				"         FROM" & _
				"             REPORT_PUPUKMANDAYS_MONTH_TAB rp" & _
				"         WHERE" & _
				"             TAHUN = " & v_txtYear & "" & _
				"             AND BULAN <= " & v_txtMonth & "" & _
				"          " & whereConsMain & ")" & _
				"     GROUP BY" & _
				"         FCBA" & _
				"         , JOBCODE" & _
				"         , ownership" & _
				"         , JOBNAME" & _
				"         , AFDELING" & _
				"         , TAHUNTANAM" & _
				"         , BLOK" & _
				"         , ACTIVEMATERIAL"
						
	strHK	=	"     , JOBCODE" & _
				"     , JOBNAME" & _
				"     , ACTIVEMATERIAL" & _
				"     , SUM(MONTH_1) MONTH_1" & _
				"     , SUM(MONTH_2) MONTH_2" & _
				"     , SUM(MONTH_3) MONTH_3" & _
				"     , SUM(MONTH_4) MONTH_4" & _
				"     , SUM(MONTH_5) MONTH_5" & _
				"     , SUM(MONTH_6) MONTH_6" & _
				"     , SUM(MONTH_7) MONTH_7" & _
				"     , SUM(MONTH_8) MONTH_8" & _
				"     , SUM(MONTH_9) MONTH_9" & _
				"     , SUM(MONTH_10) MONTH_10" & _
				"     , SUM(MONTH_11) MONTH_11" & _
				"     , SUM(MONTH_12) MONTH_12" & _
				"     , SUM(TOTAL) TOTAL"
%>
<html>
<head>
<title>ePlaS Reporting - Bulanan Pemupukan</title>
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
							<font face="verdana" size="1">STATUS</font>
						</td>
						<td width="2">
							<select name="txtStatus">
								<option value="--ALL--"<%If v_txtStatus = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<option value="TMTBM"<%If v_txtStatus = "TMTBM" then%> Selected <%End If%>>TM & TBM</option>
								<option value="TM"<%If v_txtStatus = "TM" then%> Selected <%End If%>>TM</option>
								<option value="TBM"<%If v_txtStatus = "TBM" then%> Selected <%End If%>>TBM</option>
								<option value="TB"<%If v_txtStatus = "TB" then%> Selected <%End If%>>TB</option>
								<option value="MN"<%If v_txtStatus = "MN" then%> Selected <%End If%>>MN</option>
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
								strStockCode	=	" SELECT DISTINCT fccode, fcname FROM stock WHERE ( fccode = '401.102.030' OR stockgroupcode IN ('302', '402') )" & _
													" ORDER BY fccode"
								
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
							<font face="verdana" size="1">OWNERSHIP</font>
						</td>
						<td>
							<select name="txtOwnership" >
								<option value="--SUMMARY--"<%If v_txtOwnership = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtOwnership = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<option value="BIBITAN"<%If v_txtAfdeling = "BIBITAN" then%> Selected <%End If%>>BIBITAN</option>
								<option value="LAIN-LAIN"<%If v_txtAfdeling = "LAIN-LAIN" then%> Selected <%End If%>>LAIN-LAIN</option>
								<% 
								if len(trim(v_txtOwnership)) > 0 Then
									strBlok	=	"  select distinct ownership from field where fcba = '"& v_txtfcba &"' "
								else
									strBlok	=	"  select distinct ownership from field where fcba = '"& Session("fcba") &"' "
								end if 

								Set rsBlok = Server.CreateObject("ADODB.Recordset")
								rsBlok.Open strBlok, connstr, adOpenStatic 
								
								if not rsBlok.eof then
								Do While Not rsBlok.eof
								if rsBlok("ownership") = v_txtOwnership then%>
								<option value="<%=rsBlok("ownership")%>" Selected><%=rsBlok("ownership")%></option> 
								<%Else%>
								<option value="<%=rsBlok("ownership")%>"><%=rsBlok("ownership")%></option>  
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
         <th colspan="<%=var_colspan+17%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN BULANAN PEMUPUKAN <%= TitleTmTbm %></h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		 </th>
	</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Jenis Pupuk</th>
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
		
		<%If Len(Trim(v_txtOwnership)) > 0 And v_txtOwnership <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">OWNERSHIP</th>
		<%End if %>

		<%If Len(Trim(v_txtAfdeling)) > 0 And v_txtAfdeling <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<%End if %>
		
		<%If Len(Trim(v_txtTahunTanam)) > 0 And v_txtTahunTanam <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">TAHUN TANAM</th>
		<%End if %>
		
		<%If Len(Trim(v_txtBlok)) > 0 And v_txtBlok <>  "--SUMMARY--" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<%End if %>
		
		<th colspan="12" style="background-color:#00A300;color:#fff;">Bulan</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Real (Kg)</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Rekom (Kg)</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">%</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Real (HK)</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">Jan</th>
		<th style="background-color:#00A300;color:#fff;">Feb</th>
		<th style="background-color:#00A300;color:#fff;">Mar</th>
		<th style="background-color:#00A300;color:#fff;">Apr</th>
		<th style="background-color:#00A300;color:#fff;">Mei</th>
		<th style="background-color:#00A300;color:#fff;">Jun</th>
		<th style="background-color:#00A300;color:#fff;">Jul</th>
		<th style="background-color:#00A300;color:#fff;">Agu</th>
		<th style="background-color:#00A300;color:#fff;">Sep</th>
		<th style="background-color:#00A300;color:#fff;">Okt</th>
		<th style="background-color:#00A300;color:#fff;">Nov</th>
		<th style="background-color:#00A300;color:#fff;">Des</th>
	</tr>
    </thead>
	<tbody>
<%
Set rsHarianPupuk = Server.CreateObject("ADODB.Recordset")
rsHarianPupuk.Open strHarianPupuk, connstr, adOpenStatic
' response.write strHarianPupuk
Dim MONTH_1,MONTH_2,MONTH_3,MONTH_4,MONTH_5,MONTH_6,MONTH_7,MONTH_8,MONTH_9,MONTH_10,MONTH_11,MONTH_12, TOTAL

Dim RowRealTotal : RowRealTotal = 0
Dim RowRekomTotal : RowRekomTotal = 0
Dim RowVarTotal : RowVarTotal = 0

Dim GrandTotalMonth_1 : GrandTotalMonth_1 = 0
Dim GrandTotalMonth_2 : GrandTotalMonth_2 = 0
Dim GrandTotalMonth_3 : GrandTotalMonth_3 = 0
Dim GrandTotalMonth_4 : GrandTotalMonth_4 = 0
Dim GrandTotalMonth_5 : GrandTotalMonth_5 = 0
Dim GrandTotalMonth_6 : GrandTotalMonth_6 = 0
Dim GrandTotalMonth_7 : GrandTotalMonth_7 = 0
Dim GrandTotalMonth_8 : GrandTotalMonth_8 = 0
Dim GrandTotalMonth_9 : GrandTotalMonth_9 = 0
Dim GrandTotalMonth_10 : GrandTotalMonth_10 = 0
Dim GrandTotalMonth_11 : GrandTotalMonth_11 = 0
Dim GrandTotalMonth_12 : GrandTotalMonth_12 = 0
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
		<td colspan="<%=var_colspan+17%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		Dim NOX : NOX = 1
		while not rsHarianPupuk.eof 
		whereConsTtl 	= " where company is not null "
		whereConsRow	= " "
				
				If Len(Trim(v_txtGroupPlantation)) > 0 Then
					If v_txtGroupPlantation = "--SUMMARY--" then
						whereConsTtl 	=	whereConsTtl	& " "
						whereConsRow	= 	whereConsRow
					Elseif v_txtGroupPlantation = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And GroupPlantation = '" & rsHarianPupuk("GroupPlantation") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where GroupPlantation = '" & rsHarianPupuk("GroupPlantation") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND GroupPlantation = '" & rsHarianPupuk("GroupPlantation") & "'"
						End IF 
					Else
						whereConsTtl =	whereConsTtl	& " And GroupPlantation = '" & rsHarianPupuk("GroupPlantation") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where GroupPlantation = '" & rsHarianPupuk("GroupPlantation") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND GroupPlantation = '" & rsHarianPupuk("GroupPlantation") & "'"
						End IF
					End if
				End If
				
				If Len(Trim(v_txtPlantation)) > 0 Then
					If v_txtPlantation = "--SUMMARY--" then
						whereConsTtl 	=	whereConsTtl	& " "
						whereConsRow	= 	whereConsRow
					Elseif v_txtPlantation = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And Plantation = '" & rsHarianPupuk("Plantation") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where Plantation = '" & rsHarianPupuk("Plantation") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND Plantation = '" & rsHarianPupuk("Plantation") & "'"
						End IF 
					Else
						whereConsTtl =	whereConsTtl	& " And Plantation = '" & rsHarianPupuk("Plantation") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where Plantation = '" & rsHarianPupuk("Plantation") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND Plantation = '" & rsHarianPupuk("Plantation") & "'"
						End IF
					End if
				End If
				
				If Len(Trim(v_txtGroupRegion)) > 0 Then
					If v_txtGroupRegion = "--SUMMARY--" then
						whereConsTtl 	=	whereConsTtl	& " "
						whereConsRow	= 	whereConsRow
					Elseif v_txtGroupRegion = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And GroupRegion = '" & rsHarianPupuk("GroupRegion") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where GroupRegion = '" & rsHarianPupuk("GroupRegion") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND GroupRegion = '" & rsHarianPupuk("GroupRegion") & "'"
						End IF 
					Else
						whereConsTtl =	whereConsTtl	& " And GroupRegion = '" & rsHarianPupuk("GroupRegion") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where GroupRegion = '" & rsHarianPupuk("GroupRegion") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND GroupRegion = '" & rsHarianPupuk("GroupRegion") & "'"
						End IF
					End if
				End If
				
				If Len(Trim(v_txtRegion)) > 0 Then
					If v_txtRegion = "--SUMMARY--" then
						whereConsTtl 	=	whereConsTtl	& " "
						whereConsRow	= 	whereConsRow
					Elseif v_txtRegion = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And REGION = '" & rsHarianPupuk("REGION") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where REGION = '" & rsHarianPupuk("REGION") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND REGION = '" & rsHarianPupuk("REGION") & "'"
						End IF 
					Else
						whereConsTtl =	whereConsTtl	& " And REGION = '" & rsHarianPupuk("REGION") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where REGION = '" & rsHarianPupuk("REGION") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND REGION = '" & rsHarianPupuk("REGION") & "'"
						End IF
					End if
				End If
				
				If Len(Trim(v_txtCompany)) > 0 Then
					If v_txtCompany = "--SUMMARY--" then
						whereConsTtl 	=	whereConsTtl	& " "
						whereConsRow	= 	whereConsRow
					Elseif v_txtCompany = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And COMPANY = '" & rsHarianPupuk("COMPANY") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where COMPANY = '" & rsHarianPupuk("COMPANY") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND COMPANY = '" & rsHarianPupuk("COMPANY") & "'"
						End IF 
					Else
						whereConsTtl =	whereConsTtl	& " And COMPANY = '" & rsHarianPupuk("COMPANY") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where COMPANY = '" & rsHarianPupuk("COMPANY") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND COMPANY = '" & rsHarianPupuk("COMPANY") & "'"
						End IF
					End if
				End If
				
				If Len(Trim(v_txtfcba)) > 0 Then
					If v_txtfcba = "--SUMMARY--" then
						whereConsTtl 	=	whereConsTtl	& " "
						whereConsRow	= 	whereConsRow
					Elseif v_txtfcba = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And FCBA = '" & rsHarianPupuk("FCBA") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where FCBA = '" & rsHarianPupuk("FCBA") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND FCBA = '" & rsHarianPupuk("FCBA") & "'"
						End IF 
					Else
						whereConsTtl =	whereConsTtl	& " And FCBA = '" & rsHarianPupuk("FCBA") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where FCBA = '" & rsHarianPupuk("FCBA") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND FCBA = '" & rsHarianPupuk("FCBA") & "'"
						End IF
					End if
				End If
				
				If Len(Trim(v_txtStockCode)) > 0 Then
					If v_txtStockCode = "--SUMMARY--" then
						whereConsTtl =	whereConsTtl	& " "
					Elseif v_txtStockCode = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " "
					Else
						whereConsTtl =	whereConsTtl	& " "
					End if
				End If
				
				If Len(Trim(v_txtAfdeling)) > 0 Then
					If v_txtAfdeling = "--SUMMARY--" then
						whereConsTtl 	=	whereConsTtl	& " "
						whereConsRow	= 	whereConsRow
					Elseif v_txtAfdeling = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And AFDELING = '" & rsHarianPupuk("AFDELING") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where AFDELING = '" & rsHarianPupuk("AFDELING") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND AFDELING = '" & rsHarianPupuk("AFDELING") & "'"
						End IF 
					Else
						whereConsTtl =	whereConsTtl	& " And AFDELING = '" & rsHarianPupuk("AFDELING") & "'"
						IF whereConsRow = " " then
							whereConsRow	= 	whereConsRow 	& " Where AFDELING = '" & rsHarianPupuk("AFDELING") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND AFDELING = '" & rsHarianPupuk("AFDELING") & "'"
						End IF
					End if
				End If
				
				If Len(Trim(v_txtBlok)) > 0 Then
					If v_txtBlok = "--SUMMARY--" then
						whereConsTtl =	whereConsTtl	& " "
					Elseif v_txtBlok = "--ALL--" then
						whereConsTtl =	whereConsTtl	& " And BLOK = '" & rsHarianPupuk("BLOK") & "'"
						IF whereConsRow = " "then
							whereConsRow	= 	whereConsRow 	& " Where BLOK = '" & rsHarianPupuk("BLOK") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND BLOK = '" & rsHarianPupuk("BLOK") & "'"
						End IF
					Else
						whereConsTtl =	whereConsTtl	& " And BLOK = '" & rsHarianPupuk("BLOK") & "'"
						IF whereConsRow = " "then
							whereConsRow	= 	whereConsRow 	& " Where BLOK = '" & rsHarianPupuk("BLOK") & "'"
						Else
							whereConsRow	= 	whereConsRow 	& " AND BLOK = '" & rsHarianPupuk("BLOK") & "'"
						End IF
					End if
				End If
				
		strRowsCount	= "Select Count(*)RowsCount FROM(" & strHarianPupuk & ") "& whereConsRow
		'response.write strRowsCount & "</br>-</br>"
		Set rsSqlRow = Server.CreateObject("ADODB.Recordset")
		rsSqlRow.Open strRowsCount ,connstr, adOpenStatic	
%>
		<tr>
				<% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <>  "" then %>
					<%
					   tempStockCode = rsHarianPupuk("stockname")
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
					      tempStockCode = rsHarianPupuk("stockname")
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
					      tempStockCode = rsHarianPupuk("stockname")
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
					      tempStockCode = rsHarianPupuk("stockname")
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
					      tempStockCode = rsHarianPupuk("stockname")
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
					      tempStockCode = rsHarianPupuk("stockname")
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
					      tempStockCode = rsHarianPupuk("stockname")
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

				<% if v_txtOwnership <>  "--SUMMARY--" and v_txtOwnership <>  "" then %>
				    <% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("stockname")
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

					   tempownership = rsHarianPupuk("ownership")
					   tempTahunTanam = "--ALL--"
					   tempLocationCode = "--SUMMARY--"
					 %>
					
					<% if tempownership <> sprevownership then %>
						<td align="center">
							<%=rsHarianPupuk("ownership")%>
						</td>
						<% sprevownership = tempownership %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtAfdeling <>  "--SUMMARY--" and v_txtAfdeling <>  "" then %>
				    <% if v_txtStockCode <>  "--SUMMARY--"  and  v_txtStockCode <> "" then
					      tempStockCode = rsHarianPupuk("stockname")
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
					      tempStockCode = rsHarianPupuk("stockname")
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
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_1"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_2"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_3"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_4"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_5"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_6"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_7"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_8"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_9"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_10"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_11"), 2)%></td>
					<td align="right"><%=FormatNumber(rsHarianPupuk("MONTH_12"), 2)%></td>
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
					<td align="right">-</td>
		</tr>
<%
			if NOX = CDbl(rsSqlRow("RowsCount"))then
				NOX = 1
				strHKEx	= "select GroupPlantation " & selectCons & strHK &"  From (" & strMainHK & " )" & _
						whereConsTtl & _
						" group by GroupPlantation" & selectCons & ", JOBCODE, JOBNAME, ACTIVEMATERIAL" & _
						" having (SUM(TOTAL) > 0)" & _
						" Order by GroupPlantation " & orderCons & ", ACTIVEMATERIAL"
						
            ' response.write strHKEx
			
			Set rsHKEx = Server.CreateObject("ADODB.Recordset")
			rsHKEx.Open strHKEx ,connstr, adOpenStatic
			'response.write strHKEx & "</br>-</br>"
				if rsHKEx.eof and rsHKEx.bof then
					MONTH_1 = 0
					MONTH_2 = 0
					MONTH_3 = 0
					MONTH_4 = 0
					MONTH_5 = 0
					MONTH_6 = 0
					MONTH_7 = 0
					MONTH_8 = 0
					MONTH_9 = 0
					MONTH_10 = 0
					MONTH_11 = 0
					MONTH_12 = 0
					TOTAL	= 0
				Else
					MONTH_1 = rsHKEx("MONTH_1")
					MONTH_2 = rsHKEx("MONTH_2")
					MONTH_3 = rsHKEx("MONTH_3")
					MONTH_4 = rsHKEx("MONTH_4")
					MONTH_5 = rsHKEx("MONTH_5")
					MONTH_6 = rsHKEx("MONTH_6")
					MONTH_7 = rsHKEx("MONTH_7")
					MONTH_8 = rsHKEx("MONTH_8")
					MONTH_9 = rsHKEx("MONTH_9")
					MONTH_10 = rsHKEx("MONTH_10")
					MONTH_11 = rsHKEx("MONTH_11")
					MONTH_12 = rsHKEx("MONTH_12")
					TOTAL = rsHKEx("TOTAL")
				End IF
%>
	<tr bgcolor="#D1D0CE">
			<td align="Center" colspan="<%=var_colspan%>"><b>TOTAL HK</b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_1,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_2,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_3,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_4,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_5,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_6,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_7,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_8,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_9,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_10,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_11,2)%></b></td>
			<td align="Right" ><b><%=FormatNumber(MONTH_12,2)%></b></td>
			<td align="right">-</td>
			<td align="right">-</td>
			<td align="right">-</td>
			<td align="Right" ><b><%=FormatNumber(TOTAL,2)%></b></td>
	</tr>
<%
			rsHKEx.Close
			Set rsHKEx = Nothing
			Else
				NOX = NOX+ 1
			End IF			
			'Ttl_mm += Cdbl(rsHarianPupukTtl("MONTH_1"))
			GrandTotalMonth_1 =  GrandTotalMonth_1 + Cdbl(rsHarianPupuk("MONTH_1"))
			GrandTotalMonth_2 =  GrandTotalMonth_2 + Cdbl(rsHarianPupuk("MONTH_2"))
			GrandTotalMonth_3 =  GrandTotalMonth_3 + Cdbl(rsHarianPupuk("MONTH_3"))
			GrandTotalMonth_4 =  GrandTotalMonth_4 + Cdbl(rsHarianPupuk("MONTH_4"))
			GrandTotalMonth_5 =  GrandTotalMonth_5 + Cdbl(rsHarianPupuk("MONTH_5"))
			GrandTotalMonth_6 =  GrandTotalMonth_6 + Cdbl(rsHarianPupuk("MONTH_6"))
			GrandTotalMonth_7 =  GrandTotalMonth_7 + Cdbl(rsHarianPupuk("MONTH_7"))
			GrandTotalMonth_8 =  GrandTotalMonth_8 + Cdbl(rsHarianPupuk("MONTH_8"))
			GrandTotalMonth_9 =  GrandTotalMonth_9 + Cdbl(rsHarianPupuk("MONTH_9"))
			GrandTotalMonth_10 =  GrandTotalMonth_10 + Cdbl(rsHarianPupuk("MONTH_10"))
			GrandTotalMonth_11 =  GrandTotalMonth_11 + Cdbl(rsHarianPupuk("MONTH_11"))
			GrandTotalMonth_12 =  GrandTotalMonth_12 + Cdbl(rsHarianPupuk("MONTH_12"))
			GrandRowRealTotal = GrandRowRealTotal + RowRealTotal
			GrandRowRekomTotal = GrandRowRekomTotal + RowRekomTotal
			
		
			rsHarianPupuk.MoveNext
			
			RowRealTotal	=	0
			RowRekomTotal	=	0
			RowVarTotal		=	0
			Wend	
rsHarianPupuk.Close
if cdbl(GrandRowRekomTotal) > 0 then
   GrandRowVarTotal = cdbl(GrandRowRealTotal)/cdbl(GrandRowRekomTotal)*100
else
   GrandRowVarTotal = 100
end if 
Set rsHarianPupuk = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td colspan="<%=var_colspan%>" align="center"><strong>TOTAL PUPUK</strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_1, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_2, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_3, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_4, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_5, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_6, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_7, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_8, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_9, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_10, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_11, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotalMonth_12, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandRowRealTotal, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandRowRekomTotal, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandRowVarTotal, 2)%></strong></td>
		<td align="right">-</td>
	</tr>
<%End if%>
	</tbody>
</table>
</body>
</html>
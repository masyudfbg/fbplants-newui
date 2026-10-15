<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtTanggal, v_txtUraian
Dim ScriptName
Dim rsDKWH, strDKWH, strDKWH2, strMain
Dim NO, cClass, nday, nSpanTotal

nSpanTotal = 6

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, OrderbyCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where company is not null "


If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--ALL--" then
		whereCons = wherecons & " "
	Elseif v_txtGroupPlantation = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
	end if
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--ALL--" then
		selectCons = selectCons & " , Plantation"
		whereCons = wherecons & " "
	Elseif v_txtPlantation = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , Plantation"
		whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
	end if
End If

If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--ALL--" then
		selectCons = selectCons & " , groupregion"
		whereCons = wherecons & " "
	Elseif v_txtGroupRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , groupregion"
		whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
	end if
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--ALL--" then
		selectCons = selectCons & " , region"
		whereCons = wherecons & " "
	Elseif v_txtRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , region"
		whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
	end if
End If

If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--ALL--" then
		selectCons = selectCons & " , company"
		whereCons = wherecons & " "
	Elseif v_txtCompany = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , company"
		whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
		q = " Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit" & _
			" where FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		   CompanyName=rsCat("FCCOMPANYNAME")			   
		end if
	end if
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--ALL--" then
		selectCons = selectCons & " , NoUrut, FCBA"
		whereCons = wherecons & " "
	Elseif v_txtfcba = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , NoUrut, FCBA"
		whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
		dim q, FcbaName, CompanyName, MgrName, KTUName, CropName, FcbaType
		q = "Select fccode, fccompanycode, FCNAME, FCCOMPANYNAME, FCMANAGER, FCKTU, FCCROP, FCTYPE from BusinessUnit where fcCode = '" & v_txtfcba & "'"
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then	
		   fcbaname		= rsCat("fcName")
		   CompanyName	= rsCat("FCCOMPANYNAME")
		   MgrName		= rsCat("FCMANAGER")
		   KTUName		= rsCat("FCKTU")
		   CropName		= rsCat("FCCROP")
		   FcbaType		= rsCat("fcType")
		end if
	end if
End If

v_txtTanggal	=	Request("txtTanggal")
v_txtUraian		=	Request("txtUraian")


If (v_txtTanggal = "" Or Len(Trim(v_txtTanggal)) = 0) Then
	v_txtTanggal = textDate(Date)
End If

If v_txtUraian = "" Then
	v_txtUraian	=	"--ALL--"
End If

If Len(Trim(v_txtUraian)) > 0 Then
	if v_txtUraian = "--ALL--" then
		whereCons = wherecons & " "
	Elseif v_txtUraian = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		whereCons = whereCons & " And  URAIAN = '" & v_txtUraian & "' "
	end if
End If

Dim v_Tahun : v_Tahun = Year(v_txtTanggal)
Dim v_Bulan : v_Bulan = Month(v_txtTanggal)

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT " & strFieldCons & "," & _
			"          SEQ," & _
			"          URAIAN," & _
			"          SUM (HI) HI," & _
			"          SUM (SHI) SHI," & _
			"          SUM (SBI) SBI" & _
			"     FROM (  SELECT FCBA," & _
			"                    SEQ," & _
			"                    URAIAN," & _
			"                    SUM (JAM) HI," & _
			"                    0 SHI," & _
			"                    0 SBI" & _
			"               FROM (SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            'TBM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            'TM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'VH' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MN'" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TBM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'VH' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'VH'" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE NOT IN ('MN', 'FF', 'VH'))" & _
			"              WHERE TANGGAL = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
			"           GROUP BY FCBA, SEQ, URAIAN" & _
			"           UNION ALL" & _
			"             SELECT FCBA," & _
			"                    SEQ," & _
			"                    URAIAN," & _
			"                    0 HI," & _
			"                    SUM (JAM) SHI," & _
			"                    0 SBI" & _
			"               FROM (SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            'TBM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            'TM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'VH' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MN'" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TBM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'VH' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'VH'" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE NOT IN ('MN', 'FF', 'VH'))" & _
			"              WHERE TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY')) = " & v_Tahun & "" & _
			"                    AND TO_NUMBER (TO_CHAR (TANGGAL, 'MM')) = " & v_Bulan & "" & _
			"                    AND TANGGAL <= TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
			"           GROUP BY FCBA, SEQ, URAIAN" & _
			"           UNION ALL" & _
			"             SELECT FCBA," & _
			"                    SEQ," & _
			"                    URAIAN," & _
			"                    0 HI," & _
			"                    0 SHI," & _
			"                    SUM (JAM) SBI" & _
			"               FROM (SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            'TBM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            'TM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'VH' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MN'" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TBM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'VH' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'VH'" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE NOT IN ('MN', 'FF', 'VH'))" & _
			"              WHERE TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY')) = " & v_Tahun & "" & _
			"                    AND TO_NUMBER (TO_CHAR (TANGGAL, 'MM')) <= " & v_Bulan & "" & _
			"           GROUP BY FCBA, SEQ, URAIAN)" & _
			" GROUP BY FCBA, SEQ, URAIAN" & _
			" ORDER BY FCBA, SEQ, URAIAN"

			
strDKWH2	=	"        , SEQ" & _
				"        , URAIAN" & _
				"        , SUM (HI) HI" & _
				"        , SUM (SHI) SHI" & _
				"        , SUM (SBI) SBI"
			
			
strDKWH	=	" select GroupPlantation " & selectCons & strDKWH2 & "  From (" & strMain & " )" & whereCons & _
						" group by GroupPlantation" & selectCons & " , SEQ, URAIAN Order by GroupPlantation" & selectCons &" , SEQ"
						
'response.write strDKWH
%>
<html>

<head>
<title>ePlaS Reporting - Distribusi Kerja Workshop Harian</title>
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
						<td><font face="verdana" size="1">TANGGAL</font></td>
						<td align="center">
						<% if v_txtTanggal <> "" and v_txtTanggal = v_txtTanggal then %>
							<input class="datepicker" type="text" maxlength="10" name="txtTanggal" size="10" value="<%=v_txtTanggal%>" title="Type Date mm/dd/yyyy" >
						<% else %>
							<input class="datepicker" type="text" maxlength="10" name="txtTanggal" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a href="javascript:show_calendar('frmEntry.txtTanggal');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						<td><font face="verdana" size="1">URAIAN</font></td>
						<td>
							<select name="txtUraian" >
								<option value="--ALL--"<%If v_txtUraian = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<option value="BIBITAN"<%If v_txtUraian = "BIBITAN" then%> Selected <%End If%>>BIBITAN</option>
								<option value="TBM"<%If v_txtUraian = "TBM" then%> Selected <%End If%>>TBM</option>
								<option value="TM"<%If v_txtUraian = "TM" then%> Selected <%End If%>>TM</option>
								<option value="VH"<%If v_txtUraian = "VH" then%> Selected <%End If%>>VH</option>
								<option value="LAIN-LAIN"<%If v_txtUraian = "LAIN-LAIN" then%> Selected <%End If%>>LAIN-LAIN</option>
							</select>
						</td>
						<td><input type="submit" value="View" /></td>
					</tr>				
				</table>
			</div>
			</form>
 <table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
            <th colspan="13" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			</div>
			<div id="subheader" class="centered">
				<h2>DISTRIBUSI KERJA WORKSHOP HARIAN</h2>
				<h5>TANGGAL: <%=v_txtTanggal%></h5>
			</div>
		</div>
        </th>
    </tr>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">URAIAN</th>
		<% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
		%>
	    <%end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;">REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;">SITE</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<th colspan="6" style="background-color:#00A300;color:#fff;">JAM KERJA</th>
	</tr>
	<tr>
		<th colspan="2" style="background-color:#00A300;color:#fff;">HARI INI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">JAM</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">JAM</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">JAM</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
    </thead>
<tbody>

<%
Set rsDKWH = Server.CreateObject("ADODB.Recordset")
rsDKWH.Open strDKWH, connstr, adOpenStatic

Dim rsGetTotal, strGetTotal
strGetTotal = " Select Sum(HI) HI, Sum(SHI) SHI, Sum(SBI) SBI FROM (" & strDKWH & ")"
Set rsGetTotal = Server.CreateObject("ADODB.Recordset")
rsGetTotal.Open strGetTotal, connstr, adOpenStatic

Dim Temp_Total_HI : Temp_Total_HI = 0
Dim Temp_Total_SHI : Temp_Total_SHI = 0
Dim Temp_Total_SBI : Temp_Total_SBI = 0

Dim	GrandTotal_HI : GrandTotal_HI = 0
Dim GrandTotal_SHI : GrandTotal_SHI = 0
Dim GrandTotal_SBI : GrandTotal_SBI = 0

Dim Total_Persen_HI : Total_Persen_HI = 0
Dim Total_Persen_SHI : Total_Persen_SHI = 0
Dim Total_Persen_SBI : Total_Persen_SBI = 0

Dim GrandTotal_Persen_HI : GrandTotal_Persen_HI = 0
Dim GrandTotal_Persen_SHI : GrandTotal_Persen_SHI = 0
Dim GrandTotal_Persen_SBI : GrandTotal_Persen_SBI = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempUraian : tempUraian = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim SprevUraian : SprevUraian = ""

If Not rsGetTotal.EOF Then
	Temp_Total_HI = Temp_Total_HI + Cdbl(rsGetTotal("HI"))
	Temp_Total_SHI = Temp_Total_SHI + Cdbl(rsGetTotal("SHI"))
	Temp_Total_SBI = Temp_Total_SBI + Cdbl(rsGetTotal("SBI"))
Else
	Temp_Total_HI = 0
	Temp_Total_SHI = 0
	Temp_Total_SBI = 0
End If
rsGetTotal.Close
set rsGetTotal = Nothing

if (rsDKWH.bof and rsDKWH.eof) then
%>
	<tr>
		<td colspan="<%=1+nSpanTotal+8%>">NO DATA AVAILABLE</td>
	</tr>
<%else
		while not rsDKWH.eof
			Total_Persen_HI = 0
			Total_Persen_SHI = 0
			Total_Persen_SBI = 0
			
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<% if v_txtUraian <>  "--SUMMARY--" and  v_txtUraian <>  "" then %>
						<% tempUraian = rsDKWH("URAIAN")
						   tempGroupPlantation = "--ALL--"
						   tempPlantation = "--SUMMARY--"
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempUraian <> SprevUraian then %>
							<td align="left">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
								<%=rsDKWH("URAIAN")%>
								</a>
							</td>
							<% SprevUraian = tempUraian %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>

				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   tempGroupPlantation = rsDKWH("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsDKWH("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsDKWH("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDKWH("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsDKWH("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDKWH("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDKWH("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsDKWH("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDKWH("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDKWH("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsDKWH("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsDKWH("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
						    <%=rsDKWH("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
<%
			If Temp_Total_HI = 0 Then
				Total_Persen_HI	= 0
				GrandTotal_Persen_HI = 0
			Else
				Total_Persen_HI	= Round((Cdbl(rsDKWH("HI"))/Temp_Total_HI)*100, 2)
				GrandTotal_Persen_HI	=	GrandTotal_Persen_HI + (Cdbl(rsDKWH("HI"))/Temp_Total_HI)*100
			End If
			If Temp_Total_SHI = 0 Then
				Total_Persen_SHI = 0
				GrandTotal_Persen_SHI = 0
			Else
				Total_Persen_SHI = Round((Cdbl(rsDKWH("SHI"))/Temp_Total_SHI)*100, 2)
				GrandTotal_Persen_SHI	=	GrandTotal_Persen_SHI + (Cdbl(rsDKWH("SHI"))/Temp_Total_SHI)*100
			End If
			If Temp_Total_SBI = 0 Then
				Total_Persen_SBI = 0
				GrandTotal_Persen_SBI = 0
			Else
				Total_Persen_SBI = Round((Cdbl(rsDKWH("SBI"))/Temp_Total_SBI)*100, 2)
				GrandTotal_Persen_SBI	=	GrandTotal_Persen_SBI + (Cdbl(rsDKWH("SBI"))/Temp_Total_SBI)*100
			End If
%>
				
				<td align="right"><%=FormatNumber(rsDKWH("HI"), 2)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_HI, 2)%></td>

				<td align="right"><%=FormatNumber(rsDKWH("SHI"), 2)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SHI, 2)%></td>

				<td align="right"><%=FormatNumber(rsDKWH("SBI"), 2)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SBI, 2)%></td>
			</tr>
<%
			GrandTotal_HI	=	GrandTotal_HI + Cdbl(rsDKWH("HI"))
			GrandTotal_SHI	=	GrandTotal_SHI + Cdbl(rsDKWH("SHI"))
			GrandTotal_SBI	=	GrandTotal_HI + Cdbl(rsDKWH("SBI"))
			rsDKWH.MoveNext
		Wend
	rsDKWH.Close
	set rsDKWH = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td align="center" colspan="<%=1+nSpanTotal%>"><strong>TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_HI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_HI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_SHI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_SHI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_SBI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_SBI, 2)%></strong></td>
	</tr>
<%End if%>
</tbody>
</table>
		</div>
	</body>
</html>
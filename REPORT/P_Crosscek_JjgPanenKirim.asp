<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtDivisi, v_txtLocationCode
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsKJ, strKJ
Dim NO, cClass, nday

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtGangCode	=	Request("txtGangCode")
v_txtDivisi		=	Request("txtDivisi")
v_txtLocationCode		=	Request("txtLocationCode")


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if

strEmployee	=	" Select FCBA, DIVISION, BLOK, JANJANG, " & _
				" Sum(day_1) as day_1," & _
				" Sum(day_2) as day_2," & _
				" Sum(day_3) as day_3," & _
				" Sum(day_4) as day_4," & _
				" Sum(day_5) as day_5," & _
				" Sum(day_6) as day_6," & _
				" Sum(day_7) as day_7," & _
				" Sum(day_8) as day_8," & _
				" Sum(day_9) as day_9," & _
				" Sum(day_10) as day_10," & _
				" Sum(day_11) as day_11," & _
				" Sum(day_12) as day_12," & _
				" Sum(day_13) as day_13," & _
				" Sum(day_14) as day_14," & _
				" Sum(day_15) as day_15," & _
				" Sum(day_16) as day_16," & _
				" Sum(day_17) as day_17," & _
				" Sum(day_18) as day_18," & _
				" Sum(day_19) as day_19," & _
				" Sum(day_20) as day_20," & _
				" Sum(day_21) as day_21," & _
				" Sum(day_22) as day_22," & _
				" Sum(day_23) as day_23," & _
				" Sum(day_24) as day_24," & _
				" Sum(day_25) as day_25," & _
				" Sum(day_26) as day_26," & _
				" Sum(day_27) as day_27," & _
				" Sum(day_28) as day_28," & _
				" Sum(day_29) as day_29," & _
				" Sum(day_30) as day_30," & _
				" Sum(day_31) as day_31," & _
				" Sum(total) as total" & _
				" from (" & _
				" Select h.FCBA,  to_number(to_char(h.receptiondate,'yyyy')) as Tahun, to_number(to_char(h.receptiondate,'mm')) as Bulan, f.division, h.FIELDCODE as BLOK,'2. KIRIM' as Janjang," & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),1, bunch,NULL))as day_1, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),2, bunch,NULL))as day_2, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),3, bunch,NULL))as day_3, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),4, bunch,NULL))as day_4, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),5, bunch,NULL))as day_5, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),6, bunch,NULL))as day_6, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),7, bunch,NULL))as day_7, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),8, bunch,NULL))as day_8, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),9, bunch,NULL))as day_9, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),10,bunch,NULL))as day_10, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),11,bunch,NULL))as day_11, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),12,bunch,NULL))as day_12, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),13,bunch,NULL))as day_13, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),14,bunch,NULL))as day_14, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),15,bunch,NULL))as day_15, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),16,bunch,NULL))as day_16, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),17,bunch,NULL))as day_17, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),18,bunch,NULL))as day_18, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),19,bunch,NULL))as day_19, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),20,bunch,NULL))as day_20, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),21,bunch,NULL))as day_21, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),22,bunch,NULL))as day_22, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),23,bunch,NULL))as day_23, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),24,bunch,NULL))as day_24, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),25,bunch,NULL))as day_25, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),26,bunch,NULL))as day_26, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),27,bunch,NULL))as day_27, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),28,bunch,NULL))as day_28, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),29,bunch,NULL))as day_29, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),30,bunch,NULL))as day_30, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),31,bunch,NULL))as day_31," & _
				" Sum ( h.bunch) as Total" & _
				" From harvestingspb h, field f where (h.FCBA=f.FCBA and h.fieldcode=f.fccode)" & _
				" and h.FCBA = '" & Session ("FCBA") & "'" & _
				" Group by h.FCBA, to_number(to_char(h.receptiondate,'yyyy')), to_number(to_char(h.receptiondate,'mm')), f.division, h.FIELDCODE" & _
				" union all" & _
				" Select ag.FCBA, to_number(to_char(ag.fddate,'yyyy')) as Tahun, to_number(to_char(ag.fddate,'mm')) as Bulan, f.division, ag.LOCATIONCODE AS BLOK,'1. PANEN' as Janjang," & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),1, unit,NULL))as day_1, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),2, unit,NULL))as day_2, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),3, unit,NULL))as day_3, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),4, unit,NULL))as day_4, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),5, unit,NULL))as day_5, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),6, unit,NULL))as day_6, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),7, unit,NULL))as day_7, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),8, unit,NULL))as day_8, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),9, unit,NULL))as day_9, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),10,unit,NULL))as day_10, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),11,unit,NULL))as day_11, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),12,unit,NULL))as day_12, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),13,unit,NULL))as day_13, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),14,unit,NULL))as day_14, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),15,unit,NULL))as day_15, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),16,unit,NULL))as day_16, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),17,unit,NULL))as day_17, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),18,unit,NULL))as day_18, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),19,unit,NULL))as day_19, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),20,unit,NULL))as day_20, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),21,unit,NULL))as day_21, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),22,unit,NULL))as day_22, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),23,unit,NULL))as day_23, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),24,unit,NULL))as day_24, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),25,unit,NULL))as day_25, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),26,unit,NULL))as day_26, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),27,unit,NULL))as day_27, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),28,unit,NULL))as day_28, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),29,unit,NULL))as day_29, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),30,unit,NULL))as day_30, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),31,unit,NULL))as day_31," & _
				" Sum ( ag.UNIT ) as Total" & _
				" from attendance_gad ag, field f where (ag.FCBA=f.FCBA and ag.LOCATIONCODE=f.fccode)" & _
				" and jobcode in ('6502101','1323001') " & _
				" and ag.FCBA = '" & Session ("FCBA") & "'" & _
				" Group by ag.FCBA, to_number(to_char(ag.fddate,'yyyy')) ,to_number(to_char(ag.fddate,'mm')), f.division, ag.LOCATIONCODE " & _
				" union all" & _
				" Select hq.FCBA, to_number(to_char(hq.fddate,'yyyy')) as Tahun, to_number(to_char(hq.fddate,'mm')) as Bulan, f.division, hq.fieldcode AS BLOK,'3. HQ' as Janjang," & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),1, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_1, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),2, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_2, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),3, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_3, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),4, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_4, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),5, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_5, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),6, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_6, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),7, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_7, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),8, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_8, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),9, under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_9, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),10,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_10, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),11,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_11, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),12,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_12, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),13,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_13, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),14,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_14, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),15,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_15, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),16,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_16, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),17,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_17, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),18,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_18, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),19,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_19, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),20,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_20, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),21,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_21, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),22,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_22, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),23,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_23, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),24,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_24, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),25,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_25, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),26,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_26, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),27,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_27, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),28,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_28, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),29,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_29, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),30,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_30, " & _
				" Sum ( decode (to_number (to_char (hq.fddate,'dd')),31,under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond,NULL))as day_31," & _
				" Sum ( hq.under_ripe + over_ripe+ abnormal + long_stalk + eaten_by_rat + unharvest_ffb + uncollect_lf_circle + uncollect_lf_piece + unarrange_ffb + unprune_frond ) as Total" & _
				" from HarvestingQuality hq, field f where (hq.FCBA=f.FCBA and hq.fieldcode=f.fccode)" & _
				" and hq.FCBA = '" & Session ("FCBA") & "'" & _
				" Group by hq.FCBA, to_number(to_char(hq.fddate,'yyyy')) ,to_number(to_char(hq.fddate,'mm')), f.division, hq.fieldcode)" & _
				" WHERE FCBA = '" & Session ("FCBA") & "'"


if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " and Tahun = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " and Bulan =" & v_txtMonth
end if
if  LEN(TRIM(v_txtdivisi)) > 0 then
	strEmployee	=	strEmployee & " AND DIVISION = '" & v_txtdivisi & "'"
end if
if  LEN(TRIM(v_txtLocationCode)) > 0 then
	strEmployee	=	strEmployee & " AND BLOK = '" & v_txtLocationCode & "'"
end if
	strEmployee	=	strEmployee & " group by FCBA, DIVISION, BLOK, JANJANG ORDER BY BLOK, JANJANG ASC"

%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Cross Check Janjang Panen vs Kirim</title>
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
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
		<td><font face="verdana" size="1">DIVISI</font></td><td><select name="txtDivisi" ><option value=""></option>
				
		<%
		Dim rsDivisi, strDivisi
		strDivisi	=	" SELECT * FROM PARAMETER WHERE PARCODE = 'E0002' AND FCCODE LIKE 'AFD%' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
		
		Set rsDivisi = Server.CreateObject("ADODB.Recordset")
		rsDivisi.Open strDivisi, connstr, adOpenStatic 
		
		if not rsDivisi.eof then
		Do While Not rsDivisi.eof
		if rsDivisi("FCCODE") = v_txtDivisi then%>
		        <option value="<%=rsDivisi("FCCODE")%>" Selected><%=rsDivisi("FCCODE")%></option> 
		     <%Else%>
		        <option value="<%=rsDivisi("FCCODE")%>"><%=rsDivisi("FCCODE")%></option>  
		    <%End if
			rsDivisi.movenext
		loop
		end if
		rsDivisi.Close
		set rsDivisi = Nothing
		%>
		</select></td>
				<td><font face="verdana" size="1">BLOK</font></td><td><select name="txtLocationCode" ><option value=""></option>
				
		<%
		Dim rsLocationCode, strLocationCode
		strLocationCode	=	" SELECT FCCODE, FCNAME FROM FIELD WHERE  FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
		
		Set rsLocationCode = Server.CreateObject("ADODB.Recordset")
		rsLocationCode.Open strLocationCode, connstr, adOpenStatic 
		
		if not rsLocationCode.eof then
		Do While Not rsLocationCode.eof
		if rsLocationCode("FCCODE") = v_txtLocationCode then%>
		        <option value="<%=rsLocationCode("FCCODE")%>" Selected><%=rsLocationCode("FCCODE")%></option> 
		     <%Else%>
		        <option value="<%=rsLocationCode("FCCODE")%>"><%=rsLocationCode("FCCODE")%></option>  
		    <%End if
			rsLocationCode.movenext
		loop
		end if
		rsLocationCode.Close
		set rsLocationCode = Nothing
		%>
		</select></td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>

</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
		<h2>Cross Check Janjang Panen vs Kirim</h2>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	    </div>
	</div>
</th>
		</tr>
	<tr>
		<th>NO</th>
		<th>BLOK</th>
		<th>JANJANG</th>
<%for nday = 1 to 31%>
<th><%=nday%></th>
<%next%>
<th>TOTAL</th>
	</tr>

</thead>
<tbody>
<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
'response.write strEmployee
rsEmployee.Open strEmployee, connstr, adOpenStatic

if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="36">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1	
		while not rsEmployee.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="left"><%=rsEmployee("BLOK")%></td>
				<td align="left"><%=rsEmployee("JANJANG")%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_3"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_4"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_5"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_6"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_7"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_8"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_9"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_10"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_11"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_12"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_13"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_14"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_15"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_16"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_17"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_18"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_19"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_20"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_21"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_22"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_23"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_24"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_25"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_26"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_27"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_28"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_29"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_30"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_31"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("TOTAL"))%></td>
<%
			NO = NO + 1
		
			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>
</tbody>
</table>

   <table width="90%" class="data">
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>

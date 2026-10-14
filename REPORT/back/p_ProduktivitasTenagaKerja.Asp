<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, v_txtDate, v_txtJobCode
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
v_txtDate	=	trim(request("txtDate"))
v_txtJobCode=	trim(request("txtJobCode"))
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Produktivitas Tenaga Kerja</title>
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
@media print
{
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>

</head>

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<div class="printhide">
	<!--#include file="menutop.asp"-->
	<table id="form-table" class="centered-margin">
		<form name="frmEntry" method="post" action="<%=ScriptName%>">
			<tr>
				<td><font face="verdana" size="1">TAHUN</font></td>
				<td>
<% 
    MM=Request.Form("vMonth")
    yyyy=Request.Form("vYear")

    mon = Month(Session("v_Costbook"))
    If len(MM)=0 then MM=mon

    Yr = Year(Session("v_Costbook"))
    MinYr = Yr - 2
    MaxYr = Yr  + 2
    If len(yyyy)= 0 then yyyy = Yr
%>
				<select name="vYear">
<%
    For i = MinYr to MaxYr
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
<%
    Next
%>
				</select>
				</td>
				<td><font face="verdana" size="1">BULAN</font></td>
				<td>
					<select name="vMonth">
<%
    For i = 1 to 12
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=Bulan(i)%></option>
<%
    Next
%>
					</select>
				</td>
				<td><font face="verdana" size="1">TANGGAL</font></td>
				<td align="center">
				<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
					<input type="text" maxlength="10" name="txtDate" size="10" value="<%=v_txtDate%>" title="Type Date mm/dd/yyyy" >
				<% else %>
					<input type="text" maxlength="10" name="txtDate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
				<% end if%>
					<a href="javascript:show_calendar('frmEntry.txtDate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
						<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
					</a>
				</td>
				<td><font face="verdana" size="1">JOBCODE</font></td><td><select name="txtJobCode" ><option value=""></option>
				<%
				Dim rsJobCode, strJobCode
				strJobCode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
				
				Set rsJobCode = Server.CreateObject("ADODB.Recordset")
				rsJobCode.Open strJobCode, connstr, adOpenStatic 
				
				if not rsJobCode.eof then
				Do While Not rsJobCode.eof
				if rsJobCode("FCCODE") = v_txtJobCode then%>
						<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option> 
					 <%Else%>
						<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>  
					<%End if
					rsJobCode.movenext
				loop
				end if
				rsJobCode.Close
				set rsJobCode = Nothing
				%>
				</select></td>
				<td><input type="submit" value="View" /></td>
			</tr>
		</form>
	</table>
</div>
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h3>LAPORAN PRODUKTIVITAS TENAGA KERJA</h3>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4>
	</div>
	</div>
<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF" rowspan="3">KODE PEKERJAAN</th>
		<th style="color:#FFFFFF" rowspan="3">JENIS PEKERJAAN</th>
		<th style="color:#FFFFFF" colspan="7">HARI INI</th>
		<th style="color:#FFFFFF" colspan="7">SAMPAI DENGAN HARI INI</th>
		<th style="color:#FFFFFF" colspan="7">SAMPAI DENGAN BULAN INI</th>
		<th style="color:#FFFFFF" rowspan="2">BUDGET</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" colspan="3">TENAGA KERJA (HK)</th>
		<th style="color:#FFFFFF" colspan="3">HASIL KERJA (HA)</th>
		<th style="color:#FFFFFF">OUTPUT</th>
		<th style="color:#FFFFFF" colspan="3">TENAGA KERJA (HK)</th>
		<th style="color:#FFFFFF" colspan="3">HASIL KERJA (HA)</th>
		<th style="color:#FFFFFF">OUTPUT</th>
		<th style="color:#FFFFFF" colspan="3">TENAGA KERJA (HK)</th>
		<th style="color:#FFFFFF" colspan="3">HASIL KERJA (HA)</th>
		<th style="color:#FFFFFF">OUTPUT</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF">SKU</th>
		<th style="color:#FFFFFF">KHL</th>
		<th style="color:#FFFFFF">TOTAL</th>
		<th style="color:#FFFFFF">HARIAN</th>
		<th style="color:#FFFFFF">BORONGAN</th>
		<th style="color:#FFFFFF">TOTAL</th>
		<th style="color:#FFFFFF">HK/HA/UNIT</th>
		<th style="color:#FFFFFF">SKU</th>
		<th style="color:#FFFFFF">KHL</th>
		<th style="color:#FFFFFF">TOTAL</th>
		<th style="color:#FFFFFF">HARIAN</th>
		<th style="color:#FFFFFF">BORONGAN</th>
		<th style="color:#FFFFFF">TOTAL</th>
		<th style="color:#FFFFFF">HK/HA/UNIT</th>
		<th style="color:#FFFFFF">SKU</th>
		<th style="color:#FFFFFF">KHL</th>
		<th style="color:#FFFFFF">TOTAL</th>
		<th style="color:#FFFFFF">HARIAN</th>
		<th style="color:#FFFFFF">BORONGAN</th>
		<th style="color:#FFFFFF">TOTAL</th>
		<th style="color:#FFFFFF">HK/HA/UNIT</th>
		<th style="color:#FFFFFF">HK/HA/UNIT</th>
	</tr>
<%
Dim rsProduktivitas, strProduktivitas, strGroupBy, rowno, rsalokasiBTL, stralokasiBTL
rowno		=	0
strProduktivitas		=	" SELECT fcba," & _
							" jobcode," & _
							" jobname," & _
							" SUM (sku) sku," & _
							" SUM (khl) khl," & _
							" SUM (sku) + SUM (khl) total," & _
							" SUM (unit_daily) unit_daily," & _
							" SUM (unit_contract) unit_contract," & _
							" SUM (unit_daily) + SUM (unit_contract) totalha," & _
							" (SUM (sku) + SUM (khl))" & _
							" / DECODE (SUM (unit_daily), 0, 1, SUM (unit_daily))" & _
							" output_daily," & _
							" SUM (sku_1) sku_1," & _
							" SUM (khl_1) khl_1," & _
							" SUM (sku_1) + SUM (khl_1) total_1," & _
							" SUM (unit_daily_1) unit_daily_1," & _
							" SUM (unit_contract_1) unit_contract_1," & _
							" SUM (unit_daily_1) + SUM (unit_contract_1) totalha_1," & _
							" (SUM (sku_1) + SUM (khl_1))" & _
							" / DECODE (SUM (unit_daily_1), 0, 1, SUM (unit_daily_1))" & _
							" output_daily_1," & _
							" SUM (sku_2) sku_2," & _
							" SUM (khl_2) khl_2," & _
							" SUM (sku_2) + SUM (khl_2) total_2," & _
							" SUM (unit_daily_2) unit_daily_2," & _
							" SUM (unit_contract_2) unit_contract_2," & _
							" SUM (output_budget) output_budget, " & _
							" SUM (unit_daily_2) + SUM (unit_contract_2) totalha_2," & _
							" (SUM (sku_2) + SUM (khl_2))" & _
							" / DECODE (SUM (unit_daily_2), 0, 1, SUM (unit_daily_2))" & _
							" output_daily_2" & _
							" FROM ( SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" SUM (NVL (mandays, 0)) sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 Unit_Daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2," & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND c.presentgrade IN ('KHT', 'KBT')" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') = a.fddate" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" SUM (NVL (mandays, 0)) khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 Unit_Daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2, " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND c.presentgrade IN ('KHL')" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') = a.fddate" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" SUM (NVL (unit, 0)) unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 Unit_Daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2 ," & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') = a.fddate" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 Unit_Daily," & _
							" SUM (NVL (unit, 0)) unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 Unit_Daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2, " & _
							" 0 output_budget " & _														
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND a.attendance = 'KB'" & _
							" AND TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') = a.fddate" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" SUM (NVL (mandays, 0)) sku_1," & _
							" 0 khl_1," & _
							" 0 unit_daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2 , " & _
							" 0 output_budget " & _							
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND c.presentgrade IN ('KHT', 'KBT')" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'MM')) = '" & MM & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" SUM (NVL (mandays, 0)) khl_1," & _
							" 0 unit_daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2, " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND c.presentgrade IN ('KHL')" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'MM')) = '" & MM & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" SUM (NVL (unit, 0)) unit_daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2, " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'MM')) = '" & MM & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 Unit_Daily_1," & _
							" SUM (NVL (unit, 0)) unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" 0 unit_contract_2 , " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND a.attendance = 'KB'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'MM')) = '" & MM & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 unit_daily_1," & _
							" 0 unit_contract_1," & _
							" SUM (NVL (mandays, 0)) sku_2," & _
							" 0 khl_2," & _
							" 0 unit_daily_2," & _
							" 0 unit_contract_2 , " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND c.presentgrade IN ('KHT', 'KBT')" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 unit_daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" SUM (NVL (mandays, 0)) khl_2," & _
							" 0 unit_daily_2," & _
							" 0 unit_contract_2, " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND c.presentgrade IN ('KHL')" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION" & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 unit_daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" SUM (NVL (unit, 0)) unit_daily_2," & _
							" 0 unit_contract_2 , " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND a.attendance = 'KJ'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname" & _
							" UNION " & _
							"	SELECT a.fcba, " & _
							"		   a.jobcode, " & _
							"		   c.fcname jobname, " & _
							"		   0 sku, " & _
							"		   0 khl, " & _
							"		   0 unit_daily, " & _
							"		   0 unit_contract, " & _
							"		   0 sku_1, " & _
							"		   0 khl_1, " & _
							"		   0 unit_daily_1, " & _
							"		   0 unit_contract_1, " & _
							"		   0 sku_2, " & _
							"		   0 khl_2, " & _
							"		   0 Unit_Daily_2, " & _
							"		   0 unit_contract_2, " & _
							"		   a.budget / b.total_count output_budget " & _
							"	  FROM (  SELECT fcba, " & _
							"					 jobcode, " & _
							"					 SUM ( " & _
							"						 (	OUTPUT_01 " & _
							"						  + OUTPUT_02 " & _
							"						  + OUTPUT_03 " & _
							"						  + OUTPUT_04 " & _
							"						  + OUTPUT_05 " & _
							"						  + OUTPUT_06 " & _
							"						  + OUTPUT_07 " & _
							"						  + OUTPUT_08 " & _
							"						  + OUTPUT_09 " & _
							"						  + OUTPUT_10 " & _
							"						  + OUTPUT_11 " & _
							"						  + OUTPUT_12) " & _
							"						 / 12) " & _
							"						 budget " & _
							"				FROM budget_transaction_estate " & _
							"			   WHERE fcba = '" & Session("FCBA") & "' AND budget_year = '" & yyyy & "' AND sourcetype = 'EM' " & _
							"			GROUP BY fcba, jobcode) a, " & _
							"		   (  SELECT fcba, jobcode, SUM (total_count) total_count " & _
							"				FROM (SELECT fcba, " & _
							"							 jobcode, " & _
							"							 1 total_count, " & _
							"							 (	BUDGET_01 " & _
							"							  + BUDGET_02 " & _
							"							  + BUDGET_03 " & _
							"							  + BUDGET_04 " & _
							"							  + BUDGET_05 " & _
							"							  + BUDGET_06 " & _
							"							  + BUDGET_07 " & _
							"							  + BUDGET_08 " & _
							"							  + BUDGET_09 " & _
							"							  + BUDGET_10 " & _
							"							  + BUDGET_11 " & _
							"							  + BUDGET_12) " & _
							"								 total_budget " & _
							"						FROM budget_transaction_estate " & _
							"					   WHERE	 fcba = '" & Session("FCBA") & "'" & _
							"							 AND budget_year = '" & yyyy & "' " & _
							"							 AND sourcetype = 'EM') " & _
							"			   WHERE total_budget > 0 " & _
							"			GROUP BY fcba, jobcode) b, " & _
							"			JOB c " & _
							"	 WHERE " & _ 
							"	 a.jobcode = b.jobcode AND " & _
							"	 a.jobcode = c.fccode AND " & _
							"	 a.fcba = b.fcba AND  " & _
							"	 a.fcba = c.fcba " & _
							" UNION " & _
							" SELECT a.fcba," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" 0 sku," & _
							" 0 khl," & _
							" 0 unit_daily," & _
							" 0 unit_contract," & _
							" 0 sku_1," & _
							" 0 khl_1," & _
							" 0 Unit_Daily_1," & _
							" 0 unit_contract_1," & _
							" 0 sku_2," & _
							" 0 khl_2," & _
							" 0 Unit_Daily_2," & _
							" SUM (NVL (unit, 0)) unit_contract_2, " & _
							" 0 output_budget " & _
							" FROM attendance_gad a, job b, employee c" & _
							" WHERE A.JOBCODE = B.FCCODE" & _
							" AND A.EMPLOYEECODE = C.FCCODE" & _
							" AND a.fcba = '" & Session("FCBA") & "'" & _
							" AND a.fcba = b.fcba" & _
							" AND a.fcba = c.fcba" & _
							" AND a.attendance = 'KB'" & _
							" AND TO_NUMBER(TO_CHAR (a.fddate, 'YYYY')) = '" & yyyy & "'" & _
							" AND a.fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
							" GROUP BY a.fcba, a.jobcode, b.fcname)"
if  LEN(TRIM(v_txtJobCode)) > 0 then
	strProduktivitas	=	strProduktivitas & " WHERE JOBCODE = '" & v_txtjobCode & "'"
end if
	strProduktivitas	=	strProduktivitas & " GROUP BY fcba, jobcode, jobname"


Set rsProduktivitas = Server.CreateObject("ADODB.Recordset")
'response.write strProduktivitas
rsProduktivitas.Open strProduktivitas, connstr, adOpenStatic

Dim nSKU, nKHL, nTOTAL, nUNIT_DAILY, nUNIT_CONTRACT, nTOTALHA, nOUTPUT_DAILY
Dim nSKU_1, nKHL_1, nTOTAL_1, nUNIT_DAILY_1, nUNIT_CONTRACT_1, nTOTALHA_1, nOUTPUT_DAILY_1
Dim nSKU_2, nKHL_2, nTOTAL_2, nUNIT_DAILY_2, nUNIT_CONTRACT_2, nTOTALHA_2, nOUTPUT_DAILY_2
Dim nOUTPUT_BUDGET

if Not rsProduktivitas.eof then

	nSKU = 0
	nKHL = 0
	nTOTAL = 0
	nUNIT_DAILY = 0
	nUNIT_CONTRACT = 0
	nTOTALHA = 0
	nOUTPUT_DAILY = 0
	nSKU_1 = 0
	nKHL_1 = 0
	nTOTAL_1 = 0
	nUNIT_DAILY_1 = 0
	nUNIT_CONTRACT_1 = 0
	nTOTALHA_1 = 0
	nOUTPUT_DAILY_1 = 0
	nSKU_2 = 0
	nKHL_2 = 0
	nTOTAL_2 = 0
	nUNIT_DAILY_2 = 0
	nUNIT_CONTRACT_2 = 0
	nTOTALHA_2 = 0
	nOUTPUT_DAILY_2 = 0
	nOUTPUT_BUDGET = 0


	While Not rsProduktivitas.eof
			rowno	=	rowno	+	1
%>
	<tr>
		<td align="center"><%=rsProduktivitas("JOBCODE")%></td>
		<td align="left"><%=rsProduktivitas("JOBNAME")%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("SKU"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("KHL"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("TOTAL"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("UNIT_DAILY"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("UNIT_CONTRACT"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("TOTALHA"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("OUTPUT_DAILY"),2)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("SKU_1"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("KHL_1"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("TOTAL_1"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("UNIT_DAILY_1"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("UNIT_CONTRACT_1"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("TOTALHA_1"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("OUTPUT_DAILY_1"),2)%></td>		
		<td align="right"><%=FormatNumber(rsProduktivitas("SKU_2"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("KHL_2"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("TOTAL_2"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("UNIT_DAILY_2"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("UNIT_CONTRACT_2"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("TOTALHA_2"),0)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("OUTPUT_DAILY_2"),2)%></td>
		<td align="right"><%=FormatNumber(rsProduktivitas("OUTPUT_BUDGET"),2)%></td>
	</tr>
<%

		nSKU = nSKU + Cdbl(rsProduktivitas("SKU"))
		nKHL = nKHL + Cdbl(rsProduktivitas("KHL"))
		nTOTAL = nTOTAL + Cdbl(rsProduktivitas("TOTAL"))
		nUNIT_DAILY = nUNIT_DAILY + Cdbl(rsProduktivitas("UNIT_DAILY"))
		nUNIT_CONTRACT = nUNIT_CONTRACT + Cdbl(rsProduktivitas("UNIT_CONTRACT"))
		nTOTALHA = nTOTALHA + Cdbl(rsProduktivitas("TOTALHA"))
		nOUTPUT_DAILY = nOUTPUT_DAILY + Cdbl(rsProduktivitas("OUTPUT_DAILY"))
		nSKU_1 = nSKU_1 + Cdbl(rsProduktivitas("SKU_1"))
		nKHL_1 = nKHL_1 + Cdbl(rsProduktivitas("KHL_1"))
		nTOTAL_1 = nTOTAL_1 + Cdbl(rsProduktivitas("TOTAL_1"))
		nUNIT_DAILY_1 = nUNIT_DAILY_1 + Cdbl(rsProduktivitas("UNIT_DAILY_1"))
		nUNIT_CONTRACT_1 = nUNIT_CONTRACT_1 + Cdbl(rsProduktivitas("UNIT_CONTRACT_1"))
		nTOTALHA_1 = nTOTALHA_1 + Cdbl(rsProduktivitas("TOTALHA_1"))
		nOUTPUT_DAILY_1 = nOUTPUT_DAILY_1 + Cdbl(rsProduktivitas("OUTPUT_DAILY_1"))
		nSKU_2 = nSKU_2 + Cdbl(rsProduktivitas("SKU_2"))
		nKHL_2 = nKHL_2 + Cdbl(rsProduktivitas("KHL_2"))
		nTOTAL_2 = nTOTAL_2 + Cdbl(rsProduktivitas("TOTAL_2"))
		nUNIT_DAILY_2 = nUNIT_DAILY_2 + Cdbl(rsProduktivitas("UNIT_DAILY_2"))
		nUNIT_CONTRACT_2 = nUNIT_CONTRACT_2 + Cdbl(rsProduktivitas("UNIT_CONTRACT_2"))
		nTOTALHA_2 = nTOTALHA_2 + Cdbl(rsProduktivitas("TOTALHA_2"))
		nOUTPUT_DAILY_2 = nOUTPUT_DAILY_2 + Cdbl(rsProduktivitas("OUTPUT_DAILY_2"))
		nOUTPUT_BUDGET = nOUTPUT_BUDGET + Cdbl(rsProduktivitas("OUTPUT_BUDGET"))

		rsProduktivitas.MoveNext
	Wend
end if
%>

	<tr bgcolor="#F4F4A0">
		<td align="center" colspan="2"><strong>TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(nSKU,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nKHL,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nTOTAL,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nUNIT_DAILY,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nUNIT_CONTRACT,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nTOTALHA,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nOUTPUT_DAILY,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nSKU_1,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nKHL_1,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nTOTAL_1,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nUNIT_DAILY_1,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nUNIT_CONTRACT_1,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nTOTALHA_1,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nOUTPUT_DAILY_1,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nSKU_2,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nKHL_2,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nTOTAL_2,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nUNIT_DAILY_2,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nUNIT_CONTRACT_2,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nTOTALHA_2,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nOUTPUT_DAILY_2,2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nOUTPUT_BUDGET,2)%></strong></td>
	</tr>


</table>
   
</div>
</body>
</html>
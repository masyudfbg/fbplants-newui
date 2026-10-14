<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsPerawatanTM, strPerawatanTM, strGroup, strOrder, strQuery
	Dim v_txtdatefrom, v_txtdateto, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass, v_status_tm_tbm, v_txtjobcode, v_txtfcba
	
	v_txtdatefrom		=	trim(request.Form("txtdatefrom"))
	v_txtdateto			=	trim(request.Form("txtdateto"))
	v_txtDivision		=	trim(request.Form("txtDivision"))
	v_txtBlok			=	trim(request.Form("txtBlok"))
	v_txtPlantingDate	=	trim(request.Form("txtPlantingDate"))
	v_status_tm_tbm     =   trim(request.Form("txtStatusTmTbm"))
	v_txtjobcode     	=   trim(request.Form("txtjobcode"))
	v_txtfcba     		=   trim(request.Form("txtfcba"))
	
	strPerawatanTM	=	" SELECT" & _
						" a.fcba," & _
						" a.division," & _
						" a.fieldcode," & _
						" a.fcname," & _
						" a.status," & _
						" a.plantingyear," & _
						" a.jobcode," & _
						" b.fcname," & _
						" SUM (a.sku) sku," & _
						" SUM (a.khl) khl," & _
						" SUM (a.total) total," & _
						" SUM (a.ha_harian) ha_harian," & _
						" SUM (a.bo_harian) bo_harian," & _
						" SUM (a.total_bo_harian) total_bo_harian," & _
						" DECODE ( SUM (a.total), 0, 0, SUM (a.ha_harian) / SUM (a.total))" & _
						" ha_per_hk," & _
						" sum(a.gli_ltr) gli_ltr," & _
						" DECODE (sum(a.total_bo_harian),0,0,sum(a.gli_ltr)/sum(a.total_bo_harian)) gli_ltr_ha," & _
						" sum(a.par_ltr) par_ltr," & _
						" DECODE (sum(a.total_bo_harian),0,0,sum(a.par_ltr)/sum(a.total_bo_harian)) par_ltr_ha," & _
						" sum(a.met_kg) met_kg," & _
						" DECODE (sum(a.total_bo_harian),0,0, sum(a.met_kg)/sum(a.total_bo_harian)) met_kg_ha," & _
						" sum(flo_kg) flo_kg," & _
						" DECODE (sum(a.total_bo_harian),0,0, sum(flo_kg)/sum(a.total_bo_harian)) flo_kg_ha" & _
						" FROM" & _
						" view_attendance_tm_smry a, job b" & _
						" WHERE" & _
						" (a.fcba = b.fcba and a.jobcode = b.fccode)" & _
						" and" & _
						" a.fcba = '" & v_txtfcba & "'" & _
						" AND" & _
						" fddate" & _
						" BETWEEN" & _
						" TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY')" & _
						" AND" & _
						" TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY')"
						if len(trim(v_status_tm_tbm)) > 0 then
							strPerawatanTM = strPerawatanTM & " AND a.STATUS = '" & v_status_tm_tbm & "'"
						end if
						if len(trim(v_txtDivision)) > 0 then
							strPerawatanTM = strPerawatanTM & " AND a.DIVISION = '" & v_txtDivision & "'"
						end if
						if len(trim(v_txtPlantingDate)) > 0 then
							strPerawatanTM = strPerawatanTM & " AND a.PLANTINGYEAR = '" & v_txtplantingdate & "'"
						end if
						if len(trim(v_txtBlok)) > 0 then
							strPerawatanTM = strPerawatanTM & " AND a.FIELDCODE = '" & v_txtBlok & "'"
						end if
						if len(trim(v_txtjobcode)) > 0 then
							strPerawatanTM = strPerawatanTM & " AND a.JOBCODE = '" & v_txtjobcode & "'"
						end if
						
	strGroup	=	" GROUP BY a.fcba, a.division, a.plantingyear, a.fieldcode, a.fcname, a.status, a.jobcode, b.fcname"
	strOrder	=	" ORDER BY a.fcba, a.division, a.plantingyear, a.fieldcode, a.fcname, a.status, a.jobcode"
	
	strQuery	=	strPerawatanTM + strGroup + strOrder

	Set rsPerawatanTM = Server.CreateObject("ADODB.Recordset")
	rsPerawatanTM.Open strQuery, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Maintenance Report</title>
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
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN HARIAN PERAWATAN</h3>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>DARI</th>
		<th>SAMPAI</th>
		<th>KEBUN</th>
		<th>AFDELING</th>
		<th>STATUS</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
		<th>PEKERJAAN</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_PerawatanTM_BA.asp">
		<td align="center">
		<% if v_txtdatefrom <> "" and v_txtdatefrom = v_txtdatefrom then %>
			<input type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtdatefrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
		<% if v_txtdateto <> "" and v_txtdateto = v_txtdateto then %>
			<input type="text" maxlength="10" name="txtdateto" size="10" value="<%=v_txtdateto%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtdateto" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtdateto');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
		<select name="txtfcba">	
		<option value="">--NONE--</option>		
		<%
			Dim rsFcba
			Set rsFcba = Server.CreateObject("ADODB.Recordset")
			rsFcba.CursorLocation = adUseClient
			strsql =  " SELECT fccode FROM businessunit "
			rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsFcba.eof then
				DO while not rsFcba.eof
				if v_txtFcba = rsFcba("fccode") then
		%>			<option value="<%=rsFcba("fccode")%>" Selected><%=rsFcba("fccode")%></option>
				<%else%>
					<option value="<%=rsFcba("fccode")%>"><%=rsFcba("fccode")%></option>
				<%end if%>
		<%			rsFcba.movenext
				loop
			End if
			rsFcba.Close
			set rsFcba = nothing
		%>
		</select>		
		</td>
		<td align="center">
		<select name="txtDivision">	
				<option value="">--ALL--</option>
		<%
			Dim rsDivision
			Set rsDivision = Server.CreateObject("ADODB.Recordset")
			rsDivision.CursorLocation = adUseClient
			strsql =  " SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
			rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsDivision.eof then
				DO while not rsDivision.eof
				if v_txtDivision = rsDivision("Division") then
		%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
				<%else%>
					<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
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
		<select name="txtStatusTmTbm">	
				<option value="">--ALL--</option>
		<%
			Dim rsStatusTmTbm
			Set rsStatusTmTbm = Server.CreateObject("ADODB.Recordset")
			rsStatusTmTbm.CursorLocation = adUseClient
			strsql =  " SELECT FCCODE as STATUSTMTBM from PARAMETER where parcode = 'F0006' and FCBA = '" & Session("fcBA") & "'"
			rsStatusTmTbm.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStatusTmTbm.eof then
				DO while not rsStatusTmTbm.eof
				if v_status_tm_tbm = rsStatusTmTbm("STATUSTMTBM") then
		%>			<option value="<%=rsStatusTmTbm("STATUSTMTBM")%>" Selected><%=rsStatusTmTbm("STATUSTMTBM")%></option>
				<%else%>
					<option value="<%=rsStatusTmTbm("STATUSTMTBM")%>"><%=rsStatusTmTbm("STATUSTMTBM")%></option>
				<%end if%>
		<%			rsStatusTmTbm.movenext
				loop
			End if
			rsStatusTmTbm.Close
			set rsStatusTmTbm = nothing
		%>
		</select>		
		</td>
		
		
		<td align="center">
		<select name="txtPlantingDate">	
				<option value="">--ALL--</option>
		<%
			Dim rsPlantingDate, strsql
			Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
			rsPlantingDate.CursorLocation = adUseClient
			strsql =  " SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
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
				<option value="">--ALL--</option>
		<%
			Dim rsBlok
			Set rsBlok = Server.CreateObject("ADODB.Recordset")
			rsBlok.CursorLocation = adUseClient
			strsql =  " SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
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
		<td>
				<select name="txtjobcode">	
				<option value="">--ALL--</option>
		<%
			Dim rsJob
			Set rsJob = Server.CreateObject("ADODB.Recordset")
			rsJob.CursorLocation = adUseClient
			strsql =  	" SELECT FCCODE, FCNAME FROM Job WHERE FCCODE IN (" & _
						" '6500101', '6500102', '6500103', '6500104', '6500201', '6500202', '6500203'," & _
						" '1322001', '1322002', '1322003', '1322004', '1322101', '1322102') AND FCBA = '" & Session("fcBA") & "' ORDER BY FCCODE ASC"
			rsJob.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsJob.eof then
				DO while not rsJob.eof
				if v_txtjobcode = rsJob("FCCODE") then
		%>			<option value="<%=rsJob("FCCODE")%>" Selected><%=rsJob("FCCODE")%> - <%=rsJob("FCNAME")%></option>
				<%else%>
					<option value="<%=rsJob("FCCODE")%>"><%=rsJob("FCCODE")%> - <%=rsJob("FCNAME")%></option>
				<%end if%>
		<%			rsJob.movenext
				loop
			End if
			rsJob.Close
			set rsJob = nothing
		%>
		</select>		
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="4">NO</th>
		<th rowspan="4">KEBUN</th>
		<th rowspan="4">AFDELING</th>
		<th rowspan="4">TT</th>
		<th rowspan="4">BLOK</th>
		<th rowspan="4">STATUS</th>
		<th rowspan="4">KODE PEKERJAAN</th>
		<th rowspan="4">NAMA PEKERJAAN</th>
		<th colspan="3">TENAGA KERJA</th>
		<th colspan="3">HASIL KERJA</th>
		<th>OUTPUT</th>
		<th colspan="8">PEMAKAIAN BAHAN</th>
	</tr>
	<tr>
		<th rowspan="3">SKU</th>
		<th rowspan="3">KHL</th>
		<th rowspan="3">TOTAL</th>
		<th rowspan="3">HARIAN</th>
		<th rowspan="3">BORONGAN</th>
		<th rowspan="3">TOTAL</th>
		<th rowspan="3">HA/HK</th>
		<th colspan="2">GLIFOSAT</th>
		<th colspan="2">PARAQUAT</th>
		<th colspan="2">METYL</th>
		<th colspan="2">FLOROKSIPIR</th>
	</tr>
	<tr>
		<th>QTY</th>
		<th>DOSIS</th>
		<th>QTY</th>
		<th>DOSIS</th>
		<th>QTY</th>
		<th>DOSIS</th>
		<th>QTY</th>
		<th>DOSIS</th>
	</tr>
	<tr>
		<th>LTR</th>
		<th>LTR/HA</th>
		<th>LTR</th>
		<th>LTR/HA</th>
		<th>KG</th>
		<th>KG/HA</th>
		<th>KG</th>
		<th>KG/HA</th>
	</tr>
<%
Dim nsubsku, nsubkhl, nsubtotal, nsubha_harian, nsubbo_harian
Dim nsubtotal_bo_harian, nsubha_per_hk, nsubgli_ltr, nsubgli_ltr_ha, nsubpar_ltr
Dim nsubpar_ltr_ha, nsubmet_kg, nsubmet_kg_ha, nsubflo_kg, nsubflo_kg_ha
	if not rsPerawatanTM.EOF then
		NO = 1
		nsubsku = 0
		nsubkhl = 0
		nsubtotal = 0
		nsubha_harian = 0
		nsubbo_harian = 0
		nsubtotal_bo_harian = 0
		nsubha_per_hk = 0
		nsubgli_ltr = 0
		nsubgli_ltr_ha = 0
		nsubpar_ltr = 0
		nsubpar_ltr_ha = 0
		nsubmet_kg = 0
		nsubmet_kg_ha = 0
		nsubflo_kg = 0
		nsubflo_kg_ha = 0

		while not rsPerawatanTM.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="center"><%=rsPerawatanTM(0)%></td>
				<td align="center"><%=rsPerawatanTM(1)%></td>
				<td align="center"><%=rsPerawatanTM(5)%></td>
				<td align="center"><%=rsPerawatanTM(2)%></td>
				<td align="center"><%=rsPerawatanTM(4)%></td>
				<td align="center"><%=rsPerawatanTM(6)%></td>
				<td align="left"><%=rsPerawatanTM(7)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(8),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(9),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(10),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(11),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(12),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(13),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(14),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(15),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(16),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(17),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(18),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(19),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(20),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(21),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM(22),2)%></td>
			</tr>
<%
			NO = NO + 1			
			nsubsku = nsubsku + Cdbl(rsPerawatanTM(8))
			nsubkhl = nsubkhl + Cdbl(rsPerawatanTM(9))
			nsubtotal = nsubtotal + Cdbl(rsPerawatanTM(10))
			nsubha_harian = nsubha_harian + Cdbl(rsPerawatanTM(11))
			nsubbo_harian = nsubbo_harian + Cdbl(rsPerawatanTM(12))
			nsubtotal_bo_harian = nsubtotal_bo_harian + Cdbl(rsPerawatanTM(13))
			nsubha_per_hk = nsubha_per_hk + Cdbl(rsPerawatanTM(14))
			nsubgli_ltr = nsubgli_ltr + Cdbl(rsPerawatanTM(15))
			nsubgli_ltr_ha = nsubgli_ltr_ha + Cdbl(rsPerawatanTM(16))
			nsubpar_ltr = nsubpar_ltr + Cdbl(rsPerawatanTM(17))
			nsubpar_ltr_ha = nsubpar_ltr_ha + Cdbl(rsPerawatanTM(18))
			nsubmet_kg = nsubmet_kg + Cdbl(rsPerawatanTM(19))
			nsubmet_kg_ha = nsubmet_kg_ha + Cdbl(rsPerawatanTM(20))
			nsubflo_kg = nsubflo_kg + Cdbl(rsPerawatanTM(21))
			nsubflo_kg_ha = nsubflo_kg_ha + Cdbl(rsPerawatanTM(22))

			rsPerawatanTM.MoveNext
		Wend
	end if
	rsPerawatanTM.Close
	set rsPerawatanTM = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="8">TOTAL</td>
				<td align="right"><%=FormatNumber(nsubsku,2)%></td>
				<td align="right"><%=FormatNumber(nsubkhl,2)%></td>
				<td align="right"><%=FormatNumber(nsubtotal,2)%></td>
				<td align="right"><%=FormatNumber(nsubha_harian,2)%></td>
				<td align="right"><%=FormatNumber(nsubbo_harian,2)%></td>
				<td align="right"><%=FormatNumber(nsubtotal_bo_harian,2)%></td>
				<td align="right"><%=FormatNumber(nsubha_per_hk,2)%></td>
				<td align="right"><%=FormatNumber(nsubgli_ltr,2)%></td>
				<td align="right"><%=FormatNumber(nsubgli_ltr_ha,2)%></td>
				<td align="right"><%=FormatNumber(nsubpar_ltr,2)%></td>
				<td align="right"><%=FormatNumber(nsubpar_ltr_ha,2)%></td>
				<td align="right"><%=FormatNumber(nsubmet_kg,2)%></td>
				<td align="right"><%=FormatNumber(nsubmet_kg_ha,2)%></td>
				<td align="right"><%=FormatNumber(nsubflo_kg,2)%></td>
				<td align="right"><%=FormatNumber(nsubflo_kg_ha,2)%></td>
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->

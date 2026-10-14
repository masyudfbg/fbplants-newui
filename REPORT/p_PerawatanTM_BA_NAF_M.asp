<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsPerawatanTM, strPerawatanTM, strGroup, strOrder, strQuery
	Dim v_txtdatefrom, v_txtdateto, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass, v_status_tm_tbm, v_txtjobcode, v_txtfcba
	Dim nSpanTotal
	Dim txtdatefrom, txtdateto, txtDivision, txtBlok, txtPlantingDate, txtStatusTmTbm, txtjobcode, txtfcba
	Dim sprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode, sprevStatus, sprevJobCode
	Dim tempDivision, tempPlantingDate,  tempBlok,  tempStatus,  tempJobcode, tempFcba 
	Dim ExtendedQuery, supportExtendedQuery, strExtendedQuery
    Dim rscomp , strcomp,txtcompany ,txcompanyname,txtComp,v_txtComp
	
	txtcompany          =  Request("txtComp") 
	txtComp             =  Request("txtComp") 
	v_txtComp           =   txtComp
	
	strcomp= " select a.fccompanycode,a.fccompanyname from BUSINESSUNIT a  where a.fccompanycode = '" & txtcompany & "'"
	Set rscomp = Server.CreateObject("ADODB.Recordset")
	rscomp.Open  strcomp,connstr, adOpenStatic
	if not rscomp.eof then

		txcompanyname   = 	rscomp("fccompanyname")

	End if
	
	
	nSpanTotal = 7
	
	txtdatefrom			= Request("txtdatefrom")
	txtdateto			= Request("txtdateto")
	txtDivision			= Request("txtDivision")
	txtBlok				= Request("txtBlok")
	txtPlantingDate		= Request("txtPlantingDate")
	txtStatusTmTbm     	= Request("txtStatusTmTbm")
	txtjobcode 			= Request("txtjobcode")
    txtfcba 			= Session("fcBA")

	
	
	
	
	v_txtdatefrom		=	txtdatefrom
	v_txtdateto			=	txtdateto
	v_txtDivision		=	txtDivision
	v_txtBlok			=	txtBlok
	v_txtPlantingDate	=	txtPlantingDate
	v_status_tm_tbm     =   txtStatusTmTbm
	v_txtjobcode     	=   txtjobcode
	v_txtfcba     		=   txtfcba
	
	
	
	
    strPerawatanTM = " SELECT a.status "
	
	supportExtendedQuery = 0
	if len(trim(v_txtjobcode)) > 0 then 
		if v_txtjobcode <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.jobcode, b.fcname as jobname "
		end if
	else
	    strPerawatanTM = strPerawatanTM & " , a.jobcode, b.fcname as jobname "
		supportExtendedQuery = 1
	end if	
	
	if len(trim(v_txtfcba)) > 0 then 
		if v_txtfcba <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.fcba "
		end if
	else
	    strPerawatanTM = strPerawatanTM & " , a.fcba "
	end if


	if len(trim(v_txtDivision)) > 0 then 
		if v_txtDivision <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.DIVISION "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , a.DIVISION "
	end if

	if len(trim(v_txtBlok)) > 0 then 
		if v_txtBlok <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.FIELDCODE, a.fcname "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , a.FIELDCODE, a.fcname  "	
	end if

	
	
	if len(trim(v_txtplantingdate)) > 0 then 
		if v_txtplantingdate <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.plantingyear "
		end if
	else
	    strPerawatanTM = strPerawatanTM & " , a.plantingyear "
	end if
	
	

	strPerawatanTM 		=  strPerawatanTM & " , round(SUM(nvl(a.sku,0)),2) sku, round(SUM(nvl(a.khl,0)),2) khl, " & _ 
							   " round(SUM(nvl(a.total,0)),2) total, round(SUM(nvl(a.ha_harian,0)),2) ha_harian, " & _ 
							   " round(SUM(nvl(a.bo_harian,0)),2) bo_harian, round(SUM(nvl(a.total_bo_harian,0)),2) total_bo_harian, " & _ 
							   " round(DECODE (SUM (a.total), " & _ 
									   " 0, 0, " & _ 
									   " SUM (a.ha_harian) / SUM (a.total) " & _ 
									  " ),2) ha_per_hk, " & _ 
							   " round(SUM(nvl(a.gli_ltr,0)),2) gli_ltr, " & _ 
							   " round(DECODE (SUM (a.total_bo_harian), " & _ 
									   " 0, 0, " & _ 
									   " SUM (a.gli_ltr) / SUM (a.total_bo_harian) " & _ 
									  " ),2) gli_ltr_ha, " & _ 
							   " round(SUM(nvl(a.par_ltr,0)),2) par_ltr, " & _ 
							   " round(DECODE (SUM (a.total_bo_harian), " & _ 
									   " 0, 0, " & _ 
									   " SUM (a.par_ltr) / SUM (a.total_bo_harian) " & _ 
									  " ),2) par_ltr_ha, " & _ 
							   " round(SUM(nvl(a.met_kg,0)),0) met_kg, " & _ 
							   " round(DECODE (SUM (a.total_bo_harian), " & _ 
									   " 0, 0, " & _ 
									   " SUM (a.met_kg) / SUM (a.total_bo_harian) " & _ 
									  " ),2) met_kg_ha, " & _ 
							   " round(SUM(nvl(flo_kg,0)),2) flo_kg, " & _ 
							   " round(DECODE(SUM (a.total_bo_harian), " & _ 
									   " 0, 0, " & _ 
									   " SUM (flo_kg) / SUM (a.total_bo_harian) " & _ 
									  " ),2) flo_kg_ha " & _ 
						  " FROM view_attendance_tm_smry a, job b " & _ 
						 " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) " 
	
	
	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
	        strPerawatanTM = strPerawatanTM & " AND a.fcba = '" & v_txtFcba& "' " 
		end if
	end if
	
	if len(trim(v_txtdatefrom)) > 0 and  len(trim(v_txtdateto)) then 
	    strPerawatanTM = strPerawatanTM & " AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _ 
										  " AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') "				  
	end if
	
	
    ' Where Condion
	if len(trim(v_status_tm_tbm)) > 0 then 
		if v_status_tm_tbm <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " AND a.STATUS = '" & v_status_tm_tbm & "'"
		end if
	end if
	
    if len(trim(v_txtjobcode)) > 0  then
		if v_txtjobcode <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " AND a.JOBCODE = '" & v_txtjobcode & "'"
		end if 
	end if	

	
	if len(trim(v_txtDivision)) > 0 then 
		if  v_txtDivision <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " AND a.DIVISION = '" & v_txtDivision & "'"
		end if
	end if
	
	if len(trim(v_txtPlantingDate)) > 0  then
		if v_txtPlantingDate <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " AND a.PLANTINGYEAR = '" & v_txtplantingdate & "'"
		end if 
	end if
	
	
    if len(trim(v_txtBlok)) > 0  then
		if v_txtBlok <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " AND a.FIELDCODE = '" & v_txtBlok & "'"
		end if 
	end if
	


	' Group By
	strPerawatanTM = strPerawatanTM & " GROUP BY a.status "
	

	
	if len(trim(v_txtjobcode)) > 0 then 
		if v_txtjobcode <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.jobcode, b.fcname "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , a.jobcode, b.fcname  "	
	end if	
	
	if len(trim(v_txtfcba)) > 0 then 
		if v_txtfcba <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.fcba "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , a.fcba "	
	end if	
	
	if len(trim(v_txtDivision)) > 0 then 
		if v_txtDivision <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.DIVISION "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , a.DIVISION "
	end if
	
	if len(trim(v_txtPlantingDate)) > 0 then 
		if v_txtPlantingDate <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.PLANTINGYEAR "
		end if
	else
	    strPerawatanTM = strPerawatanTM & " , a.PLANTINGYEAR "
	end if	

	if len(trim(v_txtBlok)) > 0 then 
		if v_txtBlok <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , a.FIELDCODE, a.fcname "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , a.FIELDCODE, a.fcname "	
	end if


	
	' Order By
	strPerawatanTM = strPerawatanTM & " ORDER BY status "
	
	if len(trim(v_txtjobcode)) > 0 then 
		if v_txtjobcode <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , jobcode "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , jobcode  "	
	end if	

	if len(trim(v_txtfcba)) > 0 then 
		if v_txtfcba <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , fcba "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , fcba "	
	end if

	
	if len(trim(v_txtDivision)) > 0 then 
		if v_txtDivision <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , DIVISION "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , DIVISION "
	end if
	
	if len(trim(v_txtPlantingDate)) > 0 then 
		if v_txtPlantingDate <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , PLANTINGYEAR "
		end if
	else
	    strPerawatanTM = strPerawatanTM & " , PLANTINGYEAR "
	end if	

	if len(trim(v_txtBlok)) > 0 then 
		if v_txtBlok <> "--SUMMARY--" then
			strPerawatanTM = strPerawatanTM & " , FIELDCODE "
		end if
	else
		strPerawatanTM = strPerawatanTM & " , FIELDCODE "	
	end if



	' Make String combination if the Value of Jobcode is ZERO
	' By NAF
	' the fcba, division, plantingdate and block must be in --SUMMARY-- condition
	
	
	ExtendedQuery = 0
	if v_txtfcba = "--SUMMARY--" then 
	    if v_txtDivision = "--SUMMARY--" then 
			if v_txtPlantingDate = "--SUMMARY--" then 
				if v_txtblok = "--SUMMARY--" then 
					if supportExtendedQuery = 1 then 
				        ExtendedQuery = 1
					end if
				end if
			end if
		end if
	end if
			    
    if ExtendedQuery = 1 then 
	    strExtendedQuery = " select status, jobcode, jobname, " & _
									 " SUM (sku) sku, " & _
									 " SUM (khl) khl,  " & _
									 " SUM (total) total,  " & _
									 " SUM (ha_harian) ha_harian, " & _
									 " SUM (bo_harian) bo_harian,  " & _
									 " SUM (total_bo_harian) total_bo_harian, " & _
									 " SUM (ha_per_hk) ha_per_hk, " & _
									 " SUM (gli_ltr) gli_ltr, " & _
									 " SUM (gli_ltr_ha) gli_ltr_ha, " & _
									 " SUM (par_ltr) par_ltr, " & _
									 " SUM (par_ltr_ha) par_ltr_ha, " & _
									 " SUM (met_kg) met_kg, " & _
									 " SUM (met_kg_ha) met_kg_ha, " & _
									 " SUM (flo_kg) flo_kg, " & _
									 " SUM (flo_kg_ha) flo_kg_ha " & _
							" from ( " & _
							" select a_temp.job_forfieldstatus status,  a_temp.fccode jobcode, a_temp.fcname jobname, " & _
							  " 0 sku, 0 khl, 0 total,  " & _
							  " 0 ha_harian, 0 bo_harian, 0 total_bo_harian, " & _
							  " 0 ha_per_hk, " & _
							  " 0 gli_ltr, 0 gli_ltr_ha, " & _
							  " 0 par_ltr, 0 par_ltr_ha, " & _
							  " 0 met_kg, 0 met_kg_ha, " & _
							  " 0 flo_kg, 0 flo_kg_ha from job a_temp, " & _
							" (select fcba, param_1 from parameterdetail " & _ 
							" where  " & _
							" parcode = 'SR002' and " & _
							" parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') ) b_temp " & _
							" where " & _
							" a_temp.fccode = b_temp.param_1 " & _
							" and a_temp.fcba = b_temp.fcba " 
							
		if len(trim(v_status_tm_tbm)) > 0  then
			if v_status_tm_tbm <> "--SUMMARY--" then
				strExtendedQuery = strExtendedQuery & " AND a_temp.job_forfieldstatus = '" & v_status_tm_tbm & "' "
			end if 
		end if							
		
		strExtendedQuery = strExtendedQuery & " UNION " & strPerawatanTM
        strExtendedQuery = strExtendedQuery & ") " & _ 
												" GROUP BY status, jobcode, jobname " & _
												" ORDER BY status, jobcode " 
												
	
	
		Set rsPerawatanTM = Server.CreateObject("ADODB.Recordset")
		response.write strExtendedQuery
		
		rsPerawatanTM.Open strExtendedQuery, connstr, adOpenStatic 
	else
		Set rsPerawatanTM = Server.CreateObject("ADODB.Recordset")
		response.write strPerawatanTM
		
		rsPerawatanTM.Open strPerawatanTM, connstr, adOpenStatic 	
    end if	

	
	
	' End of Modification
	
	
	


	
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
<form name="frmEntry" method="post" action="p_PerawatanTM_BA_NAF.asp">

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/DashboardGroupManager.asp">Back to Report</a></td>
		</tr>
	</table>
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN HARIAN PERAWATAN</h3>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%>, KEBUN : <%=v_txtFcba%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%>, KEBUN : <%=v_txtFcba%></h5>
		<% end if %>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>DARI</th>
		<th>SAMPAI</th>
		<th>STATUS</th>
		<th>PEKERJAAN</th>
		<th>AFDELING</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
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
		<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
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
		<select name="txtStatusTmTbm">	
			<option value="">--ALL--</option>
		<%
			Dim rsStatusTmTbm
			Set rsStatusTmTbm = Server.CreateObject("ADODB.Recordset")
			rsStatusTmTbm.CursorLocation = adUseClient
			
			
			if len(trim(v_status_tm_tbm)) > 0 then 
				strsql =  "  SELECT FCCODE as STATUSTMTBM from PARAMETER where parcode = 'F0006' and FCBA = '" & v_txtFcba & "'"
			else
				strsql =  "  SELECT FCCODE as STATUSTMTBM from PARAMETER where parcode = 'F0006' and FCBA = '" & _
				Session("fcBA") & "'"
			end if
			
			
			
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
		
		
		<td>
				<select name="txtjobcode">	
				<option value="">--ALL--</option>
		<%
			Dim rsJob
			Set rsJob = Server.CreateObject("ADODB.Recordset")
			rsJob.CursorLocation = adUseClient
			
			' Get the Activity from ParameterDetail Table
			
			if len(trim(v_txtBlok)) > 0 then 
			    strsql = " SELECT '--SUMMARY--' as fccode, '' as fcname  from DUAL UNION SELECT FCCODE, FCNAME FROM Job WHERE FCCODE IN ( " & _
						 " select param_1 from parameterdetail " & _
						 " where " & _
						 " fcba = '" & v_txtFcba & "' and " & _
						 " parcode = 'SR002' and " & _
						 " parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 " ) AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC " 
			
			else
			    strsql = " SELECT '--SUMMARY--' as fccode, '' as fcname  from DUAL UNION SELECT FCCODE, FCNAME FROM Job WHERE FCCODE IN ( " & _
						 " select param_1 from parameterdetail " & _
						 " where " & _
						 " fcba = '" & Session("fcBA") & "' and " & _
						 " parcode = 'SR002' and " & _
						 " parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 " ) AND FCBA = '" & Session("fcBA") & "' ORDER BY FCCODE ASC " 

			end if		
			
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
		
		
		
		
		
		<td align="center">
		<select name="txtDivision">	
			<option value="">--ALL--</option>
		<%
			Dim rsDivision
			Set rsDivision = Server.CreateObject("ADODB.Recordset")
			rsDivision.CursorLocation = adUseClient
			

		
			
			if len(trim(v_txtDivision)) > 0 then 
				strsql =  " SELECT '--SUMMARY--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--SUMMARY--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
			end if
			
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
		<select name="txtPlantingDate">	
				<option value="">--ALL--</option>
		<%
			Dim rsPlantingDate, strsql
			Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
			rsPlantingDate.CursorLocation = adUseClient
			

			if len(trim(v_status_tm_tbm)) > 0 then 
				strsql =  " SELECT '--SUMMARY--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & _
				v_txtFcba & "'"
			else
				strsql =  " SELECT '--SUMMARY--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & _
				Session("fcBA") & "'"
			end if

			
			
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
			
			if len(trim(v_txtBlok)) > 0 then 
				strsql =  " SELECT '--SUMMARY--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--SUMMARY--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
			end if			
			
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
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="4">NO</th>

		<% if trim(request("txtStatusTmTbm")) <> "--SUMMARY--" then %>
		    <th rowspan="4">STATUS</th>
		<% else
		    nSpanTotal = nSpanTotal - 1 		
	    end if %>

		<% if trim(request("txtjobcode")) <> "--SUMMARY--" then %>
			<th rowspan="4">KODE PEKERJAAN</th>
			<th rowspan="4">NAMA PEKERJAAN</th>
		<% else
		    nSpanTotal = nSpanTotal - 2 		
	    end if %>
		
		<% if trim(request("txtfcba")) <> "--SUMMARY--" then %>		
		    <th rowspan="4">KEBUN</th>
		<% else
		    nSpanTotal = nSpanTotal - 1 		
	    end if %>

		
		<% if trim(request("txtDivision")) <> "--SUMMARY--" then %>
		    <th rowspan="4">AFDELING</th>
		<% else
			nSpanTotal = nSpanTotal - 1 			
	    end if %>
		
		<% if trim(request("txtPlantingDate")) <> "--SUMMARY--" then %>
		    <th rowspan="4">TT</th>
		<% else
		    nSpanTotal = nSpanTotal - 1 			
	    end if %>

		<% if trim(request("txtBlok")) <> "--SUMMARY--" then %>
		    <th rowspan="4">BLOK</th>
		<% else
		    nSpanTotal = nSpanTotal - 1 		
	    end if %>		

		
		
		
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
				
				
				<% if trim(request("txtStatusTmTbm")) <> "--SUMMARY--" then %>				
					<% if sprevStatus <> rsPerawatanTM("status") then %>

						<% tempStatus = rsPerawatanTM("status") %>
						<% tempJobcode = "" %>
						<% tempfcba = "--SUMMARY--" %>
						<% tempDivision = "--SUMMARY--" %>
						<% tempPlantingDate = "--SUMMARY--" %>
						<% tempBlok = "--SUMMARY--" %>
					
						

						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtComp=<%=v_txtComp%>">
							<%=rsPerawatanTM("status")%>
							</a>
						</td>
						<% sprevStatus = rsPerawatanTM("status") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>	
				
				<% if trim(request("txtjobcode")) <> "--SUMMARY--" then %>								
					<% if sprevJobCode <> rsPerawatanTM("jobcode") then %>
					
						
					    <% if trim(request("txtStatusTmTbm")) <> "--SUMMARY--" then %>
							<% tempStatus = rsPerawatanTM("status") %>
						<% else %>
							<% tempStatus = "" %>
						<% end if %>

						<% tempJobcode = rsPerawatanTM("jobcode") %>
						<% tempFcba = "" %>
						<% tempDivision = "--SUMMARY--" %>
					    <% tempPlantingDate = "--SUMMARY--" %>
						<% tempBlok = "--SUMMARY--" %>
						
						
						
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=tempFcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtComp=<%=v_txtComp%>">
							<%=rsPerawatanTM("jobcode") %>
							</a>
							<td align="left"><%=rsPerawatanTM("jobname")%></td>
						</td>
						
						
						<% sprevJobCode = rsPerawatanTM("jobcode") %>	
					<% else	%>
						<td align="center"></td>					
						<td align="center"></td>					
					<% end if %>	
				<% end if %>	
				
				
				<% if trim(request("txtfcba")) <> "--SUMMARY--" then %>								
					<% if sprevFcba <> rsPerawatanTM("fcba") then %>
					
					    <% if trim(request("txtStatusTmTbm")) <> "--SUMMARY--" then %>
							<% tempStatus = rsPerawatanTM("status") %>
						<% else %>
							<% tempStatus = "" %>
						<% end if %>

					    <% if trim(request("txtjobcode")) <> "--SUMMARY--" then %>
							<% tempJobcode = rsPerawatanTM("jobcode") %>
						<% else %>
							<% tempJobcode = "" %>
						<% end if %>						
					
						<% tempfcba = rsPerawatanTM("fcba") %>
						
						<% tempDivision = "" %>
						<% tempPlantingDate = "--SUMMARY--" %>
						<% tempBlok = "--SUMMARY--" %>
						
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtComp=<%=v_txtComp%>">
							<%=rsPerawatanTM("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsPerawatanTM("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>							 
				<% end if %>							 


				<% if trim(request("txtDivision")) <> "--SUMMARY--" then %>
					<% if sprevDivision <> rsPerawatanTM("division") then %>
					
					
					    <% if trim(request("txtStatusTmTbm")) <> "--SUMMARY--" then %>
							<% tempStatus = rsPerawatanTM("status") %>
						<% else %>
							<% tempStatus = "" %>
						<% end if %>

					    <% if trim(request("txtjobcode")) <> "--SUMMARY--" then %>
							<% tempJobcode = rsPerawatanTM("jobcode") %>
						<% else %>
							<% tempJobcode = "" %>
						<% end if %>
						
						<% if trim(request("txtfcba")) <> "--SUMMARY--" then %>
							<% tempfcba = rsPerawatanTM("fcba") %>
						<% else %>
							<% tempfcba = "" %>
						<% end if %>
					
					    <% tempDivision = rsPerawatanTM("division") %>
						
						
						<% tempPlantingDate = "" %>
						<% tempBlok = "--SUMMARY--" %>
						
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=tempFcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtComp=<%=v_txtComp%>">
							<%=rsPerawatanTM("division")%>
							</a>
						</td>
						<% sprevDivision = rsPerawatanTM("division") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>	
				
				<% if trim(request("txtPlantingDate")) <> "--SUMMARY--" then %>				
					<% if sprevPlantingYear <> rsPerawatanTM("plantingyear") then %>
					    <% if trim(request("txtStatusTmTbm")) <> "--SUMMARY--" then %>
							<% tempStatus = rsPerawatanTM("status") %>
						<% else %>
							<% tempStatus = "" %>
						<% end if %>

					    <% if trim(request("txtjobcode")) <> "--SUMMARY--" then %>
							<% tempJobcode = rsPerawatanTM("jobcode") %>
						<% else %>
							<% tempJobcode = "" %>
						<% end if %>
						
						<% if trim(request("txtfcba")) <> "--SUMMARY--" then %>
							<% tempfcba = rsPerawatanTM("fcba") %>
						<% else %>
							<% tempfcba = "" %>
						<% end if %>
						
						<% if trim(request("txtdivision")) <> "--SUMMARY--" then %>
							<% tempDivision = rsPerawatanTM("division") %>
						<% else %>
							<% tempdivision = "" %>
						<% end if %>
						
						<% tempPlantingDate = rsPerawatanTM("plantingyear") %>
						
						<% tempBlok = "" %>
						
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtComp=<%=v_txtComp%>">
							<%=rsPerawatanTM("plantingyear")%>
							</a>
						</td>
						<% sprevPlantingYear = rsPerawatanTM("plantingyear") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>	
					

				<% if trim(request("txtBlok")) <> "--SUMMARY--" then %>				
					<% if sprevFieldCode <> rsPerawatanTM("fieldcode") then %>

					    <% if trim(request("txtStatusTmTbm")) <> "--SUMMARY--" then %>
							<% tempStatus = rsPerawatanTM("status") %>
						<% else %>
							<% tempStatus = "" %>
						<% end if %>

					    <% if trim(request("txtjobcode")) <> "--SUMMARY--" then %>
							<% tempJobcode = rsPerawatanTM("jobcode") %>
						<% else %>
							<% tempJobcode = "" %>
						<% end if %>
						
						<% if trim(request("txtfcba")) <> "--SUMMARY--" then %>
							<% tempfcba = rsPerawatanTM("fcba") %>
						<% else %>
							<% tempfcba = "" %>
						<% end if %>
						
						<% if trim(request("txtdivision")) <> "--SUMMARY--" then %>
							<% tempDivision = rsPerawatanTM("division") %>
						<% else %>
							<% tempdivision = "" %>
						<% end if %>
						
						<% if trim(request("txtPlantingDate")) <> "--SUMMARY--" then %>
							<% tempPlantingDate = rsPerawatanTM("plantingyear") %>
						<% else %>
							<% tempPlantingDate = "" %>
						<% end if %>
						
						<% tempBlok = rsPerawatanTM("fieldcode") %>
						
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=rsPerawatanTM("fcba")%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtComp=<%=v_txtComp%>">
							<%=rsPerawatanTM("fieldcode")%>
							</a>
						</td>
						<% sprevFieldCode = rsPerawatanTM("fieldcode") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>	


				
				
				<td align="right"><%=FormatNumber(rsPerawatanTM("sku"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("khl"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("total"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("ha_harian"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("bo_harian"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("total_bo_harian"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("ha_per_hk"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("gli_ltr"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("gli_ltr_ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("par_ltr"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("par_ltr_ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("met_kg"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("met_kg_ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("flo_kg"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("flo_kg_ha"),2)%></td>
			</tr>
<%
			NO = NO + 1			
			nsubsku = nsubsku + Cdbl(rsPerawatanTM("sku"))
			nsubkhl = nsubkhl + Cdbl(rsPerawatanTM("khl"))
			nsubtotal = nsubtotal + Cdbl(rsPerawatanTM("total"))
			nsubha_harian = nsubha_harian + Cdbl(rsPerawatanTM("ha_harian"))
			nsubbo_harian = nsubbo_harian + Cdbl(rsPerawatanTM("bo_harian"))
			nsubtotal_bo_harian = nsubtotal_bo_harian + Cdbl(rsPerawatanTM("total_bo_harian"))
			nsubha_per_hk = nsubha_per_hk + Cdbl(rsPerawatanTM("ha_per_hk"))
			nsubgli_ltr = nsubgli_ltr + Cdbl(rsPerawatanTM("gli_ltr"))
			nsubgli_ltr_ha = nsubgli_ltr_ha + Cdbl(rsPerawatanTM("gli_ltr_ha"))
			nsubpar_ltr = nsubpar_ltr + Cdbl(rsPerawatanTM("par_ltr"))
			nsubpar_ltr_ha = nsubpar_ltr_ha + Cdbl(rsPerawatanTM("par_ltr_ha"))
			nsubmet_kg = nsubmet_kg + Cdbl(rsPerawatanTM("met_kg"))
			nsubmet_kg_ha = nsubmet_kg_ha + Cdbl(rsPerawatanTM("met_kg_ha"))
			nsubflo_kg = nsubflo_kg + Cdbl(rsPerawatanTM("flo_kg"))
			nsubflo_kg_ha = nsubflo_kg_ha + Cdbl(rsPerawatanTM("flo_kg_ha"))

			rsPerawatanTM.MoveNext
		Wend
	end if
	rsPerawatanTM.Close
	set rsPerawatanTM = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 1 + nSpanTotal %>">TOTAL</td>
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

</form>

<!--#include file="../Bottom_Report.asp"-->

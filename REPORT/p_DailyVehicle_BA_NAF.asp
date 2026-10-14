<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim txtDateFrom, txtDateTo, txtVehicleGroup, txtVehicleCode, txtJobCode, txtFcba
	Dim v_txtDateFrom, v_txtDateTo, v_txtVehicleGroup, v_txtVehicleCode, v_txtJobCode, v_txtFcba
	Dim rsDailyVehicle, strDailyVehicle, strMain, strsql
	Dim NO, cClass, nSpanTotal
	Dim sprevVehicleCode, sprevVehicleGroup, sprevFcba
	
	txtDateFrom		=	trim(request("txtDateFrom"))
	txtDateTo		=	trim(request("txtDateTo"))
	txtVehicleGroup	=	trim(request("txtVehicleGroup"))
	txtVehicleCode	=	trim(request("txtVehicleCode"))
	txtJobCode		=	trim(request("txtJobCode"))
	txtFcba			=	trim(request("txtFcba"))
	
	v_txtDateFrom		=	txtDateFrom
	v_txtDateTo			=	txtDateTo
	v_txtVehicleGroup	=	txtVehicleGroup
	v_txtVehicleCode	=	txtVehicleCode
	v_txtJobCode		=	txtJobCode
	v_txtFcba			=	txtFcba
	
	nSpanTotal = 6
	
	strMain			 = 	" SELECT" & _
						" a.fcba fcba," & _
						" a.vehicledate vehicledate," & _
						" b.vehiclegroupcode vehiclegroupcode," & _
						" c.description vehiclegroupname," & _
						" a.vehiclecode vehiclecode," & _
						" b.fcname vehiclename," & _
						" a.jobcode jobcode," & _
						" d.fcname jobname," & _
						" a.locationtype locationtype," & _
						" a.locationcode locationcode," & _
						" SUM(nvl(a.unit,0)) unit," & _
						" SUM(nvl(a.hours,0)) hours," & _
						" SUM(nvl(a.output,0)) output" & _
						" FROM" & _
						" vehicleactivity a," & _
						" vehicle b," & _
						" vehiclegroup c," & _
						" job d" & _
						" WHERE" & _
						" vehicledate BETWEEN TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
						" AND TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
						" AND a.jobcode = d.fccode" & _
						" AND a.vehiclecode = b.fccode" & _
						" AND b.vehiclegroupcode = c.vehiclegroupcode" & _
						" AND a.fcba = b.fcba" & _
						" AND a.fcba = d.fcba" & _
						" AND b.fcba = c.fcba" & _
						" GROUP BY" & _
						" a.fcba," & _
						" a.vehicledate," & _
						" b.vehiclegroupcode," & _
						" c.description," & _
						" a.vehiclecode," & _
						" b.fcname," & _
						" a.jobcode," & _
						" d.fcname," & _
						" a.locationtype," & _
						" a.locationcode"
						
	strDailyVehicle	=	" SELECT FCBA"
	
	if len(trim(v_txtVehicleGroup)) > 0 then 
		if v_txtVehicleGroup = "--SUMMARY--" then
			strDailyVehicle = strDailyVehicle & ""
		else
			strDailyVehicle = strDailyVehicle & " , VEHICLEGROUPCODE, VEHICLEGROUPNAME"
		end if
	else
		strDailyVehicle = strDailyVehicle & " , VEHICLEGROUPCODE, VEHICLEGROUPNAME"
	end if
	
	if len(trim(v_txtVehicleCode)) > 0 then 
		if v_txtVehicleCode = "--SUMMARY--" then
			strDailyVehicle = strDailyVehicle & ""
		else
			strDailyVehicle = strDailyVehicle & " , VEHICLECODE, VEHICLENAME, JOBCODE, JOBNAME, LOCATIONTYPE, LOCATIONCODE"
		end if
	else
		strDailyVehicle = strDailyVehicle & " , VEHICLECODE, VEHICLENAME, JOBCODE, JOBNAME, LOCATIONTYPE, LOCATIONCODE"
	end if
	
	strDailyVehicle	=	strDailyVehicle & " , SUM(UNIT) UNIT, SUM(HOURS) HOURS, SUM(OUTPUT) OUTPUT"
	strDailyVehicle	=	strDailyVehicle & " FROM ( " & strMain & " )  "

	if len(trim(v_txtFcba)) > 0 then 
	    strDailyVehicle = strDailyVehicle & " WHERE FCBA = '" & v_txtFcba & "'" 
	end if
	
	if len(trim(v_txtVehicleGroup)) > 0 then 
		if v_txtVehicleGroup = "--SUMMARY--" then
			strDailyVehicle = strDailyVehicle & ""
		else
			strDailyVehicle = strDailyVehicle & " AND VEHICLEGROUPCODE = '" & v_txtVehicleGroup & "'" 
		end if
	end if
	
	if len(trim(v_txtVehicleCode)) > 0 then
		if v_txtVehicleCode = "--SUMMARY--" then
			strDailyVehicle = strDailyVehicle & ""
		else
			strDailyVehicle = strDailyVehicle & " AND VEHICLECODE = '" & v_txtVehicleCode & "'" 
		end if
	end if
	
	strDailyVehicle	=	strDailyVehicle & " GROUP BY FCBA"

	if len(trim(v_txtVehicleGroup)) > 0 then 
		if v_txtVehicleGroup = "--SUMMARY--" then
			strDailyVehicle = strDailyVehicle & ""
		else
			strDailyVehicle = strDailyVehicle & " , VEHICLEGROUPCODE, VEHICLEGROUPNAME"
		end if
	else
		strDailyVehicle = strDailyVehicle & " , VEHICLEGROUPCODE, VEHICLEGROUPNAME"
	end if

	if len(trim(v_txtVehicleCode)) > 0 then 
		if v_txtVehicleCode = "--SUMMARY--" then
			strDailyVehicle = strDailyVehicle & ""
		else
			strDailyVehicle = strDailyVehicle & " , VEHICLECODE, VEHICLENAME, JOBCODE, JOBNAME, LOCATIONTYPE, LOCATIONCODE"
		end if
	else
		strDailyVehicle = strDailyVehicle & " , VEHICLECODE, VEHICLENAME, JOBCODE, JOBNAME, LOCATIONTYPE, LOCATIONCODE"
	end if
	
	response.write strDailyVehicle
	Set rsDailyVehicle = Server.CreateObject("ADODB.Recordset")
	rsDailyVehicle.Open strDailyVehicle, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Daily Vehicle Report</title>
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
<form name="frmEntry" method="post" action="p_DailyVehicle_BA_NAF.asp">
	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/DashboardManager.asp">Back to Report</a></td>
		</tr>
	</table>
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN HARIAN KENDARAAN</h3>
		<h5>TANGGAL: <%=v_txtDateFrom%> s/d <%=v_txtDateTo%></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>DARI TANGGAL</th>
		<th>SAMPAI TANGGAL</th>
		<th>KEBUN</th>
		<th>VEHICLE GROUP</th>
		<th>VEHICLE CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>

		<td align="center">
		<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
			<input type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDateFrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>


		<td align="center">
		<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
			<input type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDateTo');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		
		<td align="center">
		<select name="txtFcba">	
		<option value="">--ALL--</option>		
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
	
				<td>
				<select name="txtVehicleGroup">	
				<option value="">--ALL--</option>
		<%
			Dim rsVehicleGroup
			Set rsVehicleGroup = Server.CreateObject("ADODB.Recordset")
			rsVehicleGroup.CursorLocation = adUseClient
			
			if len(trim(v_txtVehicleGroup)) > 0 then 
				strsql =	" SELECT '--SUMMARY--' as VehicleGroupCode, '--SUMMARY--' as Description  from DUAL UNION" & _
							" SELECT VehicleGroupCode, Description FROM VehicleGroup WHERE FCBA = '" & v_txtFcba & "' ORDER BY Description ASC"
			else
				strsql =	" SELECT '--SUMMARY--' as VehicleGroupCode, '--SUMMARY--' as Description  from DUAL UNION" & _
							" SELECT VehicleGroupCode, Description FROM VehicleGroup WHERE FCBA = '" & Session("fcBA") & "' ORDER BY Description ASC"
			end if
			
			rsVehicleGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsVehicleGroup.eof then
				DO while not rsVehicleGroup.eof
				if v_txtVehicleGroup = rsVehicleGroup("VehicleGroupCode") then
		%>			<option value="<%=rsVehicleGroup("VehicleGroupCode")%>" Selected><%=rsVehicleGroup("Description")%></option>
				<%else%>
					<option value="<%=rsVehicleGroup("VehicleGroupCode")%>"><%=rsVehicleGroup("Description")%></option>
				<%end if%>
		<%			rsVehicleGroup.movenext
				loop
			End if
			rsVehicleGroup.Close
			set rsVehicleGroup = nothing
		%>
		</select>		
		</td>
				<td>
				<select name="txtVehicleCode">	
				<option value="">--ALL--</option>
		<%
			Dim rsVehicleCode
			Set rsVehicleCode = Server.CreateObject("ADODB.Recordset")
			rsVehicleCode.CursorLocation = adUseClient
			
			if len(trim(v_txtVehicleCode)) > 0 then 
				strsql =	" SELECT '--SUMMARY--' as FCCODE, '--SUMMARY--' as FCNAME  from DUAL UNION" & _
							" SELECT FCCODE, FCNAME FROM Vehicle WHERE FCBA = '" & v_txtFcba & "' ORDER BY FCNAME ASC"
			else
				strsql =	" SELECT '--SUMMARY--' as FCCODE, '--SUMMARY--' as FCNAME  from DUAL UNION" & _
							" SELECT FCCODE, FCNAME FROM Vehicle WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCNAME ASC"
			end if
			rsVehicleCode.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsVehicleCode.eof then
				DO while not rsVehicleCode.eof
				if v_txtVehicleCode = rsVehicleCode("FCCODE") then
		%>			<option value="<%=rsVehicleCode("FCCODE")%>" Selected><%=rsVehicleCode("FCNAME")%></option>
				<%else%>
					<option value="<%=rsVehicleCode("FCCODE")%>"><%=rsVehicleCode("FCNAME")%></option>
				<%end if%>
		<%			rsVehicleCode.movenext
				loop
			End if
			rsVehicleCode.Close
			set rsVehicleCode = nothing
		%>
		</select>		
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">

	<tr>
		<th>NO</th>
		<th>KEBUN</th>
		
		<% if v_txtVehicleGroup <> "--SUMMARY--" then %>
			<th>GROUP KENDARAAN</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtVehicleCode <> "--SUMMARY--" then %>
			<th>NAMA KENDARAAN</th>
			<th>KODE PEKERJAAN</th>
			<th>JENIS PEKERJAAN</th>
			<th>TARGET TYPE</th>
			<th>TARGET CODE</th>
		<% else
			nSpanTotal = nSpanTotal - 5
	    end if %>
		
		<th>JUMLAH KM/ HM</th>
		<th>JUMLAH JAM KERJA</th>
		<th>OUTPUT</th>
	</tr>

<%
Dim nTotal_0
Dim nTotal_1
Dim nTotal_2
	if not rsDailyVehicle.EOF then
		nTotal_0 = 0
		nTotal_1 = 0
		nTotal_2 = 0
		NO = 1
		
		while not rsDailyVehicle.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<% if v_txtFcba = rsDailyVehicle("FCBA") then %>			
					    <td align="center"><%=rsDailyVehicle("FCBA")%></td>
						<% sprevFcba  = rsDailyVehicle("FCBA") %>
				<% else	%>
					    <td align="center"></td>									 
				<% end if %>
				
				<% if v_txtVehicleGroup <> "--SUMMARY--" then %>
				    <% if sprevVehicleGroup  <> rsDailyVehicle("VEHICLEGROUPCODE") then %>				
					    <td align="left"><%=rsDailyVehicle("VEHICLEGROUPNAME")%></td>
						<% sprevVehicleGroup  = rsDailyVehicle("VEHICLEGROUPCODE") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
				<% if v_txtVehicleCode <> "--SUMMARY--" then %>
				    <% if sprevVehicleCode <> rsDailyVehicle("VEHICLECODE") then %>				
						<td align="left"><%=rsDailyVehicle("VEHICLENAME")%></td>
						<td align="center"><%=rsDailyVehicle("JOBCODE")%></td>
						<td align="left"><%=rsDailyVehicle("JOBNAME")%></td>
						<td align="center"><%=rsDailyVehicle("LOCATIONTYPE")%></td>
						<td align="left"><%=rsDailyVehicle("LOCATIONCODE")%></td>
						<% sprevVehicleCode = rsDailyVehicle("VEHICLECODE") %>
					<% else	%>
						<td align="left"><%=rsDailyVehicle("VEHICLENAME")%></td>
						<td align="center"><%=rsDailyVehicle("JOBCODE")%></td>
						<td align="left"><%=rsDailyVehicle("JOBNAME")%></td>
						<td align="center"><%=rsDailyVehicle("LOCATIONTYPE")%></td>
						<td align="left"><%=rsDailyVehicle("LOCATIONCODE")%></td>				
					<% end if %>				
				<% end if %>
				
				<td align="right"><%=FormatNumber(rsDailyVehicle("unit"),2)%></%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicle("hours"),2)%></%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicle("output"),2)%></%></td>

			</tr>
<%
			NO = NO + 1
			nTotal_0 = nTotal_0 + Cdbl(rsDailyVehicle("unit"))
			nTotal_1 = nTotal_1 + Cdbl(rsDailyVehicle("hours"))
			nTotal_2 = nTotal_2 + Cdbl(rsDailyVehicle("output"))

			
			rsDailyVehicle.MoveNext
		Wend
	end if
	rsDailyVehicle.Close
	set rsDailyVehicle = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 2 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nTotal_0,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>
				
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->

<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsDailyVehicle, strDailyVehicle, strGroup, strOrder, strsql, v_txtVehicleGroup, v_txtVehicleCode, v_txtJobCode
	Dim v_txtDateFrom, v_txtDateTo, NO, cClass
	
	v_txtDateFrom		=	trim(request.Form("txtDateFrom"))
	v_txtDateTo			=	trim(request.Form("txtDateTo"))
	v_txtVehicleGroup	=	trim(request.Form("txtVehicleGroup"))
	v_txtVehicleCode	=	trim(request.Form("txtVehicleCode"))
	v_txtJobCode		=	trim(request.Form("txtJobCode"))
	
	
	strDailyVehicle = 	" SELECT" & _
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
						" SUM (nvl(a.unit,0)) unit," & _
						" SUM (nvl(a.hours,0)) hours," & _
						" SUM (nvl(a.output,0)) output" & _
						" FROM" & _
						" vehicleactivity a," & _
						" vehicle b," & _
						" vehiclegroup c," & _
						" job d" & _
						" WHERE" & _
						" vehicledate BETWEEN TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
						" AND TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
						" AND a.fcba = '" & Session("FCBA") & "'" & _
						" AND a.jobcode = d.fccode" & _
						" AND a.vehiclecode = b.fccode" & _
						" AND b.vehiclegroupcode = c.vehiclegroupcode" & _
						" AND a.fcba = b.fcba" & _
						" AND a.fcba = d.fcba" & _
						" AND b.fcba = c.fcba"
	if len(trim(v_txtVehicleGroup)) > 0 then
		strDailyVehicle = strDailyVehicle & " AND b.vehiclegroupcode = '" & v_txtVehicleGroup & "'"
	end if
	if len(trim(v_txtVehicleCode)) > 0 then
		strDailyVehicle = strDailyVehicle & " AND a.vehiclecode = '" & v_txtVehicleCode & "'"
	end if
	if len(trim(v_txtJobCode)) > 0 then
		strDailyVehicle = strDailyVehicle & " AND a.jobcode = '" & v_txtJobCode & "'"
	end if
	strDailyVehicle = strDailyVehicle &	" GROUP BY" & _
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
		<h5><%=Session("BA")%></h5>
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
		<th>VEHICLE GROUP</th>
		<th>VEHICLE CODE</th>
		<th>JOB CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_DailyVehicle_BA_NAF_M.asp">

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
				<td>
				<select name="txtVehicleGroup">	
				<option value="">--ALL--</option>
		<%
			Dim rsVehicleGroup
			Set rsVehicleGroup = Server.CreateObject("ADODB.Recordset")
			rsVehicleGroup.CursorLocation = adUseClient
			strsql =  " SELECT VehicleGroupCode, Description FROM VehicleGroup WHERE FCBA = '" & Session("fcBA") & "' ORDER BY Description ASC"
			rsVehicleGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsVehicleGroup.eof then
				DO while not rsVehicleGroup.eof
				if v_txtVehicleGroup = rsVehicleGroup("VehicleGroupCode") then
		%>			<option value="<%=rsVehicleGroup("VehicleGroupCode")%>" Selected><%=rsVehicleGroup("VehicleGroupCode")%> - <%=rsVehicleGroup("Description")%></option>
				<%else%>
					<option value="<%=rsVehicleGroup("VehicleGroupCode")%>"><%=rsVehicleGroup("VehicleGroupCode")%> - <%=rsVehicleGroup("Description")%></option>
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
			strsql =  " SELECT FCCODE, FCNAME FROM Vehicle WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCNAME ASC"
			rsVehicleCode.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsVehicleCode.eof then
				DO while not rsVehicleCode.eof
				if v_txtVehicleCode = rsVehicleCode("FCCODE") then
		%>			<option value="<%=rsVehicleCode("FCCODE")%>" Selected><%=rsVehicleCode("FCCODE")%> - <%=rsVehicleCode("FCNAME")%></option>
				<%else%>
					<option value="<%=rsVehicleCode("FCCODE")%>"><%=rsVehicleCode("FCCODE")%> - <%=rsVehicleCode("FCNAME")%></option>
				<%end if%>
		<%			rsVehicleCode.movenext
				loop
			End if
			rsVehicleCode.Close
			set rsVehicleCode = nothing
		%>
		</select>		
		</td>
				<td>
				<select name="txtJobCode">	
				<option value="">--ALL--</option>
		<%
			Dim rsJobCode
			Set rsJobCode = Server.CreateObject("ADODB.Recordset")
			rsJobCode.CursorLocation = adUseClient
			strsql =  " SELECT FCCODE, FCNAME FROM Job WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCNAME ASC"
			rsJobCode.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsJobCode.eof then
				DO while not rsJobCode.eof
				if v_txtJobCode = rsJobCode("FCCODE") then
		%>			<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>
				<%else%>
					<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>
				<%end if%>
		<%			rsJobCode.movenext
				loop
			End if
			rsJobCode.Close
			set rsJobCode = nothing
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
<th>GROUP KENDARAAN</th>
<th>NAMA KENDARAAN</th>
<th>KODE PEKERJAAN</th>
<th>JENIS PEKERJAAN</th>
<th>TARGET TYPE</th>
<th>TARGET CODE</th>
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
				<td align="left"><%=rsDailyVehicle("vehiclegroupname")%></td>
				<td align="left"><%=rsDailyVehicle("vehiclename")%></td>
				<td align="left"><%=rsDailyVehicle("jobcode")%></td>
				<td align="left"><%=rsDailyVehicle("jobname")%></td>
				<td align="left"><%=rsDailyVehicle("locationtype")%></td>
				<td align="left"><%=rsDailyVehicle("locationcode")%></td>
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
				<td align="center" colspan="7">TOTAL</td>
				<td align="right"><%=FormatNumber(nTotal_0,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>
				
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->

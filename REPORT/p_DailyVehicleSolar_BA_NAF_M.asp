<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsDailyVehicleSolar, strDailyVehicleSolar, strGroup, strOrder, strsql, v_txtVehicleGroup, v_txtVehicleCode
	Dim v_txtDateFrom, v_txtDateTo, NO, cClass
	
	v_txtDateFrom		=	trim(request.Form("txtDateFrom"))
	v_txtDateTo			=	trim(request.Form("txtDateTo"))
	v_txtVehicleGroup	=	trim(request.Form("txtVehicleGroup"))
	v_txtVehicleCode	=	trim(request.Form("txtVehicleCode"))
	
	
	strDailyVehicleSolar = 	"  SELECT" & _
							"  C.VEHICLEGROUPCODE VEHICLEGROUPCODE," & _
							"  C.DESCRIPTION DESCRIPTION," & _
							"  A.VEHICLECODE VEHICLECODE," & _
							"  B.FCNAME VEHICLENAME," & _
							"  A.UNIT UNIT," & _
							"  A.HOURS HOURS," & _
							"  A.SOLAR SOLAR," & _
							"  ROUND (DECODE (A.UNIT, 0, 0, A.SOLAR / A.UNIT), 2) RASIO" & _
							"  FROM VIEW_VEHICLEPRODUCTIVITY A, VEHICLE B, VEHICLEGROUP C" & _
							"  WHERE A.FCBA = '" &  Session("FCBA") & "'" & _
						    "  AND A.ACTIONDATE BETWEEN TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
						    "  AND TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
							"  AND A.VEHICLECODE = B.FCCODE" & _
							"  AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE" & _
							"  AND A.FCBA = B.FCBA" & _
							"  AND B.FCBA = C.FCBA"

	if len(trim(v_txtVehicleGroup)) > 0 then
		strDailyVehicleSolar = strDailyVehicleSolar & " AND B.VEHICLEGROUPCODE = '" & v_txtVehicleGroup & "'"
	end if
	if len(trim(v_txtVehicleCode)) > 0 then
		strDailyVehicleSolar = strDailyVehicleSolar & " AND A.VEHICLECODE = '" & v_txtVehicleCode & "'"
	end if
	
	Set rsDailyVehicleSolar = Server.CreateObject("ADODB.Recordset")
	'response.write strDailyVehicleSolar
	rsDailyVehicleSolar.Open strDailyVehicleSolar, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Daily Vehicle Solar Report</title>
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
		<h2>LAPORAN PEMAKAIAN SOLAR KENDARAAN</h3>
		<h5>TANGGAL: <%=v_txtDateFrom%> s/d <%=v_txtDateTo%></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>DARI TANGGAL</th>
		<th>SAMPAI TANGGAL</th>
		<th>VEHICLE GROUP</th>
		<th>VEHICLE CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_DailyVehicleSolar_BA_NAF_M.asp">
		
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
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">

	<tr>
<th>NO</th>
<th>GROUP KENDARAAN</th>
<th>KODE KENDARAAN</th>
<th>NAMA KENDARAAN</th>
<th>JUMLAH KM/ HM</th>
<th>JUMLAH JAM KERJA</th>
<th>JUMLAH SOLAR</th>
<th>RASIO</th>
	</tr>

<%
Dim nTotal_0
Dim nTotal_1
Dim nTotal_2, nTotal_3
	if not rsDailyVehicleSolar.EOF then
		nTotal_0 = 0
		nTotal_1 = 0
		nTotal_2 = 0
		nTotal_3 = 0
		NO = 1
		
		while not rsDailyVehicleSolar.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="left"><%=rsDailyVehicleSolar("DESCRIPTION")%></td>
				<td align="left"><%=rsDailyVehicleSolar("VEHICLECODE")%></td>
				<td align="left"><%=rsDailyVehicleSolar("VEHICLENAME")%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT"),2)%></%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("HOURS"),2)%></%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR"),2)%></%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("RASIO"),2)%></%></td>
			</tr>
<%
			NO = NO + 1
			nTotal_0 = nTotal_0 + Cdbl(rsDailyVehicleSolar("UNIT"))
			nTotal_1 = nTotal_1 + Cdbl(rsDailyVehicleSolar("HOURS"))
			nTotal_2 = nTotal_2 + Cdbl(rsDailyVehicleSolar("SOLAR"))
			nTotal_3 = nTotal_3 + Cdbl(rsDailyVehicleSolar("RASIO"))
			
			rsDailyVehicleSolar.MoveNext
		Wend
	end if
	rsDailyVehicleSolar.Close
	set rsDailyVehicleSolar = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="4">TOTAL</td>
				<td align="right"><%=FormatNumber(nTotal_0,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_3,2)%></td>
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->

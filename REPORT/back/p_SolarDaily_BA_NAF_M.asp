<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim txtDate, txtFcba
	Dim v_txtDate, v_txtFcba
	Dim rsAttDaily, strAttDaily, strGroup, strOrder, strSolarDaily, strMain
	Dim NO, cClass, nSpanTotal
	
	txtDate				=	trim(request("txtDate"))
	'txtFcba 		    =   trim(request("txtFcba"))

	v_txtDate			=	txtDate
	'v_txtFcba 		    =   txtFcba
	
	nSpanTotal = 4
	
	strMain = 	" SELECT FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" SUM (QUANTITY_0) QUANTITY_0," & _
				" SUM (QUANTITY_1) QUANTITY_1," & _
				" SUM (QUANTITY_2) QUANTITY_2" & _
				" FROM (SELECT FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" QUANTITY QUANTITY_0," & _
				" 0 QUANTITY_1," & _
				" 0 QUANTITY_2" & _
				" FROM VIEW_SOLARUSED_SUMMARY" & _
				" WHERE SIVDATE = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
				" UNION" & _
				" SELECT FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" 0 QUANTITY_0," & _
				" SUM (QUANTITY) QUANTITY_1," & _
				" 0 QUANTITY_2" & _
				" FROM VIEW_SOLARUSED_SUMMARY" & _
				" WHERE SIVDATE >=" & _
				" TO_DATE (" & _
				" TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')," & _
				" 'MM')" & _
				" || '/01/'" & _
				" || TO_CHAR (" & _
				" TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')," & _
				" 'YYYY')," & _
				" 'MM/DD/YYYY')" & _
				" AND SIVDATE <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
				" GROUP BY FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME" & _
				" UNION" & _
				" SELECT FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" 0 QUANTITY_0," & _
				" 0 QUANTITY_1," & _
				" SUM (QUANTITY) QUANTITY_2" & _
				" FROM VIEW_SOLARUSED_SUMMARY" & _
				" WHERE SIVDATE <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
				" GROUP BY FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME)" & _
				" GROUP BY FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME"

	strSolarDaily	=	" SELECT FCBA, LOCATIONTYPE, DESCRIPTION, LOCATIONCODE, FCNAME"

	strSolarDaily	=	strSolarDaily & ", SUM (QUANTITY_0) QUANTITY_0, SUM (QUANTITY_1) QUANTITY_1, SUM (QUANTITY_2) QUANTITY_2"
	strSolarDaily 	= 	strSolarDaily &  " FROM ( " & strMain & " )  "
	
	strSolarDaily = strSolarDaily & " WHERE FCBA = '" & Session("fcba") & "'"

	strSolarDaily	=	strSolarDaily & " GROUP BY FCBA, LOCATIONTYPE, DESCRIPTION, LOCATIONCODE, FCNAME"
	
	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strSolarDaily, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Daily Solar Report</title>
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
<form name="frmEntry" method="post" action="p_SolarDaily_BA_NAF_M.asp">
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
		<h5><%=Session("BA") %></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN PEMAKAIAN SOLAR</h3>
		<h5>TANGGAL: <%=v_txtDate%></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>	
		<th>&nbsp;</th>
	</tr>
	<tr>
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
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>

<table border="0" width="100%" id="product-table">

	<tr>
		<th rowspan="2">NO</th>
		<th rowspan="2">HEADER PEMAKAI</th>
		<th rowspan="2">GROUP PEMAKAI</th>
		<th rowspan="2">KODE PEMAKAI</th>
		<th rowspan="2">NAMA PEMAKAI</th>	
		
		
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th>HARI INI <%=v_txtDate%></th>
		<% else %>
		<th>HARI INI</th>
		<% end if%>
		
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th>SAMPAI HARI INI <%=v_txtDate%></th>
		<% else %>
		<th>SAMPAI HARI INI</th>
		<% end if%>
		<th>SAMPAI DENGAN BULAN INI</th>
	</tr>
	<tr>
		<th>TOTAL</th>
		<th>TOTAL</th>
		<th>TOTAL</th>
	</tr>

<%
Dim sprevFcba, tempjob_forfieldstatus
Dim nTotal_0, nTotal_1, nTotal_2

	if not rsAttDaily.EOF then
		nTotal_0 = 0
		nTotal_1 = 0
		nTotal_2 = 0
		NO = 1
		
		sprevFcba=""
		
		while not rsAttDaily.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
			<td align="center"><%=NO%></td>
			<td align="center"><%=rsAttDaily("locationtype")%></td>
			<td align="left"><%=rsAttDaily("description")%></td>
			<td align="left"><%=rsAttDaily("locationcode")%></td>
			<td align="left"><%=rsAttDaily("fcname")%></td>
			<td align="right"><%=FormatNumber(rsAttDaily("QUANTITY_0"),2)%></td>
			<td align="right"><%=FormatNumber(rsAttDaily("QUANTITY_1"),2)%></td>
			<td align="right"><%=FormatNumber(rsAttDaily("QUANTITY_2"),2)%></td>
			</tr>
<%
			NO = NO + 1
			nTotal_0 = nTotal_0 + Cdbl(rsAttDaily("QUANTITY_0"))
			nTotal_1 = nTotal_1 + Cdbl(rsAttDaily("QUANTITY_1"))
			nTotal_2 = nTotal_2 + Cdbl(rsAttDaily("QUANTITY_2"))

			
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 1 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nTotal_0,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>				
			</tr>
</table>
</form>

<!--#include file="../Bottom_Report.asp"-->

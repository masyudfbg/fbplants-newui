<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsUseSolar, strUseSolar, strGroup, strOrder, strQuery
	Dim v_txttodaydate, strsql, NO, cClass
	
	v_txttodaydate		=	trim(request.Form("txttodaydate"))
	
	strUseSolar	=	" select" & _
					" fcba," & _
					" locationtype," & _
					" sivdate," & _
					" description," & _
					" locationcode," & _
					" fcname," & _
					" sum(quantity) quantity" & _
					" from (" & _
					" SELECT a.fcba, a.locationtype, a.sivdate, c.description, a.locationcode, b.fcname," & _
					" SUM (a.quantity) quantity" & _
					" FROM view_solarused a, vehicle b, vehiclegroup c" & _
					" WHERE a.locationcode = b.fccode" & _
					" AND a.locationtype = 'VH'" & _
					" AND b.vehiclegroupcode = c.vehiclegroupcode" & _
					" AND a.fcba = b.fcba" & _
					" AND b.fcba = c.fcba" & _
					" GROUP BY a.fcba, a.locationtype, a.sivdate, c.description, a.locationcode, b.fcname" & _
					" UNION" & _
					" SELECT a.fcba, a.locationtype, a.sivdate, b.division description, a.locationcode," & _
					" b.fcname, SUM (a.quantity) quantity" & _
					" FROM view_solarused a, FIELD b" & _
					" WHERE a.locationcode = b.fccode AND a.locationtype = 'FF'" & _
					" AND a.fcba = b.fcba" & _
					" GROUP BY a.fcba, a.locationtype, a.sivdate, b.division, a.locationcode, b.fcname" & _
					" UNION" & _
					" SELECT a.fcba, a.locationtype, a.sivdate, a.locationcode description," & _
					" a.locationcode, a.locationcode fcname, SUM (a.quantity) quantity" & _
					" FROM view_solarused a" & _
					" WHERE a.locationtype IN ('GC', 'LOC')" & _
					" GROUP BY a.fcba," & _
					" a.locationtype," & _
					" a.sivdate," & _
					" a.locationcode," & _
					" a.locationcode," & _
					" a.locationcode" & _
					" UNION" & _
					" SELECT a.fcba, a.locationtype, a.sivdate, b.fcname description, a.locationcode," & _
					" a.locationcode fcname, SUM (a.quantity) quantity" & _
					" FROM view_solarused a, job b" & _
					" WHERE a.jobcode = b.fccode AND a.fcba = b.fcba AND a.locationtype IN" & _
					" ('CB')" & _
					" GROUP BY a.fcba, a.locationtype, a.sivdate, b.fcname, a.locationcode, a.locationcode" & _
					" )" & _
					" where" & _
					" fcba = '" & Session("fcba") & "' and" & _
					" sivdate = to_date('" & v_txttodaydate & "','MM/DD/YYYY')"		
						
	strGroup	=	" group by fcba, locationtype, sivdate, description, locationcode, fcname" 
	strOrder	=	" order by locationtype, sivdate, description, locationcode, fcname"
	
	strQuery	=	strUseSolar + strGroup + strOrder

	Set rsUseSolar = Server.CreateObject("ADODB.Recordset")
	rsUseSolar.Open strQuery, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Pemakaian Solar</title>
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
	<div id="subheader"class="centered">
		<h2>LAPORAN HARIAN PEMAKAIAN SOLAR</h2>
		<h5>TANGGAL: <%=v_txttodaydate%></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_UseSolar.asp">
		<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
				<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
		<% else %>
				<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if %>
			<a href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th>NO</th>
		<th>LOCATION TYPE</th>
		<th>SIV DATE</th>
		<th>DESCRIPTION</th>
		<th>LOCATION CODE</th>
		<th>DESCRIPTION</th>
		<th>QUANTITY</th>
	</tr>
<%
Dim nQuantity

	if not rsUseSolar.EOF then
		NO = 1
		nQuantity	= 0
		while not rsUseSolar.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="left"><%=rsUseSolar(1)%></td>
				<td align="left"><%=rsUseSolar(2)%></td>
				<td align="left"><%=rsUseSolar(3)%></td>
				<td align="left"><%=rsUseSolar(4)%></td>
				<td align="left"><%=rsUseSolar(5)%></td>
				<td align="right"><%=FormatNumber(rsUseSolar(6),2)%></td>
			</tr>
<%
			NO = NO + 1
			nQuantity = nQuantity + Cdbl(rsUseSolar(6))

			rsUseSolar.MoveNext
		Wend
	end if
	rsUseSolar.Close
	set rsUseSolar = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="6">TOTAL</td>
				<td align="right"><%=FormatNumber(nQuantity,2)%></td>
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->

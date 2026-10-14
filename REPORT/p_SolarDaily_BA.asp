<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsAttDaily, strAttDaily, strGroup, strOrder, strSolarDaily,  v_txtFcba
	Dim v_txtDate, NO, cClass
	
	v_txtDate		=	trim(request.Form("txtDate"))
	v_txtFcba 		    =   trim(request.Form("txtFcba"))	
	
	
	
	strSolarDaily = " SELECT   locationtype, description, locationcode, fcname, " & _
							 " SUM (quantity_0) quantity_0, SUM (quantity_1) quantity_1 , SUM (quantity_2) quantity_2 " & _
						" FROM (SELECT locationtype, description, locationcode, fcname, " & _
									 " quantity quantity_0, 0 quantity_1, 0 quantity_2 " & _
								" FROM view_solarused_summary " & _
							   " WHERE fcba = '" & v_txtFcba& "' " & _
								 " AND sivdate = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') " & _
							  " UNION " & _
							   " SELECT locationtype, description, locationcode, fcname, " & _
									 " 0 quantity_0, sum(quantity) quantity_1, 0 quantity_2 " & _
								" FROM view_solarused_summary " & _
							   " WHERE fcba = '" & v_txtFcba& "' " & _
							     " AND sivdate >= " & _
									  " TO_DATE (   TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'MM') " & _
											   " || '/01/' " & _
											   " || TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'), " & _
														   " 'YYYY' " & _
														  " ), " & _
											   " 'MM/DD/YYYY' " & _
											  " ) " & _
								 " AND sivdate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') " & _ 
								 " GROUP BY locationtype, description, locationcode, fcname " & _
							  " UNION " & _
							  " SELECT   locationtype, description, locationcode, fcname, " & _
									   " 0 quantity_0, 0 quantity_1,  SUM (quantity) quantity_2 " & _
								  " FROM view_solarused_summary " & _
								 " WHERE fcba = '" & v_txtFcba& "' " & _
								   " AND sivdate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') " & _
							  " GROUP BY locationtype, description, locationcode, fcname) " & _
					" GROUP BY locationtype, description, locationcode, fcname"


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
	<div id="wrapper">
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN PEMAKAIAN SOLAR PER KEBUN</h3>
		<h5>TANGGAL: <%=v_txtDate%>, KEBUN : <%=v_txtFcba %></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>
	    <th>KEBUN</th>		
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_SolarDaily_BA.asp">
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
		
		<td align="center">
			<select name="txtFcba">	
			<option value="">--NONE--</option>		
			<%
				Dim rsFcba, strsql1
				Set rsFcba = Server.CreateObject("ADODB.Recordset")
				rsFcba.CursorLocation = adUseClient
				strsql1 =  " SELECT fccode FROM businessunit "
				rsFcba.Open strsql1, connStr, adOpenForwardOnly, adLockReadOnly	
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
		
		
		
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">

	<tr>
		<th rowspan="2">HEADER PEMAKAI</th>
		<th rowspan="2">GROUP PEMAKAI</th>
		<th rowspan="2">KODE PEMAKAI</th>
		<th rowspan="2">NAMA PEMAKAI</th>
		
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th colspan="2">HARI INI <%=v_txtDate%></th>
		<% else %>
		<th colspan="2">HARI INI</th>
		<% end if%>
		
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th colspan="2">SAMPAI HARI INI <%=v_txtDate%></th>
		<% else %>
		<th colspan="2">SAMPAI HARI INI</th>
		<% end if%>
		<th colspan="2">SAMPAI DENGAN BULAN INI</th>
	</tr>
	<tr>
		<th colspan="2">TOTAL</th>
		<th colspan="2">TOTAL</th>
		<th colspan="2">TOTAL</th>
	</tr>

<%
Dim nTotal_0
Dim nTotal_1
Dim nTotal_2
	if not rsAttDaily.EOF then
		nTotal_0 = 0
		nTotal_1 = 0
		nTotal_2 = 0
		
		
		while not rsAttDaily.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="left"><%=rsAttDaily(0)%></td>
				<td align="left"><%=rsAttDaily(1)%></td>
				<td align="left"><%=rsAttDaily(2)%></td>
				<td align="left"><%=rsAttDaily(3)%></td>
				
				<td align="right" colspan="2"><%=FormatNumber(rsAttDaily(4),2)%></td>
				<td align="right" colspan="2"><%=FormatNumber(rsAttDaily(5),2)%></td>
				<td align="right" colspan="2"><%=FormatNumber(rsAttDaily(6),2)%></td>
			</tr>
<%
			nTotal_0 = nTotal_0 + Cdbl(rsAttDaily(4))
			nTotal_1 = nTotal_1 + Cdbl(rsAttDaily(5))
			nTotal_2 = nTotal_2 + Cdbl(rsAttDaily(6))

			
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="4">TOTAL</td>
				<td align="right"  colspan="2"><%=FormatNumber(nTotal_0,2)%></td>
				<td align="right"  colspan="2"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"  colspan="2"><%=FormatNumber(nTotal_2,2)%></td>
				
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->
